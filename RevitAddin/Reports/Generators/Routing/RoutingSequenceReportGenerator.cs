//-----------------------------------------------------------------------------
// <copyright file="RoutingSequenceReportGenerator.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright//>
// <summary>
//   Generates the Routing Sequence report using Dijkstra's algorithm for shortest path,
//   with virtual pathing between disconnected containment islands.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Made helper classes (Pathfinder, ContainmentIsland, etc.) public to allow access from diagnostic tools.
// - [APPLIED FIX]: Corrected virtual path length calculation to include the physical length of segments.
// - [APPLIED FIX]: Corrected GetRelevantNetwork to properly expand the containment graph.
// - [APPLIED FIX]: Added detailed status messages for virtual pathfinding failures.
// - [APPLIED FIX]: Removed network expansion logic to ensure graph only uses assigned containment.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Reports.Base;
using RTS.Reports.Generators.Routing;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Threading;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Main class for generating the Routing Sequence Report.
    /// This class orchestrates the process of reading cable data, analyzing Revit model containment,
    /// calculating cable routes, and exporting the results to a CSV file.
    /// </summary>
    public class RoutingSequenceReportGenerator : ReportGeneratorBase
    {
        // GUID for the shared parameter "RTS_ID", used to uniquely identify containment and equipment.
        private readonly Guid _rtsIdGuid = SharedParameters.General.RTS_ID;
        // A list of GUIDs for all shared parameters that can hold a cable reference.
        private readonly List<Guid> _cableGuids = SharedParameters.Cable.AllCableGuids;

        /// <summary>
        /// Initializes a new instance of the <see cref="RoutingSequenceReportGenerator"/> class.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="commandData">The external command data from Revit.</param>
        /// <param name="pcExtensible">An instance of the extensible class for data handling.</param>
        public RoutingSequenceReportGenerator(Document doc, ExternalCommandData commandData, PC_ExtensibleClass pcExtensible)
            : base(doc, commandData, pcExtensible)
        {
        }

        /// <summary>
        /// The main method to generate the report. It handles data loading, validation,
        /// processing, and file output.
        /// </summary>
        public override void GenerateReport()
        {
            try
            {
                // Step 1: Recall primary cable data from the project's extensible storage.
                List<PC_ExtensibleClass.CableData> primaryData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    Document, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName);

                // Validate that cable data exists.
                if (primaryData == null || !primaryData.Any())
                {
                    ShowInfo("No Data Found", "No primary cable data was found in the project's extensible storage.");
                    return;
                }

                // Step 2: Collect all containment elements and check if any cables have been assigned to them.
                var containmentCollector = new ContainmentCollector(Document);
                var allContainmentInProject = containmentCollector.GetAllContainmentElements();
                if (!allContainmentInProject.Any(elem => _cableGuids.Any(guid => elem.get_Parameter(guid)?.HasValue ?? false)))
                {
                    ShowInfo("No Assignments Found", "No cables have been assigned to any containment elements.");
                    return;
                }

                // Step 3: Prompt user for the output file path.
                string filePath = GetOutputFilePath("Routing_Sequence.csv", "Save Routing Sequence Report");
                if (string.IsNullOrEmpty(filePath)) return; // User cancelled the save dialog.

                // Step 4: Prepare the report header and a lookup dictionary for quick cable data access.
                var projectInfo = new FilteredElementCollector(Document).OfCategory(BuiltInCategory.OST_ProjectInformation).WhereElementIsNotElementType().Cast<ProjectInfo>().FirstOrDefault();
                var sb = ReportFormatter.CreateReportHeader(projectInfo, Document);
                var cableLookup = primaryData.Where(c => !string.IsNullOrWhiteSpace(c.CableReference)).GroupBy(c => c.CableReference).ToDictionary(g => g.Key, g => g.First());

                // Step 5: Show a progress window to the user.
                var progressWindow = new RoutingReportProgressBarWindow { Owner = System.Windows.Application.Current?.MainWindow };
                progressWindow.Show();

                try
                {
                    // Pre-collect all necessary Revit elements to avoid repeated queries inside the loop.
                    var allCandidateEquipment = containmentCollector.GetAllElectricalEquipment();
                    var allFittingsInProject = allContainmentInProject.Where(e => !(e is MEPCurve)).ToList();
                    int totalCables = cableLookup.Count;
                    int currentCable = 0;

                    // Step 6: Process each cable to determine its route.
                    foreach (var cableRef in cableLookup.Keys)
                    {
                        currentCable++;
                        var cableInfo = cableLookup[cableRef];

                        // Update the progress window.
                        double percentage = (double)currentCable / totalCables * 100.0;
                        progressWindow.UpdateProgress(currentCable, totalCables, cableRef, cableInfo?.From ?? "N/A", cableInfo?.To ?? "N/A", percentage, "Building local network...");
                        Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background); // Allow UI to update.

                        if (progressWindow.IsCancelled) break;

                        // Process the individual cable route and append the result to the report string builder.
                        var result = ProcessCableRoute(cableRef, cableInfo, allContainmentInProject, allFittingsInProject, allCandidateEquipment);
                        sb.AppendLine(result);
                    }

                    // Step 7: Finalize the report.
                    if (!progressWindow.IsCancelled)
                    {
                        File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                        progressWindow.ShowCompletion(filePath);
                    }
                }
                catch (Exception ex)
                {
                    // Handle errors during the report generation loop.
                    progressWindow.ShowError($"An unexpected error occurred: {ex.Message}");
                    ShowError($"Failed to export Routing Sequence report: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Handle critical errors during setup.
                ShowError($"A critical error occurred: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Processes a single cable to find its route, handling both connected and disconnected (virtual) paths.
        /// /// </summary>
        /// <returns>A CSV-formatted string representing the routing result for the cable.</returns>
        private string ProcessCableRoute(string cableRef, PC_ExtensibleClass.CableData cableInfo, List<Element> allContainmentInProject, List<Element> allFittingsInProject, List<Element> allCandidateEquipment)
        {
            // Initialize result variables.
            string status = "Processing Error";
            string fromStatus = "Not Specified";
            string toStatus = "Not Specified";
            string routingSequence = "Could not determine route.";
            string branchSequence = "N/A";
            double supportedLengthFeet = 0.0;
            double unsupportedLengthFeet = 0.0;
            string assignedContainmentStr = "N/A";
            string graphedContainmentStr = "N/A";
            int islandCount = 0;
            string traySystems = "N/A";
            string containmentRating = "N/A";

            try
            {
                // Step 1: Find the start equipment in the Revit model based on the 'From' field.
                List<Element> startCandidates;
                bool startFound = false;
                if (!string.IsNullOrWhiteSpace(cableInfo.From))
                {
                    startCandidates = FindMatchingEquipment(cableInfo.From, allCandidateEquipment, _rtsIdGuid, Document);
                    if (startCandidates.Any())
                    {
                        string matchedName = startCandidates.First().get_Parameter(_rtsIdGuid)?.AsString() ?? startCandidates.First().get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "Unknown";
                        fromStatus = $"Found: {matchedName}";
                        startFound = true;
                    }
                    else
                    {
                        fromStatus = "Not Found";
                    }
                }
                else
                {
                    startCandidates = new List<Element>();
                }

                // Step 2: Find the end equipment in the Revit model based on the 'To' field.
                List<Element> endCandidates;
                bool endFound = false;
                if (!string.IsNullOrWhiteSpace(cableInfo.To))
                {
                    endCandidates = FindMatchingEquipment(cableInfo.To, allCandidateEquipment, _rtsIdGuid, Document);
                    if (endCandidates.Any())
                    {
                        string matchedName = endCandidates.First().get_Parameter(_rtsIdGuid)?.AsString() ?? endCandidates.First().get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "Unknown";
                        toStatus = $"Found: {matchedName}";
                        endFound = true;
                    }
                    else
                    {
                        toStatus = "Not Found";
                    }
                }
                else
                {
                    endCandidates = new List<Element>();
                }

                // Determine if a high-accuracy node-based search should be used.
                bool useHighAccuracy = false;
                if (startFound && endFound)
                {
                    var startPoint = (startCandidates.First().Location as LocationPoint)?.Point;
                    var endPoint = (endCandidates.First().Location as LocationPoint)?.Point;

                    if (startPoint != null && endPoint != null)
                    {
                        const double feetToMeters = 0.3048;
                        double distanceMeters = startPoint.DistanceTo(endPoint) * feetToMeters;
                        if (distanceMeters < 15)
                        {
                            useHighAccuracy = true;
                        }
                    }
                }

                // Step 3: Find all containment elements assigned to this specific cable.
                var cableContainmentElements = allContainmentInProject
                    .Where(elem => _cableGuids.Any(guid => string.Equals(elem.get_Parameter(guid)?.AsString(), cableRef, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                if (!cableContainmentElements.Any())
                {
                    status = "No Containment Assigned";
                    routingSequence = "This cable reference was not found on any containment elements.";
                }
                else
                {
                    // Step 4: Build a graph of the containment network.
                    // The graph is built *only* from the elements explicitly assigned to this cable.
                    // This prevents the pathfinder from using containment that is not allocated for this route.
                    var containmentGraph = new ContainmentGraph(cableContainmentElements, Document);
                    var elementIdToLengthMap = containmentGraph.GetElementLengthMap();
                    var rtsIdToElementMap = ReportFormatter.BuildRtsIdMap(cableContainmentElements, _rtsIdGuid);

                    // Format the list of assigned and graphed containment for the diagnostic columns.
                    assignedContainmentStr = ReportFormatter.FormatRtsIdList(cableContainmentElements, _rtsIdGuid, elementIdToLengthMap);
                    graphedContainmentStr = ReportFormatter.FormatRtsIdList(cableContainmentElements, _rtsIdGuid, elementIdToLengthMap);
                    traySystems = ReportFormatter.FormatTraySystems(cableContainmentElements, _rtsIdGuid);
                    containmentRating = ReportFormatter.FormatContainmentRatings(cableContainmentElements, _rtsIdGuid);

                    List<ElementId> bestPath = null;

                    // Step 5: Determine the routing strategy based on island count.
                    var islands = Pathfinder.GroupIntoIslands(cableContainmentElements, containmentGraph, Document);
                    islandCount = islands.Count;

                    bool forceVirtualPath = islandCount > 1 || !startFound || !endFound;

                    if (!forceVirtualPath)
                    {
                        // Attempt to find a direct, confirmed path as it's a single network.
                        var pathResult = Pathfinder.FindConfirmedPath(startCandidates, endCandidates, cableContainmentElements, containmentGraph, elementIdToLengthMap, Document, useHighAccuracy);
                        if (pathResult != null && pathResult.Any())
                        {
                            status = "Route Confirmed";
                            bestPath = pathResult;
                            supportedLengthFeet = bestPath.Sum(id => elementIdToLengthMap.ContainsKey(id) ? elementIdToLengthMap[id] : 0.0);
                        }
                    }

                    // Step 6: If no confirmed path was found (or if virtual path was forced), use island hopping.
                    if (bestPath == null)
                    {
                        var virtualResult = Pathfinder.FindBestDisconnectedSequence(startCandidates, endCandidates, cableContainmentElements, containmentGraph, rtsIdToElementMap, elementIdToLengthMap, Document, useHighAccuracy);

                        if (!string.IsNullOrEmpty(virtualResult.StatusMessage))
                        {
                            status = "Virtual Path Error";
                            routingSequence = virtualResult.StatusMessage;
                        }
                        else
                        {
                            status = (startFound && endFound) ? "Route Unconfirmed (Virtual Path)" : "Route Incomplete (Start/End Not Found)";
                            routingSequence = virtualResult.RoutingSequence;
                        }

                        bestPath = virtualResult.StitchedPath;
                        supportedLengthFeet = bestPath?.Sum(id => elementIdToLengthMap.ContainsKey(id) ? elementIdToLengthMap[id] : 0.0) ?? 0.0;
                        unsupportedLengthFeet = virtualResult.VirtualLength * 1.20; // Apply 20% contingency
                        islandCount = virtualResult.IslandCount > 0 ? virtualResult.IslandCount : islandCount;
                    }


                    // Step 7: Format the final path and branch sequence strings.
                    if (bestPath != null && bestPath.Any())
                    {
                        if (status == "Route Confirmed") // Only re-format if it wasn't already formatted by the virtual pathfinder.
                        {
                            string fromPanel = ReportFormatter.FormatEquipment(startCandidates.FirstOrDefault(), Document, _rtsIdGuid);
                            string toPanel = ReportFormatter.FormatEquipment(endCandidates.FirstOrDefault(), Document, _rtsIdGuid);
                            string pathSequence = ReportFormatter.FormatPath(bestPath, rtsIdToElementMap, Document, _rtsIdGuid);
                            routingSequence = $"{fromPanel} >> {pathSequence} >> {toPanel}";
                        }
                        branchSequence = ReportFormatter.FormatBranchSequence(bestPath, rtsIdToElementMap, _rtsIdGuid);
                    }
                }
            }
            catch (Exception ex)
            {
                status = "Processing Error";
                routingSequence = $"Processing failed: {ex.Message}";
                branchSequence = "Error";
            }

            // Step 8: Format the final CSV line for this cable.
            double totalLengthFeet = supportedLengthFeet + unsupportedLengthFeet;
            double totalLengthMeters = totalLengthFeet * 0.3048;
            double supportedLengthMeters = supportedLengthFeet * 0.3048;
            double unsupportedLengthMeters = unsupportedLengthFeet * 0.3048;

            return $"\"{cableRef}\",\"{cableInfo.From ?? "N/A"}\",\"{cableInfo.To ?? "N/A"}\",\"{fromStatus}\",\"{toStatus}\",\"{status}\",\"{totalLengthMeters:F2}\",\"{supportedLengthMeters:F2}\",\"{unsupportedLengthMeters:F2}\",\"{branchSequence}\",\"{routingSequence}\",\"{assignedContainmentStr}\",\"{graphedContainmentStr}\",\"{islandCount}\",\"{traySystems}\",\"{containmentRating}\"";
        }

        /// <summary>
        /// Finds electrical equipment elements that match a given name.
        /// It first attempts to match against the "RTS_ID" shared parameter, then falls back
        /// to the built-in panel name parameter.
        /// </summary>
        /// <returns>A list of matching equipment elements.</returns>
        private static List<Element> FindMatchingEquipment(string nameToMatch, List<Element> allCandidateEquipment, Guid rtsIdGuid, Document doc)
        {
            if (string.IsNullOrWhiteSpace(nameToMatch)) return new List<Element>();

            string cleanNameToMatch = nameToMatch.Trim();

            // First, try to find a match using the RTS_ID. This is the preferred method.
            var rtsIdMatches = allCandidateEquipment.Where(e =>
            {
                var rtsIdVal = e?.get_Parameter(rtsIdGuid)?.AsString();
                if (string.IsNullOrWhiteSpace(rtsIdVal)) return false;

                string cleanRtsIdVal = rtsIdVal.Trim();
                // Use IndexOf for partial matching (e.g., "PanelA" matches "PanelA-1").
                return cleanRtsIdVal.IndexOf(cleanNameToMatch, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       cleanNameToMatch.IndexOf(cleanRtsIdVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            if (rtsIdMatches.Any()) return rtsIdMatches;

            // If no RTS_ID match, fall back to the built-in electrical panel name.
            return allCandidateEquipment.Where(e =>
            {
                var panelNameVal = e?.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString();
                if (string.IsNullOrWhiteSpace(panelNameVal)) return false;

                string cleanPanelNameVal = panelNameVal.Trim();
                return cleanPanelNameVal.IndexOf(cleanNameToMatch, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       cleanNameToMatch.IndexOf(cleanPanelNameVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();
        }
    }
}