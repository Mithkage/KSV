//-----------------------------------------------------------------------------
// <copyright file="RoutingSequenceReportGenerator.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Generates the Routing Sequence report
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible;
using RTS.Reports.Base;
using RTS.Reports.Utils;
using RTS.UI;
using RTS.Utilities; // Add reference to the utilities namespace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Threading;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Generator for the Routing Sequence Report
    /// </summary>
    public class RoutingSequenceReportGenerator : ReportGeneratorBase
    {
        // Replace hardcoded GUIDs with references to SharedParameters
        private readonly Guid _rtsIdGuid = SharedParameters.Cable.RTS_ID_GUID;
        private readonly List<Guid> _cableGuids = SharedParameters.Cable.RTS_Cable_GUIDs;

        public RoutingSequenceReportGenerator(Document doc, ExternalCommandData commandData, PC_ExtensibleClass pcExtensible) 
            : base(doc, commandData, pcExtensible)
        {
        }

        public override void GenerateReport()
        {
            string filePath = GetOutputFilePath("Routing_Sequence.csv", "Save Routing Sequence Report");
            if (string.IsNullOrEmpty(filePath))
            {
                ShowInfo("Export Cancelled", "Output file not selected. Routing Sequence export cancelled.");
                return;
            }

            var projectInfo = new FilteredElementCollector(Document)
                .OfCategory(BuiltInCategory.OST_ProjectInformation)
                .WhereElementIsNotElementType()
                .Cast<ProjectInfo>()
                .FirstOrDefault();

            string projectName = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NAME)?.AsString() ?? "N/A";
            string projectNumber = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NUMBER)?.AsString() ?? "N/A";

            var sb = new StringBuilder();
            sb.AppendLine("Cable Routing Sequence");
            sb.AppendLine($"Project:,{projectName}");
            sb.AppendLine($"Project No:,{projectNumber}");
            sb.AppendLine($"Date:,{DateTime.Now:dd/MM/yyyy}");
            sb.AppendLine();
            sb.AppendLine("Separator Legend:,,,,Tray/Conduit ID Legend:");
            sb.AppendLine(",, (Comma) = Direct connection between same containment types,,,,Format: [Prefix]-[TypeSuffix]-[Level]-[ElevMM]-[BranchNum]");
            sb.AppendLine(",, + (Plus) = Transition between different containment types,,,,Example: LA-FLS-L01-4285-0001");
            sb.AppendLine(",, || (Double Pipe) = Jump between non-connected containment of the same type,,,,Prefix: Service Type (LA=LV Sub-A, LB=LV Sub-B, GA=Gen-A)");
            sb.AppendLine(",, >> (Double Arrow) = Separates disconnected routing segments,,,,TypeSuffix: FLS=Fire-rated, ESS=Essential, DFT=Default");
            sb.AppendLine();
            sb.AppendLine("Cable Reference,From,To,Status,Branch Sequencing,Routing Sequence");

            List<PC_ExtensibleClass.CableData> primaryData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                Document, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName);

            if (primaryData == null || !primaryData.Any())
            {
                ShowInfo("No Data", "No primary cable data found.");
                return;
            }

            var cableLookup = primaryData
                .Where(c => !string.IsNullOrWhiteSpace(c.CableReference))
                .GroupBy(c => c.CableReference)
                .ToDictionary(g => g.Key, g => g.First());

            var allConduits = new FilteredElementCollector(Document)
                .OfCategory(BuiltInCategory.OST_Conduit)
                .WhereElementIsNotElementType()
                .Where(e => e.get_Parameter(_rtsIdGuid)?.HasValue ?? false)
                .ToList();

            var allCableTrays = new FilteredElementCollector(Document)
                .OfCategory(BuiltInCategory.OST_CableTray)
                .WhereElementIsNotElementType()
                .Where(e => e.get_Parameter(_rtsIdGuid)?.HasValue ?? false)
                .ToList();

            var allContainmentElements = allConduits.Concat(allCableTrays).ToList();

            var allEquipment = new FilteredElementCollector(Document)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .WhereElementIsNotElementType().ToList();

            var allFixtures = new FilteredElementCollector(Document)
                .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
                .WhereElementIsNotElementType().ToList();

            var allCandidateEquipment = allEquipment.Concat(allFixtures).ToList();

            var progressWindow = new RoutingReportProgressBarWindow
            {
                Owner = System.Windows.Application.Current?.MainWindow
            };
            progressWindow.Show();

            try
            {
                int totalCables = cableLookup.Count;
                int currentCable = 0;

                foreach (var cableRef in cableLookup.Keys)
                {
                    currentCable++;
                    var cableInfo = cableLookup[cableRef];

                    progressWindow.UpdateProgress(currentCable, totalCables, cableRef,
                        cableInfo?.From ?? "N/A", cableInfo?.To ?? "N/A",
                        (double)currentCable / totalCables * 100.0, "Building routing sequence...");
                    Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background);

                    if (progressWindow.IsCancelled)
                    {
                        progressWindow.TaskDescriptionText.Text = "Operation cancelled by user.";
                        break;
                    }

                    string status = "Error";
                    string routingSequence = "Could not determine route.";
                    string branchSequence = "N/A";
                    List<ElementId> bestPath = null;

                    try
                    {
                        var cableContainmentElements = new List<Element>();
                        foreach (Element elem in allContainmentElements)
                        {
                            foreach (var guid in _cableGuids)
                            {
                                Parameter cableParam = elem.get_Parameter(guid);
                                string paramValue = cableParam?.AsString();
                                if (string.Equals(ReportHelpers.CleanCableReference(paramValue), cableRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    cableContainmentElements.Add(elem);
                                    break;
                                }
                            }
                        }

                        if (!cableContainmentElements.Any())
                        {
                            status = "Pending";
                            routingSequence = "No cable containment assigned in model.";
                        }
                        else
                        {
                            var rtsIdToElementMap = new Dictionary<string, Element>();
                            foreach (var elem in cableContainmentElements)
                            {
                                string rtsId = elem.get_Parameter(_rtsIdGuid)?.AsString();
                                if (!string.IsNullOrWhiteSpace(rtsId) && !rtsIdToElementMap.ContainsKey(rtsId))
                                {
                                    rtsIdToElementMap.Add(rtsId, elem);
                                }
                            }

                            var adjacencyGraph = BuildAdjacencyGraph(cableContainmentElements);

                            List<Element> startCandidates = FindMatchingEquipment(cableInfo.From, allCandidateEquipment);
                            List<Element> endCandidates = FindMatchingEquipment(cableInfo.To, allCandidateEquipment);

                            if (startCandidates.Any() && endCandidates.Any())
                            {
                                double maxLength = -1.0;

                                foreach (var startCandidate in startCandidates)
                                {
                                    foreach (var endCandidate in endCandidates)
                                    {
                                        var currentPath = FindConfirmedPath(startCandidate, endCandidate, cableContainmentElements, adjacencyGraph);
                                        if (currentPath != null && currentPath.Any())
                                        {
                                            double currentLength = GetPathLength(currentPath);
                                            if (currentLength > maxLength)
                                            {
                                                maxLength = currentLength;
                                                bestPath = currentPath;
                                            }
                                        }
                                    }
                                }

                                if (bestPath != null)
                                {
                                    status = "Confirmed";
                                    routingSequence = FormatPath(bestPath, adjacencyGraph, rtsIdToElementMap);
                                }
                                else
                                {
                                    status = "Unconfirmed";
                                    var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                    routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                                }
                            }
                            else if (startCandidates.Any())
                            {
                                status = "Confirmed (From)";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }
                            else if (endCandidates.Any())
                            {
                                status = "Confirmed (To)";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }
                            else
                            {
                                status = "Unconfirmed";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }

                            if (bestPath != null && bestPath.Any())
                            {
                                branchSequence = FormatBranchSequence(bestPath, rtsIdToElementMap);
                            }
                            else if (status.Contains("Unconfirmed") || status.Contains("Confirmed (From)") || status.Contains("Confirmed (To)"))
                            {
                                // For partially confirmed or unconfirmed routes, compile from all segments
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                var allPathIds = disconnectedPaths.SelectMany(p => p).ToList();
                                branchSequence = FormatBranchSequence(allPathIds, rtsIdToElementMap);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        status = "Error";
                        routingSequence = $"Processing failed: {ex.Message}";
                        branchSequence = "Error";
                    }

                    sb.AppendLine($"\"{cableRef}\",\"{cableInfo.From ?? "N/A"}\",\"{cableInfo.To ?? "N/A"}\",\"{status}\",\"{branchSequence}\",\"{routingSequence}\"");
                }

                if (!progressWindow.IsCancelled)
                {
                    System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                    progressWindow.UpdateProgress(totalCables, totalCables, "Complete", "N/A", "N/A", 100, "Exporting file...");
                    progressWindow.TaskDescriptionText.Text = "Export complete.";
                    ShowSuccess("Export Complete", $"Routing Sequence report successfully exported to:\n{filePath}");
                }
            }
            catch (Exception ex)
            {
                progressWindow.ShowError($"An unexpected error occurred: {ex.Message}");
                ShowError($"Failed to export Routing Sequence report: {ex.Message}");
            }
            finally
            {
                progressWindow.Close();
            }
        }

        #region Routing Sequence Helper Methods

        /// <summary>
        /// Builds an adjacency graph using ElementId as keys for stability.
        /// </summary>
        private Dictionary<ElementId, List<ElementId>> BuildAdjacencyGraph(List<Element> elements)
        {
            var graph = new Dictionary<ElementId, List<ElementId>>();
            foreach (var elemA in elements)
            {
                if (!graph.ContainsKey(elemA.Id))
                {
                    graph[elemA.Id] = new List<ElementId>();
                }
                foreach (var elemB in elements)
                {
                    if (elemA.Id == elemB.Id) continue;
                    if (AreConnected(elemA, elemB))
                    {
                        graph[elemA.Id].Add(elemB.Id);
                    }
                }
            }
            return graph;
        }

        /// <summary>
        /// Finds a single path between two nodes using Breadth-First Search on ElementIds.
        /// </summary>
        private List<ElementId> FindSinglePath_BFS(ElementId startNodeId, ElementId endNodeId, Dictionary<ElementId, List<ElementId>> graph)
        {
            if (startNodeId == null || endNodeId == null || startNodeId == endNodeId)
            {
                return new List<ElementId> { startNodeId ?? endNodeId };
            }

            var queue = new Queue<ElementId>();
            var cameFrom = new Dictionary<ElementId, ElementId>();
            var visited = new HashSet<ElementId>();

            queue.Enqueue(startNodeId);
            visited.Add(startNodeId);

            while (queue.Count > 0)
            {
                var currentId = queue.Dequeue();

                if (currentId == endNodeId)
                {
                    var path = new List<ElementId>();
                    var atId = currentId;
                    while (atId != null && atId != ElementId.InvalidElementId)
                    {
                        path.Add(atId);
                        if (!cameFrom.TryGetValue(atId, out ElementId prevId)) break;
                        atId = prevId;
                    }
                    path.Reverse();
                    return path;
                }

                if (graph.TryGetValue(currentId, out List<ElementId> neighborIds))
                {
                    foreach (var neighborId in neighborIds)
                    {
                        if (!visited.Contains(neighborId))
                        {
                            visited.Add(neighborId);
                            cameFrom[neighborId] = currentId;
                            queue.Enqueue(neighborId);
                        }
                    }
                }
            }
            return null; // Path not found
        }

        /// <summary>
        /// Finds the path between two confirmed equipment endpoints.
        /// </summary>
        private List<ElementId> FindConfirmedPath(Element matchedStart, Element matchedEnd, List<Element> cableElements, Dictionary<ElementId, List<ElementId>> graph)
        {
            XYZ matchedStartPt = ReportHelpers.GetElementLocation(matchedStart);
            XYZ matchedEndPt = ReportHelpers.GetElementLocation(matchedEnd);
            if (matchedStartPt == null || matchedEndPt == null) return new List<ElementId>();

            Element startElem = cableElements
                .OrderBy(e => ReportHelpers.GetElementLocation(e)?.DistanceTo(matchedStartPt) ?? double.MaxValue)
                .FirstOrDefault();

            Element endElem = cableElements
                .OrderBy(e => ReportHelpers.GetElementLocation(e)?.DistanceTo(matchedEndPt) ?? double.MaxValue)
                .FirstOrDefault();

            if (startElem == null || endElem == null) return new List<ElementId>();

            return FindSinglePath_BFS(startElem.Id, endElem.Id, graph) ?? new List<ElementId>();
        }

        /// <summary>
        /// Finds all disconnected path segments in the graph.
        /// </summary>
        private List<List<ElementId>> FindDisconnectedPaths(List<Element> cableElements, Dictionary<ElementId, List<ElementId>> graph)
        {
            var allPaths = new List<List<ElementId>>();
            var visited = new HashSet<ElementId>();
            var elementIds = cableElements.Select(e => e.Id).ToList();

            foreach (var startNodeId in elementIds)
            {
                if (visited.Contains(startNodeId)) continue;

                var pathSegment = new List<ElementId>();
                var q = new Queue<ElementId>();

                q.Enqueue(startNodeId);
                visited.Add(startNodeId);
                pathSegment.Add(startNodeId);

                while (q.Count > 0)
                {
                    var currentId = q.Dequeue();
                    if (graph.TryGetValue(currentId, out List<ElementId> neighborIds))
                    {
                        foreach (var neighborId in neighborIds)
                        {
                            if (!visited.Contains(neighborId))
                            {
                                visited.Add(neighborId);
                                pathSegment.Add(neighborId);
                                q.Enqueue(neighborId);
                            }
                        }
                    }
                }
                allPaths.Add(pathSegment);
            }
            return allPaths;
        }

        /// <summary>
        /// Formats a list of ElementIds into the final report string.
        /// </summary>
        private string FormatPath(List<ElementId> path, Dictionary<ElementId, List<ElementId>> graph, Dictionary<string, Element> rtsIdToElementMap)
        {
            if (path == null || !path.Any()) return "Path not found";

            var elementIdToRtsIdMap = rtsIdToElementMap.ToDictionary(kvp => kvp.Value.Id, kvp => kvp.Key);
            var sequence = new List<string>();

            for (int i = 0; i < path.Count; i++)
            {
                var currentId = path[i];
                if (!elementIdToRtsIdMap.TryGetValue(currentId, out string rtsId)) continue;

                if (i > 0)
                {
                    var prevId = path[i - 1];
                    var current = Document.GetElement(currentId);
                    var prev = Document.GetElement(prevId);

                    if (current == null || prev == null) continue;

                    if (current.Category.Id != prev.Category.Id)
                    {
                        sequence.Add("+");
                    }
                    else if (graph.TryGetValue(prevId, out var neighbors) && !neighbors.Contains(currentId))
                    {
                        sequence.Add("||");
                    }
                    else
                    {
                        sequence.Add(",");
                    }
                }
                sequence.Add(rtsId);
            }
            return string.Join(" ", sequence).Replace(" , ", ", ").Replace(" + ", " + ").Replace(" || ", " || ");
        }

        /// <summary>
        /// Formats a list of disconnected paths into the final report string.
        /// </summary>
        private string FormatDisconnectedPaths(List<List<ElementId>> allPaths, Dictionary<ElementId, List<ElementId>> graph, Dictionary<string, Element> rtsIdMap)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < allPaths.Count; i++)
            {
                if (i > 0) sb.Append(" >> ");
                sb.Append(FormatPath(allPaths[i], graph, rtsIdMap));
            }
            return sb.ToString();
        }

        private bool AreConnected(Element a, Element b)
        {
            if (a == null || b == null) return false;
            try
            {
                var connectorsA = (a as MEPCurve)?.ConnectorManager?.Connectors;
                var connectorsB = (b as MEPCurve)?.ConnectorManager?.Connectors;

                if (connectorsA == null || connectorsB == null) return false;

                foreach (Connector cA in connectorsA)
                {
                    foreach (Connector cB in connectorsB)
                    {
                        if (cA.IsConnectedTo(cB) || cA.Origin.IsAlmostEqualTo(cB.Origin, 0.1))
                        {
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Finds all matching equipment elements, prioritizing RTS_ID over Panel Name.
        /// </summary>
        /// <returns>A list of all potential matching elements.</returns>
        private List<Element> FindMatchingEquipment(string idToMatch, List<Element> allCandidateEquipment)
        {
            if (string.IsNullOrWhiteSpace(idToMatch)) return new List<Element>();

            // First, try to find matches based on RTS_ID
            var rtsIdMatches = allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var rtsIdParam = e.get_Parameter(_rtsIdGuid);
                var rtsIdVal = rtsIdParam?.AsString();
                return !string.IsNullOrWhiteSpace(rtsIdVal) && idToMatch.IndexOf(rtsIdVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            // If any RTS_ID matches are found, return them and ignore Panel Name matches
            if (rtsIdMatches.Any())
            {
                return rtsIdMatches;
            }

            // If no RTS_ID matches, fall back to finding matches by Panel Name
            var panelNameMatches = allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var panelNameParam = e.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME);
                var panelNameVal = panelNameParam?.AsString();
                return !string.IsNullOrWhiteSpace(panelNameVal) && idToMatch.IndexOf(panelNameVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            return panelNameMatches;
        }

        /// <summary>
        /// Calculates the total physical length of a path of containment elements.
        /// </summary>
        private double GetPathLength(List<ElementId> pathIds)
        {
            if (pathIds == null || !pathIds.Any()) return 0.0;

            double totalLength = 0.0;
            foreach (var id in pathIds)
            {
                Element elem = Document.GetElement(id);
                if (elem is MEPCurve curve)
                {
                    totalLength += curve.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                }
            }
            return totalLength;
        }

        /// <summary>
        /// Extracts the branch number from an RTS_ID string.
        /// </summary>
        private string GetBranchNumber(string rtsId)
        {
            if (string.IsNullOrWhiteSpace(rtsId) || rtsId.Length < 4) return null;

            // The branch number is the last 4 characters of the ID
            return rtsId.Substring(rtsId.Length - 4);
        }

        /// <summary>
        /// Formats the branch sequence string from a given path.
        /// </summary>
        private string FormatBranchSequence(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap)
        {
            if (path == null || !path.Any()) return "N/A";

            var elementIdToRtsIdMap = rtsIdToElementMap.ToDictionary(kvp => kvp.Value.Id, kvp => kvp.Key);
            var uniqueBranches = new List<string>();
            var seenBranches = new HashSet<string>();
            bool hasAddedLeadingZeroBranch = false;

            foreach (var elementId in path)
            {
                if (elementIdToRtsIdMap.TryGetValue(elementId, out string rtsId))
                {
                    string branchNumber = GetBranchNumber(rtsId);
                    if (!string.IsNullOrEmpty(branchNumber) && seenBranches.Add(branchNumber))
                    {
                        // Only add apostrophe to the first branch with a leading zero
                        if (branchNumber.StartsWith("0") && !hasAddedLeadingZeroBranch)
                        {
                            uniqueBranches.Add($"'{branchNumber}");
                            hasAddedLeadingZeroBranch = true;
                        }
                        else
                        {
                            uniqueBranches.Add(branchNumber);
                        }
                    }
                }
            }
            return string.Join(", ", uniqueBranches);
        }

        #endregion
    }
}