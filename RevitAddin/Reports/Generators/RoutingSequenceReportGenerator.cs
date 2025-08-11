//-----------------------------------------------------------------------------
// <copyright file="RoutingSequenceReportGenerator.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Generates the Routing Sequence report using Dijkstra's algorithm for shortest path,
//   with virtual pathing between disconnected containment islands.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-09:
// - Implemented high-accuracy island finding. The logic now checks every element
//   within an island to find the true closest point, instead of using the
//   bounding box center, for more accurate start/end points.
// - Implemented high-accuracy jump calculation between islands using element-to-element
//   distance checks for maximum precision.
// - Added integrated completion state to the progress bar UI.
// - Fixed critical logic error in GetRelevantNetwork method.
// - Optimized GroupIntoIslands method by using a HashSet for faster lookups.
// - Implemented "just-in-time" graph building for significant performance improvement.
// - Refactored entire class for performance and maintainability.
// - Implemented a PriorityQueue for Dijkstra's algorithm.
// - Implemented intelligent separators for virtual jumps (+, ||, >>).
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Reports.Base;
using RTS.Reports.Utils;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Main class to generate the Routing Sequence Report.
    /// Orchestrates data collection, graph building, pathfinding, and report generation.
    /// </summary>
    public class RoutingSequenceReportGenerator : ReportGeneratorBase
    {
        private readonly Guid _rtsIdGuid = SharedParameters.Cable.RTS_ID_GUID;
        private readonly List<Guid> _cableGuids = SharedParameters.Cable.RTS_Cable_GUIDs;

        public RoutingSequenceReportGenerator(Document doc, ExternalCommandData commandData, PC_ExtensibleClass pcExtensible)
            : base(doc, commandData, pcExtensible)
        {
        }

        public override void GenerateReport()
        {
            // --- Global Exception Handler to prevent silent exits ---
            try
            {
                // --- PRE-CHECK 1: Verify that cable data exists in the project ---
                List<PC_ExtensibleClass.CableData> primaryData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    Document, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName);

                if (primaryData == null || !primaryData.Any())
                {
                    ShowInfo("No Data Found", "No primary cable data was found in the project's extensible storage. Please ensure cable data has been created before running this report.");
                    return;
                }

                // --- PRE-CHECK 2: Verify that at least one cable is assigned to containment ---
                var containmentCollector = new ContainmentCollector(Document);
                var allContainmentInProject = containmentCollector.GetAllContainmentElements();
                bool isAnyCableAssigned = allContainmentInProject.OfType<MEPCurve>().Any(elem =>
                    _cableGuids.Select(guid => elem.get_Parameter(guid))
                               .Any(param => param != null && param.HasValue && !string.IsNullOrWhiteSpace(param.AsString()))
                );

                if (!isAnyCableAssigned)
                {
                    ShowInfo("No Assignments Found", "Cable data was found, but no cables have been assigned to any containment elements in the model. The report would be empty.");
                    return;
                }

                // --- All checks passed, proceed with asking for a save location ---
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

                var sb = ReportFormatter.CreateReportHeader(projectInfo, Document);

                var cableLookup = primaryData
                    .Where(c => !string.IsNullOrWhiteSpace(c.CableReference))
                    .GroupBy(c => c.CableReference)
                    .ToDictionary(g => g.Key, g => g.First());

                var progressWindow = new RoutingReportProgressBarWindow
                {
                    Owner = System.Windows.Application.Current?.MainWindow
                };
                progressWindow.Show();

                try
                {
                    var allCandidateEquipment = containmentCollector.GetAllElectricalEquipment();
                    var allFittingsInProject = allContainmentInProject.Where(e => !(e is MEPCurve)).ToList();

                    int totalCables = cableLookup.Count;
                    int currentCable = 0;

                    // --- DEVELOPMENT LIMIT: Process only the first 20 cables. Remove .Take(20) for production. ---
                    foreach (var cableRef in cableLookup.Keys.Take(20))
                    {
                        currentCable++;
                        var cableInfo = cableLookup[cableRef];

                        double percentage = (double)currentCable / Math.Min(totalCables, 20) * 100.0; // Adjust percentage for the limited run
                        progressWindow.UpdateProgress(currentCable, Math.Min(totalCables, 20), cableRef,
                            cableInfo?.From ?? "N/A", cableInfo?.To ?? "N/A",
                            percentage, "Building local network...");
                        Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background);

                        if (progressWindow.IsCancelled)
                        {
                            progressWindow.TaskDescriptionText.Text = "Operation cancelled by user.";
                            break;
                        }

                        var result = ProcessCableRoute(cableRef, cableInfo, allContainmentInProject, allFittingsInProject, allCandidateEquipment);
                        sb.AppendLine(result);
                    }

                    if (!progressWindow.IsCancelled)
                    {
                        System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                        // Show the completion state on the progress window itself
                        progressWindow.ShowCompletion(filePath);
                    }
                }
                catch (Exception ex)
                {
                    progressWindow.ShowError($"An unexpected error occurred: {ex.Message}");
                    ShowError($"Failed to export Routing Sequence report: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // This will catch any error that happens anywhere in the command, including UI initialization.
                ShowError($"A critical error occurred: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Processes a single cable's route and returns the formatted CSV line.
        /// </summary>
        private string ProcessCableRoute(string cableRef, PC_ExtensibleClass.CableData cableInfo, List<Element> allContainmentInProject, List<Element> allFittingsInProject, List<Element> allCandidateEquipment)
        {
            string status = "Error";
            string routingSequence = "Could not determine route.";
            string branchSequence = "N/A";
            List<ElementId> bestPath = null;

            try
            {
                // Step 1: Find the specific trays/conduits this cable is assigned to.
                var cableContainmentElements = new List<Element>();
                foreach (Element elem in allContainmentInProject)
                {
                    if (!(elem is MEPCurve)) continue;

                    foreach (var guid in _cableGuids)
                    {
                        Parameter cableParam = elem.get_Parameter(guid);
                        if (string.Equals(ReportHelpers.CleanCableReference(cableParam?.AsString()), cableRef, StringComparison.OrdinalIgnoreCase))
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
                    // Step 2: Build a specific, local graph for ONLY the relevant elements.
                    var relevantNetwork = GetRelevantNetwork(cableContainmentElements, allFittingsInProject);
                    var containmentGraph = new ContainmentGraph(relevantNetwork, Document);
                    var elementIdToLengthMap = containmentGraph.GetElementLengthMap();
                    var rtsIdToElementMap = ReportFormatter.BuildRtsIdMap(relevantNetwork, _rtsIdGuid);

                    List<Element> startCandidates = FindMatchingEquipment(cableInfo.From, allCandidateEquipment, _rtsIdGuid);
                    List<Element> endCandidates = FindMatchingEquipment(cableInfo.To, allCandidateEquipment, _rtsIdGuid);

                    if (startCandidates.Any() && endCandidates.Any())
                    {
                        var pathResult = Pathfinder.FindConfirmedPath(startCandidates, endCandidates, cableContainmentElements, containmentGraph, elementIdToLengthMap);
                        if (pathResult != null && pathResult.Any())
                        {
                            status = "Confirmed";
                            bestPath = pathResult;
                        }
                        else
                        {
                            status = "Unconfirmed (Virtual)";
                            var virtualResult = Pathfinder.FindBestDisconnectedSequence(startCandidates, endCandidates, cableContainmentElements, containmentGraph, rtsIdToElementMap, elementIdToLengthMap, Document);
                            routingSequence = virtualResult.RoutingSequence;
                            bestPath = virtualResult.StitchedPath;
                        }
                    }
                    else
                    {
                        status = startCandidates.Any() ? "Confirmed (From, Virtual)" : endCandidates.Any() ? "Confirmed (To, Virtual)" : "Unconfirmed (Virtual)";
                        var virtualResult = Pathfinder.FindBestDisconnectedSequence(startCandidates, endCandidates, cableContainmentElements, containmentGraph, rtsIdToElementMap, elementIdToLengthMap, Document);
                        routingSequence = virtualResult.RoutingSequence;
                        bestPath = virtualResult.StitchedPath;
                    }

                    if (bestPath != null && bestPath.Any())
                    {
                        if (status == "Confirmed")
                        {
                            routingSequence = ReportFormatter.FormatPath(bestPath, containmentGraph, rtsIdToElementMap, Document);
                        }
                        branchSequence = ReportFormatter.FormatBranchSequence(bestPath, rtsIdToElementMap);
                    }
                }
            }
            catch (Exception ex)
            {
                status = "Error";
                routingSequence = $"Processing failed: {ex.Message}";
                branchSequence = "Error";
            }

            return $"\"{cableRef}\",\"{cableInfo.From ?? "N/A"}\",\"{cableInfo.To ?? "N/A"}\",\"{status}\",\"{branchSequence}\",\"{routingSequence}\"";
        }

        /// <summary>
        /// Finds all fittings and other MEPCurves connected to a given set of containment elements to build a complete local network.
        /// </summary>
        private List<Element> GetRelevantNetwork(List<Element> cableContainment, List<Element> allFittings)
        {
            var relevantElements = new Dictionary<ElementId, Element>();
            var queue = new Queue<Element>(cableContainment);

            foreach (var el in cableContainment)
            {
                if (!relevantElements.ContainsKey(el.Id))
                    relevantElements.Add(el.Id, el);
            }

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                var connectors = ContainmentGraph.GetConnectors(current);
                if (connectors == null) continue;

                foreach (Connector conn in connectors)
                {
                    if (!conn.IsConnected) continue;

                    foreach (Connector connected in conn.AllRefs)
                    {
                        var owner = connected.Owner;
                        if (owner != null && !relevantElements.ContainsKey(owner.Id))
                        {
                            relevantElements.Add(owner.Id, owner);
                            queue.Enqueue(owner);
                        }
                    }
                }
            }
            return relevantElements.Values.ToList();
        }

        private static List<Element> FindMatchingEquipment(string idToMatch, List<Element> allCandidateEquipment, Guid rtsIdGuid)
        {
            if (string.IsNullOrWhiteSpace(idToMatch)) return new List<Element>();

            var rtsIdMatches = allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var rtsIdParam = e.get_Parameter(rtsIdGuid);
                var rtsIdVal = rtsIdParam?.AsString();
                return !string.IsNullOrWhiteSpace(rtsIdVal) && idToMatch.IndexOf(rtsIdVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            if (rtsIdMatches.Any()) return rtsIdMatches;

            return allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var panelNameParam = e.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME);
                var panelNameVal = panelNameParam?.AsString();
                return !string.IsNullOrWhiteSpace(panelNameVal) && idToMatch.IndexOf(panelNameVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();
        }
    }

    #region Helper Classes

    public class ContainmentCollector
    {
        private readonly Document _doc;
        public ContainmentCollector(Document doc) { _doc = doc; }

        public List<Element> GetAllContainmentElements()
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Conduit, BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_CableTrayFitting
            };
            var multiCatFilter = new ElementMulticategoryFilter(categories);
            return new FilteredElementCollector(_doc).WherePasses(multiCatFilter).WhereElementIsNotElementType().ToList();
        }

        public List<Element> GetAllElectricalEquipment()
        {
            var equip = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToList();
            var fixtures = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().ToList();
            return equip.Concat(fixtures).ToList();
        }
    }

    public class ContainmentGraph
    {
        public Dictionary<ElementId, List<ElementId>> AdjacencyGraph { get; }
        private readonly Document _doc;

        // Constructor for cable-specific graphs (no progress reporting)
        public ContainmentGraph(List<Element> elements, Document doc)
        {
            _doc = doc;
            AdjacencyGraph = BuildAdjacencyGraph(elements);
        }

        private Dictionary<ElementId, List<ElementId>> BuildAdjacencyGraph(List<Element> elements)
        {
            var graph = new Dictionary<ElementId, List<ElementId>>();
            foreach (var elemA in elements)
            {
                if (!graph.ContainsKey(elemA.Id)) graph[elemA.Id] = new List<ElementId>();
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

        public Dictionary<ElementId, double> GetElementLengthMap()
        {
            return AdjacencyGraph.Keys.ToDictionary(id => id, id => (_doc.GetElement(id) as MEPCurve)?.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0);
        }

        public static ConnectorSet GetConnectors(Element element)
        {
            if (element is MEPCurve curve) return curve.ConnectorManager?.Connectors;
            if (element is FamilyInstance fi) return fi.MEPModel?.ConnectorManager?.Connectors;
            return null;
        }

        private static bool AreConnected(Element a, Element b)
        {
            try
            {
                var connectorsA = GetConnectors(a);
                var connectorsB = GetConnectors(b);
                if (connectorsA == null || connectorsB == null) return false;

                foreach (Connector cA in connectorsA)
                {
                    if (!cA.IsConnected) continue;
                    foreach (Connector cB in connectorsB)
                    {
                        if (cA.IsConnectedTo(cB)) return true;
                    }
                }
            }
            catch { }
            return false;
        }
    }

    public static class Pathfinder
    {
        private const double METERS_TO_FEET = 3.28084;

        public class VirtualPathResult
        {
            public string RoutingSequence { get; set; } = "Could not determine virtual route.";
            public List<ElementId> StitchedPath { get; set; } = new List<ElementId>();
        }

        #region Public Pathfinding Methods

        public static List<ElementId> FindConfirmedPath(List<Element> startCandidates, List<Element> endCandidates, List<Element> cableElements, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap)
        {
            List<ElementId> bestPath = null;
            double minLength = double.MaxValue;

            foreach (var startCandidate in startCandidates)
            {
                foreach (var endCandidate in endCandidates)
                {
                    XYZ startPt = ReportHelpers.GetElementLocation(startCandidate);
                    XYZ endPt = ReportHelpers.GetElementLocation(endCandidate);
                    if (startPt == null || endPt == null) continue;

                    Element startElem = cableElements.OrderBy(e => ReportHelpers.GetElementLocation(e)?.DistanceTo(startPt) ?? double.MaxValue).FirstOrDefault();
                    Element endElem = cableElements.OrderBy(e => ReportHelpers.GetElementLocation(e)?.DistanceTo(endPt) ?? double.MaxValue).FirstOrDefault();

                    if (startElem == null || endElem == null) continue;

                    var currentPath = FindShortestPath_Dijkstra(startElem.Id, endElem.Id, graph, lengthMap);
                    if (currentPath != null)
                    {
                        double currentLength = GetPathLength(currentPath, lengthMap);
                        if (currentLength < minLength)
                        {
                            minLength = currentLength;
                            bestPath = currentPath;
                        }
                    }
                }
            }
            return bestPath;
        }

        public static VirtualPathResult FindBestDisconnectedSequence(List<Element> startCandidates, List<Element> endCandidates, List<Element> cableElements, ContainmentGraph graph, Dictionary<string, Element> rtsIdToElementMap, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            if (!cableElements.Any()) return new VirtualPathResult();

            var islands = GroupIntoIslands(cableElements, graph, doc);
            if (!islands.Any()) return new VirtualPathResult();

            ContainmentIsland startIsland = FindClosestIsland(startCandidates, islands, doc);
            ContainmentIsland endIsland = FindClosestIsland(endCandidates, islands, doc);

            if (startIsland == null && endIsland == null)
            {
                var disconnectedPaths = islands.Select(i => i.ElementIds.ToList()).ToList();
                return new VirtualPathResult { RoutingSequence = ReportFormatter.FormatDisconnectedPaths(disconnectedPaths, graph, rtsIdToElementMap, doc) };
            }

            startIsland = startIsland ?? endIsland;
            endIsland = endIsland ?? startIsland;

            var islandGraph = BuildIslandGraph(islands, doc);
            var islandPath = FindShortestIslandPath(startIsland, endIsland, islandGraph);
            if (islandPath == null || !islandPath.Any()) return new VirtualPathResult();

            return StitchPaths(islandPath, startCandidates, endCandidates, islandGraph, graph, rtsIdToElementMap, lengthMap, doc);
        }

        #endregion

        #region Island Hopping Logic

        private enum ContainmentType { Tray, Conduit, Mixed }

        private class ContainmentIsland
        {
            public int Id { get; }
            public HashSet<ElementId> ElementIds { get; } = new HashSet<ElementId>();
            public BoundingBoxXYZ BoundingBox { get; set; }
            public ContainmentType Type { get; set; } = ContainmentType.Mixed;
            public ContainmentIsland(int id) { Id = id; }
        }

        private static List<ContainmentIsland> GroupIntoIslands(List<Element> cableElements, ContainmentGraph graph, Document doc)
        {
            var islands = new List<ContainmentIsland>();
            var visited = new HashSet<ElementId>();
            int islandIdCounter = 0;
            var cableElementIds = new HashSet<ElementId>(cableElements.Select(e => e.Id));

            foreach (var element in cableElements)
            {
                if (visited.Contains(element.Id)) continue;

                var newIsland = new ContainmentIsland(islandIdCounter++);
                var queue = new Queue<ElementId>();
                queue.Enqueue(element.Id);
                visited.Add(element.Id);

                var islandBoundingBox = new BoundingBoxXYZ { Enabled = false };

                while (queue.Count > 0)
                {
                    var currentId = queue.Dequeue();
                    newIsland.ElementIds.Add(currentId);
                    var currentElem = doc.GetElement(currentId);
                    if (currentElem != null)
                    {
                        var bb = currentElem.get_BoundingBox(null);
                        if (bb != null)
                        {
                            if (!islandBoundingBox.Enabled)
                            {
                                islandBoundingBox = bb;
                                islandBoundingBox.Enabled = true;
                            }
                            else
                            {
                                islandBoundingBox.Min = new XYZ(Math.Min(islandBoundingBox.Min.X, bb.Min.X), Math.Min(islandBoundingBox.Min.Y, bb.Min.Y), Math.Min(islandBoundingBox.Min.Z, bb.Min.Z));
                                islandBoundingBox.Max = new XYZ(Math.Max(islandBoundingBox.Max.X, bb.Max.X), Math.Max(islandBoundingBox.Max.Y, bb.Max.Y), Math.Max(islandBoundingBox.Max.Z, bb.Max.Z));
                            }
                        }
                    }

                    if (graph.AdjacencyGraph.TryGetValue(currentId, out var neighbors))
                    {
                        foreach (var neighborId in neighbors)
                        {
                            if (cableElementIds.Contains(neighborId) && !visited.Contains(neighborId))
                            {
                                visited.Add(neighborId);
                                queue.Enqueue(neighborId);
                            }
                        }
                    }
                }
                newIsland.BoundingBox = islandBoundingBox;
                newIsland.Type = GetIslandContainmentType(newIsland, doc);
                islands.Add(newIsland);
            }
            return islands;
        }

        private static ContainmentType GetIslandContainmentType(ContainmentIsland island, Document doc)
        {
            bool hasTray = false;
            bool hasConduit = false;
            foreach (var elemId in island.ElementIds)
            {
                var elem = doc.GetElement(elemId);
                if (elem == null) continue;
#if REVIT2024_OR_GREATER
                var catId = elem.Category.Id.Value;
#else
                var catId = elem.Category.Id.IntegerValue;
#endif
                if (catId == (int)BuiltInCategory.OST_CableTray || catId == (int)BuiltInCategory.OST_CableTrayFitting) hasTray = true;
                else if (catId == (int)BuiltInCategory.OST_Conduit || catId == (int)BuiltInCategory.OST_ConduitFitting) hasConduit = true;
                if (hasTray && hasConduit) return ContainmentType.Mixed;
            }
            if (hasTray) return ContainmentType.Tray;
            if (hasConduit) return ContainmentType.Conduit;
            return ContainmentType.Mixed;
        }

        private static ContainmentIsland FindClosestIsland(List<Element> equipmentCandidates, List<ContainmentIsland> islands, Document doc)
        {
            if (equipmentCandidates == null || !equipmentCandidates.Any()) return null;

            ContainmentIsland bestIsland = null;
            double minDistance = double.MaxValue;

            foreach (var equip in equipmentCandidates)
            {
                var equipLocation = ReportHelpers.GetElementLocation(equip);
                if (equipLocation == null) continue;

                foreach (var island in islands)
                {
                    foreach (var elementId in island.ElementIds)
                    {
                        var elem = doc.GetElement(elementId);
                        var elemLocation = ReportHelpers.GetElementLocation(elem);
                        if (elemLocation == null) continue;

                        double dist = equipLocation.DistanceTo(elemLocation);
                        if (dist < minDistance)
                        {
                            minDistance = dist;
                            bestIsland = island;
                        }
                    }
                }
            }
            return bestIsland;
        }

        private static Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> BuildIslandGraph(List<ContainmentIsland> islands, Document doc)
        {
            var islandGraph = new Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>>();
            for (int i = 0; i < islands.Count; i++)
            {
                islandGraph[islands[i]] = new List<Tuple<ContainmentIsland, double>>();
                for (int j = i + 1; j < islands.Count; j++)
                {
                    if (!islandGraph.ContainsKey(islands[j])) islandGraph[islands[j]] = new List<Tuple<ContainmentIsland, double>>();

                    // High-accuracy distance calculation
                    double minDistance = double.MaxValue;
                    foreach (var idA in islands[i].ElementIds)
                    {
                        var locA = ReportHelpers.GetElementLocation(doc.GetElement(idA));
                        if (locA == null) continue;

                        foreach (var idB in islands[j].ElementIds)
                        {
                            var locB = ReportHelpers.GetElementLocation(doc.GetElement(idB));
                            if (locB == null) continue;
                            minDistance = Math.Min(minDistance, locA.DistanceTo(locB));
                        }
                    }

                    if (minDistance != double.MaxValue)
                    {
                        islandGraph[islands[i]].Add(new Tuple<ContainmentIsland, double>(islands[j], minDistance));
                        islandGraph[islands[j]].Add(new Tuple<ContainmentIsland, double>(islands[i], minDistance));
                    }
                }
            }
            return islandGraph;
        }

        private static List<ContainmentIsland> FindShortestIslandPath(ContainmentIsland startIsland, ContainmentIsland endIsland, Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> islandGraph)
        {
            var distances = islandGraph.Keys.ToDictionary(i => i.Id, i => double.PositiveInfinity);
            var previous = new Dictionary<int, int?>();
            var queue = new PriorityQueue<ContainmentIsland>();

            distances[startIsland.Id] = 0;
            queue.Enqueue(startIsland, 0);

            while (queue.Count > 0)
            {
                var smallest = queue.Dequeue();
                if (smallest.Id == endIsland.Id)
                {
                    var path = new List<ContainmentIsland>();
                    int? currentId = endIsland.Id;
                    while (currentId.HasValue)
                    {
                        path.Add(islandGraph.Keys.First(i => i.Id == currentId.Value));
                        previous.TryGetValue(currentId.Value, out currentId);
                    }
                    path.Reverse();
                    return path;
                }
                if (distances[smallest.Id] == double.PositiveInfinity) break;
                foreach (var neighborTuple in islandGraph[smallest])
                {
                    var neighbor = neighborTuple.Item1;
                    var jumpDistance = neighborTuple.Item2;
                    var alt = distances[smallest.Id] + jumpDistance;
                    if (alt < distances[neighbor.Id])
                    {
                        distances[neighbor.Id] = alt;
                        previous[neighbor.Id] = smallest.Id;
                        queue.Enqueue(neighbor, alt);
                    }
                }
            }
            return null;
        }

        private static VirtualPathResult StitchPaths(List<ContainmentIsland> islandPath, List<Element> startCandidates, List<Element> endCandidates, Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> islandGraph, ContainmentGraph graph, Dictionary<string, Element> rtsIdToElementMap, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            var finalSequence = new StringBuilder();
            var finalStitchedPath = new List<ElementId>();
            ElementId lastElementId = null;

            if (startCandidates.Any())
            {
                lastElementId = FindClosestElementInIsland(startCandidates.First(), islandPath.First(), doc);
            }

            for (int i = 0; i < islandPath.Count; i++)
            {
                var currentIsland = islandPath[i];
                ElementId startOfSegment;
                ElementId endOfSegment = null;

                if (i > 0)
                {
                    var prevIsland = islandPath[i - 1];
                    var jumpTuple = islandGraph[prevIsland].FirstOrDefault(t => t.Item1.Id == currentIsland.Id);
                    double jumpDistance = jumpTuple?.Item2 ?? 0.0;
                    if (prevIsland.Type != ContainmentType.Mixed && prevIsland.Type == currentIsland.Type)
                    {
                        finalSequence.Append(jumpDistance <= METERS_TO_FEET ? " || " : " >> ");
                    }
                    else
                    {
                        finalSequence.Append(" + ");
                    }
                    startOfSegment = FindClosestElementInIsland(doc.GetElement(lastElementId), currentIsland, doc);
                }
                else
                {
                    startOfSegment = lastElementId ?? currentIsland.ElementIds.First();
                }

                if (i == islandPath.Count - 1)
                {
                    endOfSegment = endCandidates.Any() ? FindClosestElementInIsland(endCandidates.First(), currentIsland, doc) : FindFarthestElementInIsland(doc.GetElement(startOfSegment), currentIsland, doc);
                }
                else
                {
                    var nextIsland = islandPath[i + 1];
                    endOfSegment = FindClosestElementToOtherIsland(currentIsland, nextIsland, doc);
                }

                startOfSegment = startOfSegment ?? currentIsland.ElementIds.First();
                endOfSegment = endOfSegment ?? startOfSegment;

                var segmentPath = FindShortestPath_Dijkstra(startOfSegment, endOfSegment, graph, lengthMap);
                if (segmentPath != null && segmentPath.Any())
                {
                    finalSequence.Append(ReportFormatter.FormatPath(segmentPath, graph, rtsIdToElementMap, doc));
                    finalStitchedPath.AddRange(segmentPath);
                    lastElementId = segmentPath.Last();
                }
                else
                {
                    var fallbackPath = new List<ElementId> { startOfSegment };
                    finalSequence.Append(ReportFormatter.FormatPath(fallbackPath, graph, rtsIdToElementMap, doc));
                    finalStitchedPath.AddRange(fallbackPath);
                    lastElementId = startOfSegment;
                }
            }
            return new VirtualPathResult { RoutingSequence = finalSequence.ToString(), StitchedPath = finalStitchedPath };
        }

        private static ElementId FindClosestElementInIsland(Element targetElement, ContainmentIsland island, Document doc)
        {
            var targetLocation = ReportHelpers.GetElementLocation(targetElement);
            if (targetLocation == null) return island.ElementIds.FirstOrDefault();
            return island.ElementIds.OrderBy(id => ReportHelpers.GetElementLocation(doc.GetElement(id))?.DistanceTo(targetLocation) ?? double.MaxValue).FirstOrDefault();
        }

        private static ElementId FindFarthestElementInIsland(Element targetElement, ContainmentIsland island, Document doc)
        {
            var targetLocation = ReportHelpers.GetElementLocation(targetElement);
            if (targetLocation == null) return island.ElementIds.FirstOrDefault();
            return island.ElementIds.OrderByDescending(id => ReportHelpers.GetElementLocation(doc.GetElement(id))?.DistanceTo(targetLocation) ?? 0).FirstOrDefault();
        }

        private static ElementId FindClosestElementToOtherIsland(ContainmentIsland fromIsland, ContainmentIsland toIsland, Document doc)
        {
            ElementId bestElementId = fromIsland.ElementIds.FirstOrDefault();
            double minDistance = double.MaxValue;

            foreach (var idFrom in fromIsland.ElementIds)
            {
                var locFrom = ReportHelpers.GetElementLocation(doc.GetElement(idFrom));
                if (locFrom == null) continue;

                foreach (var idTo in toIsland.ElementIds)
                {
                    var locTo = ReportHelpers.GetElementLocation(doc.GetElement(idTo));
                    if (locTo == null) continue;

                    double dist = locFrom.DistanceTo(locTo);
                    if (dist < minDistance)
                    {
                        minDistance = dist;
                        bestElementId = idFrom;
                    }
                }
            }
            return bestElementId;
        }

        #endregion

        #region Core Dijkstra Algorithm

        private static List<ElementId> FindShortestPath_Dijkstra(ElementId startNodeId, ElementId endNodeId, ContainmentGraph graph, Dictionary<ElementId, double> elementIdToLengthMap)
        {
            if (startNodeId == null || endNodeId == null) return null;
            if (startNodeId == endNodeId) return new List<ElementId> { startNodeId };

            var distances = new Dictionary<ElementId, double>();
            var previous = new Dictionary<ElementId, ElementId>();
            var queue = new PriorityQueue<ElementId>();

            foreach (var nodeId in graph.AdjacencyGraph.Keys)
            {
                distances[nodeId] = double.PositiveInfinity;
            }

            if (!distances.ContainsKey(startNodeId)) return null;
            distances[startNodeId] = 0;
            queue.Enqueue(startNodeId, 0);

            while (queue.Count > 0)
            {
                var currentNodeId = queue.Dequeue();
                if (currentNodeId == endNodeId)
                {
                    var path = new List<ElementId>();
                    var at = endNodeId;
                    while (at != null && at != ElementId.InvalidElementId)
                    {
                        path.Add(at);
                        if (!previous.TryGetValue(at, out at)) break;
                    }
                    path.Reverse();
                    return path;
                }

                if (!graph.AdjacencyGraph.TryGetValue(currentNodeId, out var neighbors)) continue;

                foreach (var neighborId in neighbors)
                {
                    double costToNeighbor = elementIdToLengthMap.ContainsKey(neighborId) ? elementIdToLengthMap[neighborId] : 0.0;
                    double altDistance = distances[currentNodeId] + costToNeighbor;
                    if (altDistance < distances[neighborId])
                    {
                        distances[neighborId] = altDistance;
                        previous[neighborId] = currentNodeId;
                        queue.Enqueue(neighborId, altDistance);
                    }
                }
            }
            return null;
        }

        private static double GetPathLength(List<ElementId> pathIds, Dictionary<ElementId, double> lengthMap)
        {
            if (pathIds == null || !pathIds.Any()) return 0.0;
            return pathIds.Sum(id => lengthMap.ContainsKey(id) ? lengthMap[id] : 0.0);
        }

        #endregion

        #region Nested PriorityQueue for .NET 4.8

        /// <summary>
        /// A simple Priority Queue implementation for use in .NET Framework 4.8.
        /// </summary>
        private class PriorityQueue<T>
        {
            private readonly List<Tuple<T, double>> _elements = new List<Tuple<T, double>>();
            public int Count => _elements.Count;

            public void Enqueue(T item, double priority)
            {
                _elements.Add(Tuple.Create(item, priority));
            }

            public T Dequeue()
            {
                int bestIndex = 0;
                for (int i = 0; i < _elements.Count; i++)
                {
                    if (_elements[i].Item2 < _elements[bestIndex].Item2)
                    {
                        bestIndex = i;
                    }
                }
                T bestItem = _elements[bestIndex].Item1;
                _elements.RemoveAt(bestIndex);
                return bestItem;
            }
        }

        #endregion
    }

    /// <summary>
    /// Handles formatting of the final report strings.
    /// </summary>
    public static class ReportFormatter
    {
        public static StringBuilder CreateReportHeader(ProjectInfo projectInfo, Document doc)
        {
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
            return sb;
        }

        public static Dictionary<string, Element> BuildRtsIdMap(List<Element> elements, Guid rtsIdGuid)
        {
            var map = new Dictionary<string, Element>();
            foreach (var elem in elements)
            {
                string rtsId = elem.get_Parameter(rtsIdGuid)?.AsString();
                if (!string.IsNullOrWhiteSpace(rtsId) && !map.ContainsKey(rtsId))
                {
                    map.Add(rtsId, elem);
                }
            }
            return map;
        }

        public static string FormatPath(List<ElementId> path, ContainmentGraph graph, Dictionary<string, Element> rtsIdToElementMap, Document doc)
        {
            if (path == null || !path.Any()) return "Path not found";
            var sequence = new List<string>();
            for (int i = 0; i < path.Count; i++)
            {
                var currentId = path[i];
                var currentElem = doc.GetElement(currentId);
                if (currentElem == null) continue;
                string rtsId = rtsIdToElementMap.FirstOrDefault(kvp => kvp.Value.Id == currentId).Key;
                if (string.IsNullOrEmpty(rtsId)) continue;

                if (i > 0)
                {
                    var prevElem = doc.GetElement(path[i - 1]);
                    if (prevElem != null && currentElem.Category.Id != prevElem.Category.Id)
                    {
                        sequence.Add("+");
                    }
                    else
                    {
                        sequence.Add(",");
                    }
                }
                sequence.Add(rtsId);
            }
            return string.Join(" ", sequence).Replace(" , ", ", ").Replace(" + ", " + ");
        }

        public static string FormatDisconnectedPaths(List<List<ElementId>> allPaths, ContainmentGraph graph, Dictionary<string, Element> rtsIdMap, Document doc)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < allPaths.Count; i++)
            {
                if (i > 0) sb.Append(" >> ");
                sb.Append(FormatPath(allPaths[i], graph, rtsIdMap, doc));
            }
            return sb.ToString();
        }

        public static string FormatBranchSequence(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap)
        {
            if (path == null || !path.Any()) return "N/A";
            var uniqueBranches = new List<string>();
            var seenBranches = new HashSet<string>();
            bool hasAddedLeadingZeroBranch = false;

            foreach (var elementId in path)
            {
                string rtsId = rtsIdToElementMap.FirstOrDefault(kvp => kvp.Value.Id == elementId).Key;
                if (!string.IsNullOrEmpty(rtsId))
                {
                    string branchNumber = rtsId.Length >= 4 ? rtsId.Substring(rtsId.Length - 4) : null;
                    if (!string.IsNullOrEmpty(branchNumber) && seenBranches.Add(branchNumber))
                    {
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
    }

    #endregion
}
