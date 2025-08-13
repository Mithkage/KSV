//-----------------------------------------------------------------------------
// <copyright file="RoutingSequenceReportGenerator.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright//
// <summary>
//   Generates the Routing Sequence report using Dijkstra's algorithm for shortest path,
//   with virtual pathing between disconnected containment islands.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Made helper classes (Pathfinder, ContainmentIsland, etc.) public to allow access from diagnostic tools.
// - [APPLIED FIX]: Removed CleanCableReference logic, assuming data is pre-cleaned in the model.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Reports.Base;
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
    public class RoutingSequenceReportGenerator : ReportGeneratorBase
    {
        private readonly Guid _rtsIdGuid = SharedParameters.General.RTS_ID;
        private readonly List<Guid> _cableGuids = SharedParameters.Cable.AllCableGuids;

        public RoutingSequenceReportGenerator(Document doc, ExternalCommandData commandData, PC_ExtensibleClass pcExtensible)
            : base(doc, commandData, pcExtensible)
        {
        }

        public override void GenerateReport()
        {
            try
            {
                List<PC_ExtensibleClass.CableData> primaryData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    Document, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName);

                if (primaryData == null || !primaryData.Any())
                {
                    ShowInfo("No Data Found", "No primary cable data was found in the project's extensible storage.");
                    return;
                }

                var containmentCollector = new ContainmentCollector(Document);
                var allContainmentInProject = containmentCollector.GetAllContainmentElements();
                if (!allContainmentInProject.Any(elem => _cableGuids.Any(guid => elem.get_Parameter(guid)?.HasValue ?? false)))
                {
                    ShowInfo("No Assignments Found", "No cables have been assigned to any containment elements.");
                    return;
                }

                string filePath = GetOutputFilePath("Routing_Sequence.csv", "Save Routing Sequence Report");
                if (string.IsNullOrEmpty(filePath)) return;

                var projectInfo = new FilteredElementCollector(Document).OfCategory(BuiltInCategory.OST_ProjectInformation).WhereElementIsNotElementType().Cast<ProjectInfo>().FirstOrDefault();
                var sb = ReportFormatter.CreateReportHeader(projectInfo, Document);
                var cableLookup = primaryData.Where(c => !string.IsNullOrWhiteSpace(c.CableReference)).GroupBy(c => c.CableReference).ToDictionary(g => g.Key, g => g.First());

                var progressWindow = new RoutingReportProgressBarWindow { Owner = System.Windows.Application.Current?.MainWindow };
                progressWindow.Show();

                try
                {
                    var allCandidateEquipment = containmentCollector.GetAllElectricalEquipment();
                    var allFittingsInProject = allContainmentInProject.Where(e => !(e is MEPCurve)).ToList();
                    int totalCables = cableLookup.Count;
                    int currentCable = 0;

                    foreach (var cableRef in cableLookup.Keys.Take(20)) // DEV LIMIT
                    {
                        currentCable++;
                        var cableInfo = cableLookup[cableRef];
                        int progressTotal = Math.Min(totalCables, 20);
                        double percentage = (double)currentCable / progressTotal * 100.0;
                        progressWindow.UpdateProgress(currentCable, progressTotal, cableRef, cableInfo?.From ?? "N/A", cableInfo?.To ?? "N/A", percentage, "Building local network...");
                        Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background);

                        if (progressWindow.IsCancelled) break;

                        var result = ProcessCableRoute(cableRef, cableInfo, allContainmentInProject, allFittingsInProject, allCandidateEquipment);
                        sb.AppendLine(result);
                    }

                    if (!progressWindow.IsCancelled)
                    {
                        File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
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
                ShowError($"A critical error occurred: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
            }
        }

        private string ProcessCableRoute(string cableRef, PC_ExtensibleClass.CableData cableInfo, List<Element> allContainmentInProject, List<Element> allFittingsInProject, List<Element> allCandidateEquipment)
        {
            string status = "Processing Error";
            string fromStatus = "Not Specified";
            string toStatus = "Not Specified";
            string routingSequence = "Could not determine route.";
            string branchSequence = "N/A";
            double virtualLength = 0.0;

            try
            {
                List<Element> startCandidates;
                if (!string.IsNullOrWhiteSpace(cableInfo.From))
                {
                    startCandidates = FindMatchingEquipment(cableInfo.From, allCandidateEquipment, _rtsIdGuid, Document);
                    if (startCandidates.Any())
                    {
                        string matchedName = startCandidates.First().get_Parameter(_rtsIdGuid)?.AsString() ?? startCandidates.First().get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "Unknown";
                        fromStatus = $"Found: {matchedName}";
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

                List<Element> endCandidates;
                if (!string.IsNullOrWhiteSpace(cableInfo.To))
                {
                    endCandidates = FindMatchingEquipment(cableInfo.To, allCandidateEquipment, _rtsIdGuid, Document);
                    if (endCandidates.Any())
                    {
                        string matchedName = endCandidates.First().get_Parameter(_rtsIdGuid)?.AsString() ?? endCandidates.First().get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "Unknown";
                        toStatus = $"Found: {matchedName}";
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
                    var relevantNetwork = GetRelevantNetwork(cableContainmentElements, allFittingsInProject);
                    var containmentGraph = new ContainmentGraph(relevantNetwork, Document);
                    var elementIdToLengthMap = containmentGraph.GetElementLengthMap();
                    var rtsIdToElementMap = ReportFormatter.BuildRtsIdMap(relevantNetwork, _rtsIdGuid);

                    List<ElementId> bestPath = null;

                    if (startCandidates.Any() && endCandidates.Any())
                    {
                        var pathResult = Pathfinder.FindConfirmedPath(startCandidates, endCandidates, cableContainmentElements, containmentGraph, elementIdToLengthMap, Document);
                        if (pathResult != null && pathResult.Any())
                        {
                            status = "Route Confirmed";
                            bestPath = pathResult;
                        }
                    }

                    if (bestPath == null)
                    {
                        if (!fromStatus.StartsWith("Found") || !toStatus.StartsWith("Found"))
                        {
                            status = "Route Incomplete (Start/End Not Found)";
                        }
                        else
                        {
                            status = "Route Unconfirmed (Virtual Path)";
                        }

                        var virtualResult = Pathfinder.FindBestDisconnectedSequence(startCandidates, endCandidates, cableContainmentElements, containmentGraph, rtsIdToElementMap, elementIdToLengthMap, Document);
                        routingSequence = virtualResult.RoutingSequence;
                        bestPath = virtualResult.StitchedPath;
                        virtualLength = virtualResult.VirtualLength;
                    }

                    if (bestPath != null && bestPath.Any())
                    {
                        if (status == "Route Confirmed")
                        {
                            routingSequence = ReportFormatter.FormatPath(bestPath, rtsIdToElementMap, Document, _rtsIdGuid);
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

            double virtualLengthInMeters = virtualLength * 0.3048;
            return $"\"{cableRef}\",\"{cableInfo.From ?? "N/A"}\",\"{cableInfo.To ?? "N/A"}\",\"{fromStatus}\",\"{toStatus}\",\"{status}\",\"{virtualLengthInMeters:F2}\",\"{branchSequence}\",\"{routingSequence}\"";
        }

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

        private static List<Element> FindMatchingEquipment(string nameToMatch, List<Element> allCandidateEquipment, Guid rtsIdGuid, Document doc)
        {
            if (string.IsNullOrWhiteSpace(nameToMatch)) return new List<Element>();

            string cleanNameToMatch = nameToMatch.Trim();

            var rtsIdMatches = allCandidateEquipment.Where(e =>
            {
                var rtsIdVal = e?.get_Parameter(rtsIdGuid)?.AsString();
                if (string.IsNullOrWhiteSpace(rtsIdVal)) return false;

                string cleanRtsIdVal = rtsIdVal.Trim();
                return cleanRtsIdVal.IndexOf(cleanNameToMatch, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       cleanNameToMatch.IndexOf(cleanRtsIdVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            if (rtsIdMatches.Any()) return rtsIdMatches;

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
            return new FilteredElementCollector(_doc)
                .WherePasses(new ElementMulticategoryFilter(categories))
                .WhereElementIsNotElementType().ToList();
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

        public ContainmentGraph(List<Element> elements, Document doc)
        {
            _doc = doc;
            AdjacencyGraph = BuildAdjacencyGraph(elements);
        }

        private Dictionary<ElementId, List<ElementId>> BuildAdjacencyGraph(List<Element> elements)
        {
            var graph = new Dictionary<ElementId, List<ElementId>>();
            var elementDict = elements.ToDictionary(e => e.Id);

            foreach (var elem in elements)
            {
                if (!graph.ContainsKey(elem.Id))
                {
                    graph[elem.Id] = new List<ElementId>();
                }
            }

            foreach (var elem in elements)
            {
                try
                {
                    var connectors = GetConnectors(elem);
                    if (connectors == null || connectors.IsEmpty) continue;

#if REVIT2024_OR_GREATER
                    long catId = elem.Category.Id.Value;
#else
                    long catId = elem.Category.Id.IntegerValue;
#endif
                    bool isFitting = catId == (long)BuiltInCategory.OST_CableTrayFitting || catId == (long)BuiltInCategory.OST_ConduitFitting;

                    if (isFitting && elem is FamilyInstance fitting)
                    {
                        var connectedCurves = new List<ElementId>();
                        foreach (Connector fittingConnector in connectors)
                        {
                            if (!fittingConnector.IsConnected) continue;
                            foreach (Connector connected in fittingConnector.AllRefs)
                            {
                                if (elementDict.ContainsKey(connected.Owner.Id) && connected.Owner is MEPCurve)
                                {
                                    connectedCurves.Add(connected.Owner.Id);
                                    graph[connected.Owner.Id].Add(fitting.Id);
                                    graph[fitting.Id].Add(connected.Owner.Id);
                                }
                            }
                        }

                        if (connectedCurves.Count > 1)
                        {
                            for (int i = 0; i < connectedCurves.Count; i++)
                            {
                                for (int j = i + 1; j < connectedCurves.Count; j++)
                                {
                                    graph[connectedCurves[i]].Add(connectedCurves[j]);
                                    graph[connectedCurves[j]].Add(connectedCurves[i]);
                                }
                            }
                        }
                    }
                    else if (elem is MEPCurve curve)
                    {
                        foreach (Connector c in connectors)
                        {
                            if (!c.IsConnected) continue;
                            foreach (Connector connected in c.AllRefs)
                            {
                                if (elementDict.ContainsKey(connected.Owner.Id))
                                {
                                    graph[curve.Id].Add(connected.Owner.Id);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not process connectors for element {elem.Id}: {ex.Message}");
                }
            }

            foreach (var key in graph.Keys.ToList())
            {
                graph[key] = graph[key].Distinct().ToList();
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
    }

    // *** FIX: Made Pathfinder and its nested classes public ***
    public static class Pathfinder
    {
        private const double METERS_TO_FEET = 3.28084;

        public class VirtualPathResult
        {
            public string RoutingSequence { get; set; } = "Could not determine virtual route.";
            public List<ElementId> StitchedPath { get; set; } = new List<ElementId>();
            public double VirtualLength { get; set; } = 0.0;
        }

        #region Public Pathfinding Methods

        public static List<ElementId> FindConfirmedPath(List<Element> startCandidates, List<Element> endCandidates, List<Element> cableElements, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            List<ElementId> bestPath = null;
            double minLength = double.MaxValue;

            foreach (var startCandidate in startCandidates)
            {
                foreach (var endCandidate in endCandidates)
                {
                    XYZ startPt = GetAccurateElementLocation(startCandidate);
                    XYZ endPt = GetAccurateElementLocation(endCandidate);
                    if (startPt == null || endPt == null) continue;

                    Element startElem = cableElements.OrderBy(e => GetClosestPointOnElement(e, startPt).DistanceTo(startPt)).FirstOrDefault();
                    Element endElem = cableElements.OrderBy(e => GetClosestPointOnElement(e, endPt).DistanceTo(endPt)).FirstOrDefault();

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
                var rtsIdGuid = SharedParameters.General.RTS_ID;
                return new VirtualPathResult { RoutingSequence = ReportFormatter.FormatDisconnectedPaths(disconnectedPaths, rtsIdToElementMap, doc, rtsIdGuid) };
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

        public enum ContainmentType { Tray, Conduit, Mixed }

        public class ContainmentIsland
        {
            public int Id { get; }
            public HashSet<ElementId> ElementIds { get; } = new HashSet<ElementId>();
            public BoundingBoxXYZ BoundingBox { get; set; }
            public ContainmentType Type { get; set; } = ContainmentType.Mixed;
            public ContainmentIsland(int id) { Id = id; }
        }

        public static List<ContainmentIsland> GroupIntoIslands(List<Element> cableElements, ContainmentGraph graph, Document doc)
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

        public static ContainmentType GetIslandContainmentType(ContainmentIsland island, Document doc)
        {
            bool hasTray = false;
            bool hasConduit = false;
            foreach (var elemId in island.ElementIds)
            {
                var elem = doc.GetElement(elemId);
                if (elem == null) continue;

#if REVIT2024_OR_GREATER
                long catId = elem.Category.Id.Value;
#else
                long catId = elem.Category.Id.IntegerValue;
#endif

                if (catId == (long)BuiltInCategory.OST_CableTray || catId == (long)BuiltInCategory.OST_CableTrayFitting) hasTray = true;
                else if (catId == (long)BuiltInCategory.OST_Conduit || catId == (long)BuiltInCategory.OST_ConduitFitting) hasConduit = true;
                if (hasTray && hasConduit) return ContainmentType.Mixed;
            }
            if (hasTray) return ContainmentType.Tray;
            if (hasConduit) return ContainmentType.Conduit;
            return ContainmentType.Mixed;
        }

        public static ContainmentIsland FindClosestIsland(List<Element> equipmentCandidates, List<ContainmentIsland> islands, Document doc)
        {
            if (equipmentCandidates == null || !equipmentCandidates.Any()) return null;

            ContainmentIsland bestIsland = null;
            double minDistance = double.MaxValue;

            foreach (var equip in equipmentCandidates)
            {
                var equipLocation = GetAccurateElementLocation(equip);
                if (equipLocation == null) continue;

                foreach (var island in islands)
                {
                    foreach (var elementId in island.ElementIds)
                    {
                        var elem = doc.GetElement(elementId);
                        var elemLocation = GetClosestPointOnElement(elem, equipLocation);
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

        public static Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> BuildIslandGraph(List<ContainmentIsland> islands, Document doc)
        {
            var islandGraph = new Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>>();
            for (int i = 0; i < islands.Count; i++)
            {
                islandGraph[islands[i]] = new List<Tuple<ContainmentIsland, double>>();
                var islandA_Endpoints = GetIslandEndpoints(islands[i], doc);

                for (int j = i + 1; j < islands.Count; j++)
                {
                    if (!islandGraph.ContainsKey(islands[j])) islandGraph[islands[j]] = new List<Tuple<ContainmentIsland, double>>();

                    var islandB_Endpoints = GetIslandEndpoints(islands[j], doc);
                    double minDistance = double.MaxValue;

                    foreach (var p1 in islandA_Endpoints)
                    {
                        foreach (var p2 in islandB_Endpoints)
                        {
                            minDistance = Math.Min(minDistance, p1.DistanceTo(p2));
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

        public static List<ContainmentIsland> FindShortestIslandPath(ContainmentIsland startIsland, ContainmentIsland endIsland, Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> islandGraph)
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

        public static VirtualPathResult StitchPaths(List<ContainmentIsland> islandPath, List<Element> startCandidates, List<Element> endCandidates, Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> islandGraph, ContainmentGraph graph, Dictionary<string, Element> rtsIdToElementMap, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            var finalSequence = new StringBuilder();
            var finalStitchedPath = new List<ElementId>();
            ElementId lastElementId = null;
            var rtsIdGuid = SharedParameters.General.RTS_ID;
            double totalVirtualLength = 0.0;

            if (startCandidates.Any())
            {
                var startPanel = startCandidates.First();
                lastElementId = FindClosestElementInIsland(startPanel, islandPath.First(), doc);
                double jumpLength = CalculateManhattanDistance(GetAccurateElementLocation(startPanel), GetClosestPointOnElement(doc.GetElement(lastElementId), GetAccurateElementLocation(startPanel)));
                totalVirtualLength += jumpLength;
                string panelName = startPanel.get_Parameter(rtsIdGuid)?.AsString() ?? startPanel.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "Start";
                finalSequence.Append(panelName).Append(" >> ");
            }

            for (int i = 0; i < islandPath.Count; i++)
            {
                var currentIsland = islandPath[i];
                ElementId startOfSegment;

                if (i > 0)
                {
                    var prevIsland = islandPath[i - 1];
                    var jumpTuple = islandGraph[prevIsland].FirstOrDefault(t => t.Item1.Id == currentIsland.Id);
                    double jumpDistance = jumpTuple?.Item2 ?? 0.0;
                    totalVirtualLength += jumpDistance;

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

                ElementId endOfSegment;
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
                    finalSequence.Append(ReportFormatter.FormatPath(segmentPath, rtsIdToElementMap, doc, rtsIdGuid));
                    finalStitchedPath.AddRange(segmentPath);
                    lastElementId = segmentPath.Last();
                }
                else
                {
                    finalSequence.Append("[PATH_NOT_FOUND_IN_ISLAND]");
                    var fallbackPath = new List<ElementId> { startOfSegment };
                    finalStitchedPath.AddRange(fallbackPath);
                    lastElementId = startOfSegment;
                }
            }

            if (endCandidates.Any())
            {
                var endPanel = endCandidates.First();
                double jumpLength = CalculateManhattanDistance(GetAccurateElementLocation(doc.GetElement(lastElementId)), GetAccurateElementLocation(endPanel));
                totalVirtualLength += jumpLength;
                string panelName = endPanel.get_Parameter(rtsIdGuid)?.AsString() ?? endPanel.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString() ?? "End";
                finalSequence.Append(" >> ").Append(panelName);
            }

            return new VirtualPathResult { RoutingSequence = finalSequence.ToString(), StitchedPath = finalStitchedPath, VirtualLength = totalVirtualLength };
        }

        public static ElementId FindClosestElementInIsland(Element targetElement, ContainmentIsland island, Document doc)
        {
            var targetLocation = GetAccurateElementLocation(targetElement);
            if (targetLocation == null) return island.ElementIds.FirstOrDefault();
            return island.ElementIds.OrderBy(id => GetClosestPointOnElement(doc.GetElement(id), targetLocation).DistanceTo(targetLocation)).FirstOrDefault();
        }

        public static ElementId FindFarthestElementInIsland(Element targetElement, ContainmentIsland island, Document doc)
        {
            var targetLocation = GetAccurateElementLocation(targetElement);
            if (targetLocation == null) return island.ElementIds.FirstOrDefault();
            return island.ElementIds.OrderByDescending(id => GetClosestPointOnElement(doc.GetElement(id), targetLocation).DistanceTo(targetLocation)).FirstOrDefault();
        }

        public static ElementId FindClosestElementToOtherIsland(ContainmentIsland fromIsland, ContainmentIsland toIsland, Document doc)
        {
            ElementId bestElementId = fromIsland.ElementIds.FirstOrDefault();
            double minDistance = double.MaxValue;

            var fromEndpoints = GetIslandEndpoints(fromIsland, doc);
            var toEndpoints = GetIslandEndpoints(toIsland, doc);

            foreach (var p1 in fromEndpoints)
            {
                foreach (var p2 in toEndpoints)
                {
                    double dist = p1.DistanceTo(p2);
                    if (dist < minDistance)
                    {
                        minDistance = dist;
                        bestElementId = fromIsland.ElementIds.OrderBy(id => GetClosestPointOnElement(doc.GetElement(id), p1).DistanceTo(p1)).First();
                    }
                }
            }
            return bestElementId;
        }

        #endregion

        #region Core Dijkstra Algorithm

        private static List<ElementId> FindShortestPath_Dijkstra(ElementId startNodeId, ElementId endNodeId, ContainmentGraph graph, Dictionary<ElementId, double> elementIdToLengthMap)
        {
            if (startNodeId == null || endNodeId == null || startNodeId.Equals(ElementId.InvalidElementId) || endNodeId.Equals(ElementId.InvalidElementId)) return null;
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
                    double costFromCurrent = elementIdToLengthMap.ContainsKey(currentNodeId) ? elementIdToLengthMap[currentNodeId] : 0.0;
                    double altDistance = distances[currentNodeId] + costFromCurrent;
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

        #region Location Helpers

        private static XYZ GetAccurateElementLocation(Element element)
        {
            if (element == null || element.Location == null) return null;

            try
            {
                if (element.Location is LocationPoint locPoint) return locPoint.Point;
                if (element.Location is LocationCurve locCurve) return locCurve.Curve.Evaluate(0.5, true);
            }
            catch { }

            var bbox = element.get_BoundingBox(null);

#if REVIT2024_OR_GREATER
            return bbox != null ? (bbox.Min + bbox.Max) / 2.0 : null;
#else
            return bbox != null ? (bbox.Min + bbox.Max) / 2.0 : null;
#endif
        }

        private static XYZ GetClosestPointOnElement(Element element, XYZ targetPoint)
        {
            if (element == null || targetPoint == null) return targetPoint ?? XYZ.Zero;

            if (element.Location is LocationCurve locCurve)
            {
                return locCurve.Curve.Project(targetPoint).XYZPoint;
            }

            return GetAccurateElementLocation(element) ?? targetPoint;
        }

        private static double CalculateManhattanDistance(XYZ p1, XYZ p2)
        {
            if (p1 == null || p2 == null) return 0.0;
            return Math.Abs(p1.X - p2.X) + Math.Abs(p1.Y - p2.Y) + Math.Abs(p1.Z - p2.Z);
        }

        public static List<XYZ> GetIslandEndpoints(ContainmentIsland island, Document doc)
        {
            var endpoints = new List<XYZ>();
            foreach (var elemId in island.ElementIds)
            {
                var elem = doc.GetElement(elemId);
                if (elem is MEPCurve curve && curve.Location is LocationCurve lc)
                {
                    endpoints.Add(lc.Curve.GetEndPoint(0));
                    endpoints.Add(lc.Curve.GetEndPoint(1));
                }
                else if (elem is FamilyInstance fitting)
                {
                    var connectors = ContainmentGraph.GetConnectors(fitting);
                    if (connectors != null)
                    {
                        foreach (Connector c in connectors)
                        {
                            endpoints.Add(c.Origin);
                        }
                    }
                }
            }
            return endpoints;
        }

        #endregion

        #region Nested PriorityQueue for .NET 4.8

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
            sb.AppendLine("Cable Reference,From,To,From Status,To Status,Status,Virtual Length (m),Branch Sequencing,Routing Sequence");
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

        public static string FormatPath(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap, Document doc, Guid rtsIdGuid)
        {
            if (path == null || !path.Any()) return "Path not found";

            var sequence = new List<string>();
            var rtsIdCache = new Dictionary<ElementId, string>();

            foreach (var kvp in rtsIdToElementMap)
            {
                if (!rtsIdCache.ContainsKey(kvp.Value.Id))
                {
                    rtsIdCache.Add(kvp.Value.Id, kvp.Key);
                }
            }

            for (int i = 0; i < path.Count; i++)
            {
                var currentId = path[i];
                var currentElem = doc.GetElement(currentId);
                if (currentElem == null) continue;

                string rtsId;
                if (!rtsIdCache.TryGetValue(currentId, out rtsId))
                {
                    rtsId = currentElem.get_Parameter(rtsIdGuid)?.AsString();
                    rtsIdCache[currentId] = rtsId;
                }

                if (string.IsNullOrEmpty(rtsId))
                {
                    rtsId = "[MISSING_RTS_ID]";
                }

                if (i > 0)
                {
                    var prevElem = doc.GetElement(path[i - 1]);
                    if (prevElem != null && currentElem.Category.Id != prevElem.Category.Id)
                    {
                        sequence.Add(" + ");
                    }
                    else
                    {
                        sequence.Add(", ");
                    }
                }
                sequence.Add(rtsId);
            }
            return string.Join("", sequence);
        }

        public static string FormatDisconnectedPaths(List<List<ElementId>> allPaths, Dictionary<string, Element> rtsIdMap, Document doc, Guid rtsIdGuid)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < allPaths.Count; i++)
            {
                if (i > 0) sb.Append(" >> ");
                sb.Append(FormatPath(allPaths[i], rtsIdMap, doc, rtsIdGuid));
            }
            return sb.ToString();
        }

        public static string FormatBranchSequence(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap, Guid rtsIdGuid)
        {
            if (path == null || !path.Any()) return "N/A";
            var uniqueBranches = new List<string>();
            var seenBranches = new HashSet<string>();
            bool hasAddedLeadingZeroBranch = false;

            foreach (var elementId in path)
            {
                var elem = rtsIdToElementMap.Values.FirstOrDefault(e => e.Id == elementId);
                string rtsId = elem?.get_Parameter(rtsIdGuid)?.AsString();

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
