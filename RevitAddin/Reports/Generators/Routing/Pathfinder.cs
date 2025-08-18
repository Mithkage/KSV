//-----------------------------------------------------------------------------
// <copyright file="Pathfinder.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Handles all pathfinding logic, including Dijkstra's algorithm and virtual island hopping.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Made helper classes (Pathfinder, ContainmentIsland, etc.) public to allow access from diagnostic tools.
// - [APPLIED FIX]: Enhanced FindConfirmedPath to be "containment-type aware", allowing it to correctly route across different but disconnected containment systems (e.g., tray to conduit).
// - [APPLIED FIX]: Corrected island pathfinding to handle routes contained within a single island.
// - [APPLIED FIX]: Added detailed status messages for virtual pathfinding failures.
// - [APPLIED FIX]: Corrected BuildIslandGraph to ensure a fully connected graph is created.
// - [APPLIED FIX]: Replaced island distance calculation from center-to-center to edge-to-edge for more accurate virtual pathing.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using RTS.Reports.Utils;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Handles all pathfinding logic, including Dijkstra's algorithm and virtual island hopping.
    /// </summary>
    public static class Pathfinder
    {
        private const double METERS_TO_FEET = 3.28084;

        #region Public Pathfinding Methods

        /// <summary>
        /// Finds the best confirmed path between start and end equipment, considering routes
        /// that may span different, disconnected containment types (e.g., tray and conduit).
        /// </summary>
        /// <returns>The list of ElementIds representing the shortest confirmed path, or null if no path is found.</returns>
        public static List<ElementId> FindConfirmedPath(List<Element> startCandidates, List<Element> endCandidates, List<Element> cableElements, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            if (!startCandidates.Any() || !endCandidates.Any() || !cableElements.Any())
            {
                return null;
            }

            // Separate the assigned containment elements by their system type (Tray vs. Conduit).
            var trayElements = cableElements.Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray || e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTrayFitting).ToList();
            var conduitElements = cableElements.Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit || e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_ConduitFitting).ToList();

            var possiblePaths = new List<List<ElementId>>();

            // Scenario 1: Path is entirely within the tray network.
            if (trayElements.Any())
            {
                var path = FindShortestPathInSystem(startCandidates, endCandidates, trayElements, graph, lengthMap, doc);
                if (path != null) possiblePaths.Add(path);
            }

            // Scenario 2: Path is entirely within the conduit network.
            if (conduitElements.Any())
            {
                var path = FindShortestPathInSystem(startCandidates, endCandidates, conduitElements, graph, lengthMap, doc);
                if (path != null) possiblePaths.Add(path);
            }

            // Scenario 3: Path is a hybrid, starting in tray and ending in conduit.
            if (trayElements.Any() && conduitElements.Any())
            {
                var path = FindHybridPath(startCandidates, endCandidates, trayElements, conduitElements, graph, lengthMap, doc);
                if (path != null) possiblePaths.Add(path);
            }

            // Scenario 4: Path is a hybrid, starting in conduit and ending in tray.
            if (conduitElements.Any() && trayElements.Any())
            {
                var path = FindHybridPath(startCandidates, endCandidates, conduitElements, trayElements, graph, lengthMap, doc);
                if (path != null) possiblePaths.Add(path);
            }

            // If no paths were found, return null.
            if (!possiblePaths.Any()) return null;

            // Compare all found paths and return the one with the minimum physical length.
            return possiblePaths.OrderBy(p => GetPathLength(p, lengthMap)).FirstOrDefault();
        }


        public static VirtualPathResult FindBestDisconnectedSequence(List<Element> startCandidates, List<Element> endCandidates, List<Element> cableElements, ContainmentGraph graph, Dictionary<string, Element> rtsIdToElementMap, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            if (!cableElements.Any()) return new VirtualPathResult();

            var islands = GroupIntoIslands(cableElements, graph, doc);
            if (!islands.Any()) return new VirtualPathResult { StatusMessage = "No containment islands could be formed from the assigned elements." };

            var elementIdToIslandMap = islands.SelectMany(i => i.ElementIds.Select(id => new { ElementId = id, Island = i }))
                                              .ToDictionary(x => x.ElementId, x => x.Island);

            ContainmentIsland startIsland = FindClosestIsland(startCandidates, islands, doc);
            ContainmentIsland endIsland = FindClosestIsland(endCandidates, islands, doc);

            if (startIsland == null && endIsland == null)
            {
                var disconnectedPaths = islands.Select(i => i.ElementIds.ToList()).ToList();
                return new VirtualPathResult
                {
                    RoutingSequence = ReportFormatter.FormatDisconnectedPaths(disconnectedPaths, rtsIdToElementMap, doc, SharedParameters.General.RTS_ID),
                    IslandCount = islands.Count,
                    StatusMessage = startCandidates.Any() || endCandidates.Any() ? "Could not link start/end equipment to any containment island." : null
                };
            }

            startIsland = startIsland ?? endIsland;
            endIsland = endIsland ?? startIsland;

            var islandGraph = BuildIslandGraph(islands, doc);
            var islandPath = FindShortestIslandPath(startIsland, endIsland, islandGraph);

            if (islandPath == null || !islandPath.Any())
            {
                string statusMsg = $"Could not find a path between start island (ID: {startIsland.Id}) and end island (ID: {endIsland.Id}).";
                if (startIsland.Id == endIsland.Id)
                {
                    statusMsg = $"Start and end panels are closest to the same island (ID: {startIsland.Id}), but an internal path could not be found.";
                }
                return new VirtualPathResult { IslandCount = islands.Count, StatusMessage = statusMsg };
            }

            var result = StitchPaths(islandPath, startCandidates, endCandidates, islandGraph, graph, rtsIdToElementMap, lengthMap, doc);
            result.IslandCount = islands.Count;
            return result;
        }

        #endregion

        #region Private Pathfinding Helpers

        /// <summary>
        /// Finds the shortest path for a route that is contained entirely within a single containment system (e.g., all tray or all conduit).
        /// </summary>
        private static List<ElementId> FindShortestPathInSystem(List<Element> startCandidates, List<Element> endCandidates, List<Element> systemElements, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            Element startElem = FindClosestElementToCandidates(startCandidates, systemElements, doc);
            Element endElem = FindClosestElementToCandidates(endCandidates, systemElements, doc);

            if (startElem == null || endElem == null) return null;

            return FindShortestPath_Dijkstra(startElem.Id, endElem.Id, graph, lengthMap);
        }

        /// <summary>
        /// Finds the shortest path for a hybrid route that transitions from one containment system to another.
        /// </summary>
        private static List<ElementId> FindHybridPath(List<Element> startCandidates, List<Element> endCandidates, List<Element> startSystemElements, List<Element> endSystemElements, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap, Document doc)
        {
            // Find the best entry point into the first system.
            Element entryPoint = FindClosestElementToCandidates(startCandidates, startSystemElements, doc);
            if (entryPoint == null) return null;

            // Find the best exit point from the second system.
            Element exitPoint = FindClosestElementToCandidates(endCandidates, endSystemElements, doc);
            if (exitPoint == null) return null;

            // Find the element in the first system that is closest to the second system's network.
            Element transitionStart = FindClosestElementToOtherSystem(startSystemElements, endSystemElements, doc);
            if (transitionStart == null) return null;

            // Find the element in the second system that is closest to the first system's network.
            Element transitionEnd = FindClosestElementToOtherSystem(endSystemElements, startSystemElements, doc);
            if (transitionEnd == null) return null;

            // Calculate the path for the first segment.
            var path1 = FindShortestPath_Dijkstra(entryPoint.Id, transitionStart.Id, graph, lengthMap);
            if (path1 == null) return null;

            // Calculate the path for the second segment.
            var path2 = FindShortestPath_Dijkstra(transitionEnd.Id, exitPoint.Id, graph, lengthMap);
            if (path2 == null) return null;

            // Combine the two path segments.
            path1.AddRange(path2);
            return path1;
        }

        /// <summary>
        /// Finds the single element in a list that is closest to any of the candidate equipment.
        /// </summary>
        private static Element FindClosestElementToCandidates(List<Element> equipmentCandidates, List<Element> containmentElements, Document doc)
        {
            return containmentElements
                .OrderBy(containment => equipmentCandidates
                    .Min(equip => ReportHelpers.GetElementLocation(containment)?.DistanceTo(ReportHelpers.GetElementLocation(equip)) ?? double.MaxValue))
                .FirstOrDefault();
        }

        /// <summary>
        /// Finds the element in one system that is physically closest to any element in another system.
        /// </summary>
        private static Element FindClosestElementToOtherSystem(List<Element> systemA, List<Element> systemB, Document doc)
        {
            return systemA
                .OrderBy(elemA => systemB
                    .Min(elemB => ReportHelpers.GetElementLocation(elemA)?.DistanceTo(ReportHelpers.GetElementLocation(elemB)) ?? double.MaxValue))
                .FirstOrDefault();
        }


        #endregion

        #region Island Hopping Logic

        public static List<ContainmentIsland> GroupIntoIslands(List<Element> cableElements, ContainmentGraph graph, Document doc)
        {
            var islands = new List<ContainmentIsland>();
            var visited = new HashSet<ElementId>();
            int islandIdCounter = 0;

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
                            // Only expand to neighbors that are part of the assigned cable elements.
                            if (cableElements.Any(e => e.Id == neighborId) && !visited.Contains(neighborId))
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

                int catId = elem.Category.Id.IntegerValue;
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
            ContainmentIsland closestIsland = null;
            double minDistance = double.MaxValue;

            foreach (var equip in equipmentCandidates)
            {
                var equipLocation = ReportHelpers.GetElementLocation(equip);
                if (equipLocation == null) continue;
                foreach (var island in islands)
                {
                    if (island.BoundingBox == null || !island.BoundingBox.Enabled) continue;
                    var boxCenter = (island.BoundingBox.Min + island.BoundingBox.Max) / 2.0;
                    double dist = equipLocation.DistanceTo(boxCenter);
                    if (dist < minDistance)
                    {
                        minDistance = dist;
                        closestIsland = island;
                    }
                }
            }
            return closestIsland;
        }

        // [CHANGE]: The method for calculating the distance between islands was changed from a simple
        // center-to-center calculation to a more accurate edge-to-edge calculation.
        // [REASON]: The center-to-center method was inaccurate for large or irregularly shaped islands, causing the
        // pathfinder to choose illogical, long virtual jumps instead of shorter, geometrically obvious paths.
        // [NEW METHOD]: This method now uses the GetMinDistanceBetweenIslands helper to find the
        // true shortest distance between the bounding boxes of two islands.
        private static Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> BuildIslandGraph(List<ContainmentIsland> islands, Document doc)
        {
            var islandGraph = new Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>>();

            // First, initialize the graph with all islands as keys.
            foreach (var island in islands)
            {
                if (!islandGraph.ContainsKey(island))
                {
                    islandGraph[island] = new List<Tuple<ContainmentIsland, double>>();
                }
            }

            // Then, populate the adjacency lists with edge-to-edge distances.
            for (int i = 0; i < islands.Count; i++)
            {
                for (int j = i + 1; j < islands.Count; j++)
                {
                    var islandA = islands[i];
                    var islandB = islands[j];

                    double distance = GetMinDistanceBetweenIslands(islandA, islandB, doc);

                    islandGraph[islandA].Add(new Tuple<ContainmentIsland, double>(islandB, distance));
                    islandGraph[islandB].Add(new Tuple<ContainmentIsland, double>(islandA, distance));
                }
            }
            return islandGraph;
        }

        /// <summary>
        /// Calculates the minimum distance between the edges of two island bounding boxes. This provides a more
        /// realistic "jump" cost for the pathfinder than a simple center-to-center calculation.
        /// </summary>
        private static double GetMinDistanceBetweenIslands(ContainmentIsland islandA, ContainmentIsland islandB, Document doc)
        {
            var boxA = islandA.BoundingBox;
            var boxB = islandB.BoundingBox;

            if (boxA == null || !boxA.Enabled || boxB == null || !boxB.Enabled) return double.MaxValue;

            // Find the closest point on boxB to the center of boxA
            var centerA = (boxA.Min + boxA.Max) / 2.0;
            var closestPointOnB = GetClosestPointOnBoundingBox(boxB, centerA);

            // Find the closest point on boxA to that point on boxB
            var closestPointOnA = GetClosestPointOnBoundingBox(boxA, closestPointOnB);

            return closestPointOnA.DistanceTo(closestPointOnB);
        }

        /// <summary>
        /// Finds the point on a bounding box that is closest to a given target point.
        /// </summary>
        private static XYZ GetClosestPointOnBoundingBox(BoundingBoxXYZ box, XYZ target)
        {
            double x = Math.Max(box.Min.X, Math.Min(target.X, box.Max.X));
            double y = Math.Max(box.Min.Y, Math.Min(target.Y, box.Max.Y));
            double z = Math.Max(box.Min.Z, Math.Min(target.Z, box.Max.Z));
            return new XYZ(x, y, z);
        }

        private static List<ContainmentIsland> FindShortestIslandPath(ContainmentIsland startIsland, ContainmentIsland endIsland, Dictionary<ContainmentIsland, List<Tuple<ContainmentIsland, double>>> islandGraph)
        {
            if (startIsland.Id == endIsland.Id)
            {
                return new List<ContainmentIsland> { startIsland };
            }

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
            double totalVirtualLength = 0;
            ElementId lastElementId = null;

            // Determine the absolute start element for the entire sequence.
            if (startCandidates.Any())
            {
                lastElementId = FindClosestElementInIsland(startCandidates.First(), islandPath.First(), doc);
                finalSequence.Append(ReportFormatter.FormatEquipment(startCandidates.First(), doc, SharedParameters.General.RTS_ID));
            }

            for (int i = 0; i < islandPath.Count; i++)
            {
                var currentIsland = islandPath[i];
                ElementId startOfSegment;
                ElementId endOfSegment;

                if (i > 0) // This is not the first island, so we need a separator.
                {
                    var prevIsland = islandPath[i - 1];
                    var jumpTuple = islandGraph[prevIsland].FirstOrDefault(t => t.Item1.Id == currentIsland.Id);
                    double jumpDistance = jumpTuple?.Item2 ?? 0.0;
                    totalVirtualLength += jumpDistance;

                    // Add appropriate separator based on containment types and distance.
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
                    // This is the first island.
                    startOfSegment = lastElementId; // It was set by the start candidate logic above.
                    if (startCandidates.Any()) finalSequence.Append(" >> ");
                }

                // Determine the end of the current island segment.
                if (i == islandPath.Count - 1) // This is the last island.
                {
                    endOfSegment = endCandidates.Any() ? FindClosestElementInIsland(endCandidates.First(), currentIsland, doc) : FindFarthestElementInIsland(doc.GetElement(startOfSegment), currentIsland, doc);
                }
                else // There are more islands to come.
                {
                    var nextIsland = islandPath[i + 1];
                    endOfSegment = FindClosestElementToOtherIsland(currentIsland, nextIsland, doc);
                }

                // Handle cases where start/end points might be null or need a default.
                if (startOfSegment == null || startOfSegment.IntegerValue == -1)
                {
                    startOfSegment = currentIsland.ElementIds.FirstOrDefault();
                }
                if (endOfSegment == null || endOfSegment.IntegerValue == -1)
                {
                    endOfSegment = startOfSegment;
                }


                var segmentPath = FindShortestPath_Dijkstra(startOfSegment, endOfSegment, graph, lengthMap);
                if (segmentPath != null && segmentPath.Any())
                {
                    finalSequence.Append(ReportFormatter.FormatPath(segmentPath, rtsIdToElementMap, doc, SharedParameters.General.RTS_ID));
                    finalStitchedPath.AddRange(segmentPath);
                    lastElementId = segmentPath.Last();
                }
                else // Fallback if Dijkstra fails for a segment.
                {
                    var fallbackPath = new List<ElementId> { startOfSegment };
                    if (startOfSegment != endOfSegment) fallbackPath.Add(endOfSegment);
                    finalSequence.Append(ReportFormatter.FormatPath(fallbackPath, rtsIdToElementMap, doc, SharedParameters.General.RTS_ID));
                    finalStitchedPath.AddRange(fallbackPath);
                    lastElementId = endOfSegment;
                }
            }

            // Append the absolute end element for the entire sequence.
            if (endCandidates.Any())
            {
                finalSequence.Append(" >> ").Append(ReportFormatter.FormatEquipment(endCandidates.First(), doc, SharedParameters.General.RTS_ID));
            }

            return new VirtualPathResult { RoutingSequence = finalSequence.ToString(), StitchedPath = finalStitchedPath, VirtualLength = totalVirtualLength };
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
            ElementId bestElement = fromIsland.ElementIds.FirstOrDefault();
            double minDistance = double.MaxValue;
            var toBox = toIsland.BoundingBox;
            if (toBox == null || !toBox.Enabled) return bestElement;
            var toCenter = (toBox.Min + toBox.Max) / 2.0;

            foreach (var elementId in fromIsland.ElementIds)
            {
                var fromLocation = ReportHelpers.GetElementLocation(doc.GetElement(elementId));
                if (fromLocation == null) continue;
                double dist = fromLocation.DistanceTo(toCenter);
                if (dist < minDistance)
                {
                    minDistance = dist;
                    bestElement = elementId;
                }
            }
            return bestElement;
        }

        #endregion

        #region Core Dijkstra Algorithm

        private static List<ElementId> FindShortestPath_Dijkstra(ElementId startNodeId, ElementId endNodeId, ContainmentGraph graph, Dictionary<ElementId, double> elementIdToLengthMap)
        {
            if (startNodeId == null || endNodeId == null || startNodeId.IntegerValue == -1 || endNodeId.IntegerValue == -1) return null;
            if (startNodeId == endNodeId) return new List<ElementId> { startNodeId };

            var distances = new Dictionary<ElementId, double>();
            var previous = new Dictionary<ElementId, ElementId>();
            var queue = new PriorityQueue<ElementId>();

            foreach (var nodeId in graph.AdjacencyGraph.Keys)
            {
                distances[nodeId] = double.PositiveInfinity;
            }

            if (!distances.ContainsKey(startNodeId)) return null; // Start node not in graph
            distances[startNodeId] = 0;
            queue.Enqueue(startNodeId, 0);

            while (queue.Count > 0)
            {
                var currentNodeId = queue.Dequeue();
                if (currentNodeId == endNodeId)
                {
                    var path = new List<ElementId>();
                    var at = endNodeId;
                    while (at != null && at.IntegerValue != -1)
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

                    if (!distances.ContainsKey(neighborId)) distances[neighborId] = double.PositiveInfinity;

                    if (altDistance < distances[neighborId])
                    {
                        distances[neighborId] = altDistance;
                        previous[neighborId] = currentNodeId;
                        queue.Enqueue(neighborId, altDistance);
                    }
                }
            }
            return null; // Path not found
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
}