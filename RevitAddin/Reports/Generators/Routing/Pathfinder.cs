//-----------------------------------------------------------------------------
// <copyright file="Pathfinder.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Provides pathfinding capabilities for routing cables through containment.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Corrected FindEntryPoint to handle cases where no valid entry point is found.
// - [APPLIED FIX]: Added high-accuracy node-based search for short routes.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Provides pathfinding capabilities for routing cables through containment.
    /// </summary>
    public static class Pathfinder
    {
        /// <summary>
        /// Finds a continuous, confirmed path between start and end equipment within a single containment network.
        /// </summary>
        public static List<ElementId> FindConfirmedPath(
            List<Element> startCandidates, List<Element> endCandidates,
            List<Element> cableContainment, ContainmentGraph graph,
            Dictionary<ElementId, double> lengthMap, Document doc, bool useHighAccuracy)
        {
            if (!startCandidates.Any() || !endCandidates.Any() || !cableContainment.Any())
                return null;

            var startPoints = FindEntryPoints(startCandidates, cableContainment, doc, useHighAccuracy);
            var endPoints = FindEntryPoints(endCandidates, cableContainment, doc, useHighAccuracy);

            if (!startPoints.Any() || !endPoints.Any())
                return null;

            // Find the best path among all combinations of start and end points.
            List<ElementId> bestPath = null;
            double shortestLength = double.MaxValue;

            foreach (var start in startPoints)
            {
                foreach (var end in endPoints)
                {
                    var path = Dijkstra(start.Item1, end.Item1, graph, lengthMap);
                    if (path != null)
                    {
                        double currentLength = path.Sum(id => lengthMap.ContainsKey(id) ? lengthMap[id] : 0.0);
                        if (currentLength < shortestLength)
                        {
                            shortestLength = currentLength;
                            bestPath = path;
                        }
                    }
                }
            }
            return bestPath;
        }

        /// <summary>
        /// Finds the optimal sequence of disconnected containment islands to route a cable.
        /// </summary>
        public static VirtualPathResult FindBestDisconnectedSequence(
            List<Element> startCandidates, List<Element> endCandidates,
            List<Element> cableContainment, ContainmentGraph graph,
            Dictionary<string, Element> rtsIdMap, Dictionary<ElementId, double> lengthMap, Document doc, bool useHighAccuracy)
        {
            var result = new VirtualPathResult();
            var islands = GroupIntoIslands(cableContainment, graph, doc);
            result.IslandCount = islands.Count;

            if (!islands.Any())
            {
                result.StatusMessage = "No containment islands could be identified for the assigned elements.";
                return result;
            }

            // Find entry/exit points for each island
            foreach (var island in islands)
            {
                island.EntryPoints = FindEntryPoints(startCandidates, island.Elements, doc, useHighAccuracy);
                island.ExitPoints = FindEntryPoints(endCandidates, island.Elements, doc, useHighAccuracy);
            }

            // Filter islands that have valid connections to start/end equipment
            var reachableIslands = islands.Where(i => i.EntryPoints.Any()).ToList();
            if (!reachableIslands.Any())
            {
                result.StatusMessage = "Could not find a valid entry point from the 'From' equipment to any assigned containment.";
                return result;
            }

            // Find the best starting island (closest to the start equipment)
            var bestStartIsland = reachableIslands
                .OrderBy(i => i.EntryPoints.Min(p => p.Item2))
                .FirstOrDefault();

            if (bestStartIsland == null)
            {
                result.StatusMessage = "Failed to determine the best starting containment island.";
                return result;
            }

            // Build the path by hopping between the ordered islands
            var (stitchedPath, sequence, virtualLength) = BuildStitchedPath(bestStartIsland, islands, startCandidates, endCandidates, graph, lengthMap, rtsIdMap, doc, useHighAccuracy);

            result.StitchedPath = stitchedPath;
            result.RoutingSequence = sequence;
            result.VirtualLength = virtualLength;

            return result;
        }

        /// <summary>
        /// Groups containment elements into connected sub-graphs (islands).
        /// </summary>
        public static List<ContainmentIsland> GroupIntoIslands(List<Element> elements, ContainmentGraph graph, Document doc)
        {
            var islands = new List<ContainmentIsland>();
            var visited = new HashSet<ElementId>();

            foreach (var element in elements)
            {
                if (visited.Contains(element.Id)) continue;

                var islandElements = new List<Element>();
                var queue = new Queue<ElementId>();

                queue.Enqueue(element.Id);
                visited.Add(element.Id);

                while (queue.Count > 0)
                {
                    var currentId = queue.Dequeue();
                    var currentElement = doc.GetElement(currentId);
                    if (currentElement != null)
                    {
                        islandElements.Add(currentElement);
                    }

                    if (graph.AdjacencyGraph.TryGetValue(currentId, out var neighbors))
                    {
                        foreach (var neighborId in neighbors)
                        {
                            if (!visited.Contains(neighborId))
                            {
                                visited.Add(neighborId);
                                queue.Enqueue(neighborId);
                            }
                        }
                    }
                }
                islands.Add(new ContainmentIsland { Elements = islandElements });
            }
            return islands;
        }

        /// <summary>
        /// Finds the closest points on a set of containment elements to a piece of equipment.
        /// </summary>
        /// <returns>A list of tuples containing the ElementId of the containment and the distance.</returns>
        private static List<Tuple<ElementId, double>> FindEntryPoints(List<Element> equipmentList, List<Element> containment, Document doc, bool useHighAccuracy)
        {
            var entryPoints = new List<Tuple<ElementId, double>>();
            if (!equipmentList.Any() || !containment.Any()) return entryPoints;

            var equipment = equipmentList.First(); // Assume the first candidate is the target
            var equipmentLocation = (equipment.Location as LocationPoint)?.Point;
            if (equipmentLocation == null) return entryPoints;

            foreach (var elem in containment)
            {
                double minDistance = double.MaxValue;

                if (useHighAccuracy && elem is MEPCurve)
                {
                    // High-accuracy: check distance to the curve itself
                    var curve = (elem.Location as LocationCurve)?.Curve;
                    if (curve != null)
                    {
                        minDistance = curve.Distance(equipmentLocation);
                    }
                }
                else if (useHighAccuracy && elem is FamilyInstance fi)
                {
                    // High-accuracy: check distance to each connector
                    var connectors = ContainmentGraph.GetConnectors(fi);
                    if (connectors != null)
                    {
                        foreach (Connector c in connectors)
                        {
                            minDistance = Math.Min(minDistance, c.Origin.DistanceTo(equipmentLocation));
                        }
                    }
                }
                else
                {
                    // Standard accuracy: use bounding box
                    var bbox = elem.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        minDistance = GetDistanceToBoundingBox(equipmentLocation, bbox);
                    }
                }

                if (minDistance < double.MaxValue)
                {
                    entryPoints.Add(new Tuple<ElementId, double>(elem.Id, minDistance));
                }
            }

            // Return the top 3 closest containment elements as potential entry points
            return entryPoints.OrderBy(p => p.Item2).Take(3).ToList();
        }

        /// <summary>
        /// Builds the final stitched path and formatted sequence string by traversing islands.
        /// </summary>
        private static (List<ElementId>, string, double) BuildStitchedPath(
            ContainmentIsland startIsland, List<ContainmentIsland> allIslands,
            List<Element> startCandidates, List<Element> endCandidates,
            ContainmentGraph graph, Dictionary<ElementId, double> lengthMap,
            Dictionary<string, Element> rtsIdMap, Document doc, bool useHighAccuracy)
        {
            var stitchedPath = new List<ElementId>();
            var sequenceParts = new List<string>();
            double totalVirtualLength = 0;

            var currentIsland = startIsland;
            var remainingIslands = allIslands.Where(i => i != startIsland).ToList();

            // Add connection from start equipment to the first island
            var startEntryPoint = currentIsland.EntryPoints.First();
            totalVirtualLength += startEntryPoint.Item2;
            var fromPanelName = ReportFormatter.FormatEquipment(startCandidates.FirstOrDefault(), doc, rtsIdMap.Values.First().get_Parameter(SharedParameters.General.RTS_ID).GUID);
            sequenceParts.Add(fromPanelName);
            sequenceParts.Add(">>");

            while (currentIsland != null)
            {
                // Find the best path within the current island
                var islandExitPoint = FindBestExitPoint(currentIsland, endCandidates, remainingIslands, doc, useHighAccuracy);
                var internalPath = Dijkstra(startEntryPoint.Item1, islandExitPoint.Item1, graph, lengthMap);
                if (internalPath != null)
                {
                    stitchedPath.AddRange(internalPath);
                    sequenceParts.Add(ReportFormatter.FormatPath(internalPath, rtsIdMap, doc, rtsIdMap.Values.First().get_Parameter(SharedParameters.General.RTS_ID).GUID));
                }

                if (!remainingIslands.Any())
                {
                    // This is the last island, connect to the end equipment
                    totalVirtualLength += islandExitPoint.Item2;
                    break;
                }

                // Find the next closest island
                var (nextIsland, connection) = FindNextIsland(currentIsland, remainingIslands, doc);
                totalVirtualLength += connection.Item3; // Add jump distance
                sequenceParts.Add(">>");

                // Prepare for the next iteration
                currentIsland = nextIsland;
                remainingIslands.Remove(nextIsland);
                startEntryPoint = new Tuple<ElementId, double>(connection.Item2, 0); // The entry to the next island is the connected element
            }

            var toPanelName = ReportFormatter.FormatEquipment(endCandidates.FirstOrDefault(), doc, rtsIdMap.Values.First().get_Parameter(SharedParameters.General.RTS_ID).GUID);
            sequenceParts.Add(toPanelName);

            return (stitchedPath, string.Join(" ", sequenceParts), totalVirtualLength);
        }

        /// <summary>
        /// Finds the best exit point from an island, considering both the final destination and the next island hop.
        /// </summary>
        private static Tuple<ElementId, double> FindBestExitPoint(ContainmentIsland island, List<Element> endCandidates, List<ContainmentIsland> remainingIslands, Document doc, bool useHighAccuracy)
        {
            // If it's the last island, the best exit is the one closest to the end equipment.
            if (!remainingIslands.Any())
            {
                return island.ExitPoints.OrderBy(p => p.Item2).FirstOrDefault() ?? island.Elements.Select(e => new Tuple<ElementId, double>(e.Id, 0)).First();
            }

            // Otherwise, find the point that provides the shortest jump to the next island.
            Tuple<ElementId, ElementId, double> bestConnection = null;
            double minJump = double.MaxValue;

            foreach (var nextIsland in remainingIslands)
            {
                var (fromId, toId, dist) = FindShortestConnection(island, nextIsland, doc);
                if (dist < minJump)
                {
                    minJump = dist;
                    bestConnection = new Tuple<ElementId, ElementId, double>(fromId, toId, dist);
                }
            }

            return new Tuple<ElementId, double>(bestConnection.Item1, 0);
        }

        /// <summary>
        /// Finds the closest connection between two islands.
        /// </summary>
        /// <returns>A tuple containing the from ElementId, to ElementId, and the distance.</returns>
        private static (ElementId, ElementId, double) FindShortestConnection(ContainmentIsland fromIsland, ContainmentIsland toIsland, Document doc)
        {
            ElementId fromId = null, toId = null;
            double shortestDist = double.MaxValue;

            foreach (var fromElem in fromIsland.Elements)
            {
                var fromBbox = fromElem.get_BoundingBox(null);
                if (fromBbox == null) continue;

                foreach (var toElem in toIsland.Elements)
                {
                    var toBbox = toElem.get_BoundingBox(null);
                    if (toBbox == null) continue;

                    double dist = GetDistanceToBoundingBox(fromBbox.Min, toBbox); // Approximate distance
                    if (dist < shortestDist)
                    {
                        shortestDist = dist;
                        fromId = fromElem.Id;
                        toId = toElem.Id;
                    }
                }
            }
            return (fromId, toId, shortestDist);
        }

        /// <summary>
        /// Finds the next island in the sequence based on proximity.
        /// </summary>
        private static (ContainmentIsland, Tuple<ElementId, ElementId, double>) FindNextIsland(ContainmentIsland current, List<ContainmentIsland> candidates, Document doc)
        {
            ContainmentIsland nextIsland = null;
            Tuple<ElementId, ElementId, double> bestConnection = null;
            double minDistance = double.MaxValue;

            foreach (var candidate in candidates)
            {
                var (fromId, toId, dist) = FindShortestConnection(current, candidate, doc);
                if (dist < minDistance)
                {
                    minDistance = dist;
                    nextIsland = candidate;
                    bestConnection = new Tuple<ElementId, ElementId, double>(fromId, toId, dist);
                }
            }
            return (nextIsland, bestConnection);
        }

        /// <summary>
        /// Standard Dijkstra's algorithm implementation for finding the shortest path in the graph.
        /// </summary>
        private static List<ElementId> Dijkstra(ElementId startNode, ElementId endNode, ContainmentGraph graph, Dictionary<ElementId, double> lengthMap)
        {
            var distances = new Dictionary<ElementId, double>();
            var previous = new Dictionary<ElementId, ElementId>();
            var queue = new List<ElementId>();

            foreach (var vertex in graph.AdjacencyGraph.Keys)
            {
                distances[vertex] = double.MaxValue;
                previous[vertex] = null;
                queue.Add(vertex);
            }
            distances[startNode] = 0;

            while (queue.Count > 0)
            {
                queue.Sort((a, b) => distances[a].CompareTo(distances[b]));
                var u = queue[0];
                queue.RemoveAt(0);

                if (u == endNode)
                {
                    var path = new List<ElementId>();
                    while (previous[u] != null)
                    {
                        path.Insert(0, u);
                        u = previous[u];
                    }
                    path.Insert(0, startNode);
                    return path;
                }

                if (!graph.AdjacencyGraph.ContainsKey(u)) continue;

                foreach (var v in graph.AdjacencyGraph[u])
                {
                    double alt = distances[u] + (lengthMap.ContainsKey(u) ? lengthMap[u] : 0.0);
                    if (alt < distances[v])
                    {
                        distances[v] = alt;
                        previous[v] = u;
                    }
                }
            }
            return null; // Path not found
        }

        /// <summary>
        /// Calculates the minimum distance from a point to a bounding box.
        /// </summary>
        private static double GetDistanceToBoundingBox(XYZ point, BoundingBoxXYZ bbox)
        {
            double dx = Math.Max(0, Math.Max(bbox.Min.X - point.X, point.X - bbox.Max.X));
            double dy = Math.Max(0, Math.Max(bbox.Min.Y - point.Y, point.Y - bbox.Max.Y));
            double dz = Math.Max(0, Math.Max(bbox.Min.Z - point.Z, point.Z - bbox.Max.Z));
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
    }
}