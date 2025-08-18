using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Represents the containment network as an adjacency graph, where keys are ElementIds
    /// and values are lists of connected ElementIds.
    /// </summary>
    public class ContainmentGraph
    {
        public Dictionary<ElementId, List<ElementId>> AdjacencyGraph { get; }
        private readonly Document _doc;

        public ContainmentGraph(List<Element> elements, Document doc)
        {
            _doc = doc;
            AdjacencyGraph = BuildAdjacencyGraph(elements);
        }

        /// <summary>
        /// Builds the adjacency graph from a list of containment elements by inspecting their connectors.
        /// </summary>
        private Dictionary<ElementId, List<ElementId>> BuildAdjacencyGraph(List<Element> elements)
        {
            var graph = new Dictionary<ElementId, List<ElementId>>();
            var elementDict = elements.ToDictionary(e => e.Id);

            // Initialize the graph with all elements as keys.
            foreach (var elem in elements)
            {
                if (!graph.ContainsKey(elem.Id))
                {
                    graph[elem.Id] = new List<ElementId>();
                }
            }

            // Populate the adjacency lists by checking connectors.
            foreach (var elem in elements)
            {
                try
                {
                    var connectors = GetConnectors(elem);
                    if (connectors == null || connectors.IsEmpty) continue;

                    // Logic for fittings: connect all curves attached to the fitting to each other.
                    if (elem is FamilyInstance fitting)
                    {
                        var neighbors = new List<ElementId>();
                        foreach (Connector fittingConnector in connectors)
                        {
                            if (!fittingConnector.IsConnected) continue;
                            foreach (Connector connected in fittingConnector.AllRefs)
                            {
                                if (elementDict.ContainsKey(connected.Owner.Id))
                                {
                                    neighbors.Add(connected.Owner.Id);
                                }
                            }
                        }

                        // Create a mesh connection between all neighbors.
                        for (int i = 0; i < neighbors.Count; i++)
                        {
                            for (int j = i + 1; j < neighbors.Count; j++)
                            {
                                var neighbor1 = neighbors[i];
                                var neighbor2 = neighbors[j];
                                if (graph.ContainsKey(neighbor1) && graph.ContainsKey(neighbor2))
                                {
                                    graph[neighbor1].Add(neighbor2);
                                    graph[neighbor2].Add(neighbor1);
                                }
                            }
                        }
                    }
                    // Logic for curves (trays/conduits): connect to whatever is at their ends.
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
                                    graph[connected.Owner.Id].Add(curve.Id); // Ensure two-way connection
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log errors for elements that fail connector processing.
                    System.Diagnostics.Debug.WriteLine($"Could not process connectors for element {elem.Id}: {ex.Message}");
                }
            }

            // Ensure all adjacency lists contain unique elements.
            foreach (var key in graph.Keys.ToList())
            {
                graph[key] = graph[key].Distinct().ToList();
            }

            return graph;
        }


        /// <summary>
        /// Creates a map of ElementId to its length. For non-curve elements, length is 0.
        /// </summary>
        public Dictionary<ElementId, double> GetElementLengthMap()
        {
            return AdjacencyGraph.Keys.ToDictionary(id => id, id => (_doc.GetElement(id) as MEPCurve)?.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0);
        }

        /// <summary>
        /// A static helper to get the ConnectorSet from an element, supporting both MEPCurves and FamilyInstances.
        /// </summary>
        public static ConnectorSet GetConnectors(Element element)
        {
            if (element is MEPCurve curve) return curve.ConnectorManager?.Connectors;
            if (element is FamilyInstance fi) return fi.MEPModel?.ConnectorManager?.Connectors;
            return null;
        }
    }
}