//-----------------------------------------------------------------------------
// <copyright file="ContainmentGraph.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Represents the containment network as a graph for pathfinding.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Corrected GetConnectors to handle both MEPCurve and FamilyInstance elements.
// - [APPLIED FIX]: Enhanced GetElementLengthMap to calculate approximate lengths for fittings.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Represents the containment network as a graph for pathfinding.
    /// </summary>
    public class ContainmentGraph
    {
        /// <summary>
        /// Gets the adjacency graph where each key is an element ID and the value is a list of connected element IDs.
        /// </summary>
        public Dictionary<ElementId, List<ElementId>> AdjacencyGraph { get; } = new Dictionary<ElementId, List<ElementId>>();

        private readonly List<Element> _elements;
        private readonly Document _doc;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContainmentGraph"/> class.
        /// </summary>
        /// <param name="elements">The list of containment elements to include in the graph.</param>
        /// <param name="doc">The active Revit document.</param>
        public ContainmentGraph(List<Element> elements, Document doc)
        {
            _elements = elements;
            _doc = doc;
            BuildGraph();
        }

        /// <summary>
        /// Builds the adjacency graph from the provided containment elements.
        /// </summary>
        private void BuildGraph()
        {
            var elementDict = _elements.ToDictionary(e => e.Id);

            foreach (var element in _elements)
            {
                if (!AdjacencyGraph.ContainsKey(element.Id))
                {
                    AdjacencyGraph[element.Id] = new List<ElementId>();
                }

                var connectors = GetConnectors(element);
                if (connectors == null) continue;

                foreach (Connector connector in connectors)
                {
                    if (!connector.IsConnected) continue;

                    foreach (Connector connectedRef in connector.AllRefs)
                    {
                        var owner = connectedRef.Owner;
                        if (owner != null && elementDict.ContainsKey(owner.Id) && owner.Id != element.Id)
                        {
                            if (!AdjacencyGraph[element.Id].Contains(owner.Id))
                            {
                                AdjacencyGraph[element.Id].Add(owner.Id);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the connectors for a given containment element.
        /// </summary>
        public static ConnectorSet GetConnectors(Element element)
        {
            if (element is MEPCurve mepCurve)
            {
                return mepCurve.ConnectorManager?.Connectors;
            }
            if (element is FamilyInstance fi)
            {
                return fi.MEPModel?.ConnectorManager?.Connectors;
            }
            return null;
        }

        /// <summary>
        /// Creates a map of element IDs to their calculated lengths in feet.
        /// Includes calculated lengths for fittings.
        /// </summary>
        /// <returns>A dictionary mapping each element ID to its length.</returns>
        public Dictionary<ElementId, double> GetElementLengthMap()
        {
            var lengthMap = new Dictionary<ElementId, double>();
            foreach (var element in _elements)
            {
                double lengthInFeet = 0.0;
                if (element is MEPCurve mepCurve)
                {
                    lengthInFeet = mepCurve.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0;
                }
                else if (element is FamilyInstance fitting)
                {
                    lengthInFeet = GetFittingLength(fitting);
                }
                lengthMap[element.Id] = lengthInFeet;
            }
            return lengthMap;
        }

        /// <summary>
        /// Calculates an approximate centerline length for a fitting.
        /// </summary>
        /// <param name="fitting">The fitting element.</param>
        /// <returns>The calculated length in feet.</returns>
        private double GetFittingLength(FamilyInstance fitting)
        {
            var connectors = GetConnectors(fitting);
            if (connectors == null || connectors.Size < 2) return 0.0;

            // Case 1: Handle bends/elbows by calculating arc length
            Parameter angleParam = fitting.get_Parameter(BuiltInParameter.CONNECTOR_ANGLE);
            if (angleParam != null && angleParam.HasValue)
            {
                double angle = angleParam.AsDouble(); // Angle in radians
                // For fittings, "Bend Radius" is typically a family parameter, not a BuiltInParameter.
                Parameter radiusParam = fitting.LookupParameter("Bend Radius");

                if (radiusParam != null && radiusParam.HasValue)
                {
                    double radius = radiusParam.AsDouble();
                    return Math.Abs(angle * radius); // Arc length = angle * radius
                }
            }

            // Case 2: Fallback for other fittings (tees, crosses, transitions)
            // Approximate length by finding the maximum distance between any two connectors.
            var connectorList = connectors.Cast<Connector>().ToList();
            double maxLength = 0;
            for (int i = 0; i < connectorList.Count; i++)
            {
                for (int j = i + 1; j < connectorList.Count; j++)
                {
                    double dist = connectorList[i].Origin.DistanceTo(connectorList[j].Origin);
                    if (dist > maxLength)
                    {
                        maxLength = dist;
                    }
                }
            }
            return maxLength;
        }
    }
}