//
// File: RT_WireRoute.cs
//
// Namespace: RT_WireRoute
//
// Class: RT_WireRouteClass
//
// Function: This Revit external command looks up the RTS_ID parameter on
//           Wire elements. Where it matches the RTS_ID of Conduit or
//           Conduit Fitting elements, it updates the wire's geometry
//           to pass through those conduits using orthogonal routing where possible.
//           It achieves this by deleting the existing wire and creating a new one
//           with the calculated path. It attempts to find the shortest traversal path
//           through identified conduits, including "jumping" gaps with default pathing
//           if necessary. If no viable conduit pathway exists for a wire's RTS_ID,
//           the wire is skipped. Includes robust error handling and user reporting.
//
// Author: AI (Based on user requirements)
//
// Date: June 28, 2025 (Updated with WiringType in Wire.Create)
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO; // Used for Path, even if not for direct file IO in this main logic
using System.Linq;
using System.Text;
using System.Windows.Forms; // For TaskDialog
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Electrical; // For Wire, ElectricalSystem, Connector, WiringType (UNCOMMENTED THIS LINE)
using System.Diagnostics;
#endregion

// IMPORTANT COMPATIBILITY GUIDANCE:
// This C# code file is designed to be compatible with both Revit 2022 and Revit 2024
// (and potentially other versions with similar API). However, for successful compilation
// and execution, you MUST set up your Visual Studio project correctly for the target Revit version.
//
// Key Steps for Dual-Version Compatibility:
// 1.  PROJECT STRUCTURE:
//     * The most robust approach is to have SEPARATE Visual Studio projects for each Revit version
//       you intend to support (e.g., "RTS_Revit2022.csproj" and "RTS_Revit2024.csproj").
//     * You can then "Add Existing Item" to these projects and choose "Add as Link"
//       for your .cs source code files (like this one). This way, you maintain a single
//       set of C# source files, but compile a distinct DLL for each Revit version.
//
// 2.  TARGET .NET FRAMEWORK:
//     * For Revit 2022: Your Visual Studio project MUST target .NET Framework 4.8.
//     * For Revit 2024: Your Visual Studio project MUST target .NET 7.0.
//     (Right-click Project -> Properties -> Application/Target Framework dropdown)
//
// 3.  REVT API REFERENCES:
//     * For Revit 2022: Reference Autodesk.Revit.DB.dll and Autodesk.Revit.UI.dll from
//       your Revit 2022 installation directory (e.g., C:\Program Files\Autodesk\Revit 2022\).
//     * For Revit 2024: Reference Autodesk.Revit.DB.dll and Autodesk.Revit.UI.dll from
//       your Revit 2024 installation directory (e.g., C:\Program Files\Autodesk\Revit 2024\).
//     (Right-click Project -> Add -> Project Reference -> Browse)
//
// 4.  "COPY LOCAL" SETTING:
//     * For ALL Revit API references (RevitAPI.dll, RevitAPIUI.dll, and any others like AdWindows.dll etc.):
//       Set "Copy Local" to FALSE in the Properties window for each reference.
//       (Select reference in Solution Explorer -> Properties window -> Copy Local = False)
//
// 5.  PLATFORM TARGET:
//     * Ensure your project is configured to build for "x64" (64-bit).
//     (Right-click Project -> Properties -> Build/Platform target dropdown)
//
// Following these steps will ensure the compiler can correctly resolve API calls like
// 'Wire.WireType' and 'Wire.Create' for the specific Revit version you are building against.


namespace RTS.Commands.MEPTools.Electrical
{
    /// <summary>
    /// Revit External Command to route electrical wires through conduits based on matching RTS_ID.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_WireRouteClass : IExternalCommand
    {
        // Shared Parameter GUID for RTS_ID (Must match PC_Extensible, PC_WireData, etc.)
        private static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
        private const string RTS_ID_NAME = "RTS_ID";

        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Correctly get UIApplication and UIDocument
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            int wiresUpdatedCount = 0;
            List<string> skippedWires = new List<string>();

            // Collect all conduits and conduit fittings by their RTS_ID
            // Dictionary<RTS_ID, List<Element>>
            var conduitsByRtsId = CollectConduitsByRtsId(doc);

            // Get the active view's ElementId, required for Wire.Create
            ElementId activeViewId = uidoc.ActiveView.Id;
            // Determine a default WiringType. Chamfer is often preferred for "orthogonal" routing.
            WiringType defaultWiringType = WiringType.Chamfer;
            // If you needed to extract the old wire's WiringType, you would use:
            // Parameter oldWiringTypeParam = wire.get_Parameter(BuiltInParameter.RBS_WIRE_WIRING_TYPE);
            // if (oldWiringTypeParam != null) defaultWiringType = (WiringType)oldWiringTypeParam.AsInteger();


            using (Transaction tx = new Transaction(doc, "Route Wires Through Conduits"))
            {
                tx.Start();
                try
                {
                    // Filter for Wire elements
                    FilteredElementCollector wireCollector = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Wire)
                        .WhereElementIsNotElementType();

                    foreach (Element wireElement in wireCollector)
                    {
                        // Explicitly cast to Wire for specific wire properties.
                        // This ensures we are only processing actual Wire elements.
                        Wire wire = wireElement as Wire;
                        if (wire == null)
                        {
                            Debug.WriteLine($"Skipping element {wireElement.Id} as it is not a Wire element.");
                            continue;
                        }

                        // Get essential properties from the existing wire before deletion
                        // The WireType property returns the ElementId of the WiringType Element.
                        // FIX: Using GetTypeId() which is a general method to get the ElementId of an element's type,
                        //      as Wire.WireType might be causing compiler issues depending on the Revit API assembly version.
                        ElementId oldWireTypeId = wire.GetTypeId();
                        string oldWireRtsId = null;
                        Parameter oldWireRtsIdParam = wire.get_Parameter(RTS_ID_GUID);
                        if (oldWireRtsIdParam != null && oldWireRtsIdParam.HasValue)
                        {
                            oldWireRtsId = oldWireRtsIdParam.AsString().Trim();
                        }

                        // Get the wire's own endpoints. These are crucial for the new wire's creation.
                        List<XYZ> wireEndpoints = GetWireEndpoints(wire);
                        if (wireEndpoints.Count != 2)
                        {
                            skippedWires.Add($"Wire ID: {wire.Id} (UniqueId: {wire.UniqueId}) - Could not determine wire's two distinct endpoints. Skipping.");
                            continue;
                        }

                        XYZ wireStartPoint = wireEndpoints[0];
                        XYZ wireEndPoint = wireEndpoints[1];

                        if (string.IsNullOrEmpty(oldWireRtsId))
                        {
                            skippedWires.Add($"Wire ID: {wire.Id} (UniqueId: {wire.UniqueId}) - No '{RTS_ID_NAME}' parameter value found. Skipping.");
                            continue;
                        }

                        // Try to find conduits matching this RTS_ID
                        if (!conduitsByRtsId.TryGetValue(oldWireRtsId, out List<Element> matchingConduits))
                        {
                            skippedWires.Add($"Wire ID: {wire.Id} (RTS_ID: '{oldWireRtsId}', UniqueId: {wire.UniqueId}) - No conduits or fittings found with matching '{RTS_ID_NAME}'. Skipping.");
                            continue;
                        }

                        // Attempt to build a path using the matched conduits' connectors
                        List<XYZ> pathPoints = BuildConduitPath(uidoc, matchingConduits, wireStartPoint, wireEndPoint);

                        if (pathPoints.Count < 2) // Need at least two points (start and end)
                        {
                            skippedWires.Add($"Wire ID: {wire.Id} (RTS_ID: '{oldWireRtsId}', UniqueId: {wire.UniqueId}) - Matching conduits do not form a viable continuous path between wire endpoints. Skipping.");
                            continue;
                        }

                        // --- DELETE THE OLD WIRE AND CREATE A NEW ONE ---
                        try
                        {
                            // Delete the old wire
                            doc.Delete(wire.Id);
                            Debug.WriteLine($"Old Wire ID: {wire.Id} (UniqueId: {wire.UniqueId}) deleted.");

                            // Create a new wire with the specified path.
                            // Using the overload: Wire.Create(Document, ElementId wireTypeId, ElementId viewId, WiringType wiringType, IList<XYZ> vertexPoints, Connector startConnectorTo, Connector endConnectorTo)
                            // We pass null for startConnectorTo and endConnectorTo, as the pathPoints define the geometry explicitly.
                            Wire newWire = Wire.Create(doc, oldWireTypeId, activeViewId, defaultWiringType, pathPoints, null, null);

                            if (newWire != null)
                            {
                                // Update the RTS_ID parameter on the NEW wire
                                Parameter newWireRtsIdParam = newWire.get_Parameter(RTS_ID_GUID);
                                if (newWireRtsIdParam != null && !newWireRtsIdParam.IsReadOnly)
                                {
                                    newWireRtsIdParam.Set(oldWireRtsId);
                                }
                                else
                                {
                                    Debug.WriteLine($"Warning: New Wire ID: {newWire.Id} - Could not set RTS_ID parameter (read-only or not found).");
                                }

                                wiresUpdatedCount++;
                                Debug.WriteLine($"New Wire ID: {newWire.Id} (UniqueId: {newWire.UniqueId}) created and routed successfully.");
                            }
                            else
                            {
                                skippedWires.Add($"Original Wire ID: {wire.Id} (RTS_ID: '{oldWireRtsId}', UniqueId: {wire.UniqueId}) - Failed to create new wire with path. Skipping.");
                                Debug.WriteLine($"Failed to create new wire for original Wire ID {wire.Id} (UniqueId: {wire.UniqueId}).");
                            }
                        }
                        catch (Exception geoEx)
                        {
                            skippedWires.Add($"Original Wire ID: {wire.Id} (RTS_ID: '{oldWireRtsId}', UniqueId: {wire.UniqueId}) - Failed to delete/recreate wire: {geoEx.Message}. Skipping.");
                            Debug.WriteLine($"Error during delete/recreate process for Original Wire ID {wire.Id} (UniqueId: {wire.UniqueId}): {geoEx.Message}");
                        }
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    message = $"An unexpected error occurred during wire routing transaction: {ex.Message}";
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }
            }

            // Final Report
            StringBuilder report = new StringBuilder();
            report.AppendLine("Wire Routing Process Complete.");
            report.AppendLine($"Wires Updated: {wiresUpdatedCount}");

            if (skippedWires.Any())
            {
                report.AppendLine("\nSkipped Wires (with reasons):");
                // Limit the number of detailed messages for very large models
                foreach (string reason in skippedWires.Take(20))
                {
                    report.AppendLine($"- {reason}");
                }
                if (skippedWires.Count > 20)
                {
                    report.AppendLine($"...and {skippedWires.Count - 20} more skipped wires.");
                }
            }
            else
            {
                report.AppendLine("\nNo wires were skipped.");
            }

            TaskDialog.Show("RT_WireRoute Report", report.ToString());

            return Result.Succeeded;
        }

        /// <summary>
        /// Collects all Conduit and Conduit Fitting elements grouped by their RTS_ID.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>A dictionary where key is RTS_ID and value is a list of matching elements.</returns>
        private Dictionary<string, List<Element>> CollectConduitsByRtsId(Document doc)
        {
            var conduitsByRtsId = new Dictionary<string, List<Element>>(StringComparer.OrdinalIgnoreCase);

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };
            var categoryFilter = new ElementMulticategoryFilter(categories);

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(categoryFilter)
                .WhereElementIsNotElementType();

            foreach (Element elem in collector)
            {
                Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                if (rtsIdParam != null && rtsIdParam.HasValue)
                {
                    string rtsId = rtsIdParam.AsString().Trim();
                    if (!string.IsNullOrEmpty(rtsId))
                    {
                        if (!conduitsByRtsId.ContainsKey(rtsId))
                        {
                            conduitsByRtsId[rtsId] = new List<Element>();
                        }
                        conduitsByRtsId[rtsId].Add(elem);
                    }
                }
            }
            return conduitsByRtsId;
        }

        /// <summary>
        /// Extracts the two distinct XYZ endpoint locations of a wire.
        /// </summary>
        /// <param name="wire">The wire element.</param>
        /// <returns>A list containing exactly two XYZ points, representing the wire's start and end.</returns>
        private List<XYZ> GetWireEndpoints(Wire wire)
        {
            List<XYZ> endpoints = new List<XYZ>();
            ConnectorSet connectors = wire.ConnectorManager.Connectors;
            if (connectors != null && connectors.Size >= 2) // A wire should have at least two connectors
            {
                // Wires are modeled as curves between connectors. Get the origins of the first two.
                // In a simple wire, these should be the two distinct endpoints.
                endpoints.Add(connectors.Cast<Connector>().ElementAt(0).Origin);
                endpoints.Add(connectors.Cast<Connector>().ElementAt(1).Origin);
            }
            return endpoints;
        }


        /// <summary>
        /// Attempts to build an ordered path of XYZ points for a wire, prioritizing matched conduits.
        /// It tries to traverse connected conduits. If a gap is found, it attempts to "jump" to the
        /// next closest unvisited conduit part with the same RTS_ID, otherwise it routes directly
        /// to the wire's end point.
        /// </summary>
        /// <param name="uidoc">The Revit UI Document (for active view scale for tolerance).</param>
        /// <param name="conduitElements">List of conduit and fitting elements with matching RTS_ID for the current wire.</param>
        /// <param name="wireStartPoint">The actual start XYZ point for the wire (from its own connector).</param>
        /// <param name="wireEndPoint">The actual end XYZ point for the wire (from its own connector).</param>
        /// <returns>An ordered list of XYZ points representing the path, including wire start and end points.</returns>
        private List<XYZ> BuildConduitPath(UIDocument uidoc, List<Element> conduitElements, XYZ wireStartPoint, XYZ wireEndPoint) // Added UIDocument
        {
            List<XYZ> pathPoints = new List<XYZ>();
            HashSet<ElementId> visitedElements = new HashSet<ElementId>(); // Tracks visited conduit/fitting elements within this specific pathfinding attempt

            // Define a small tolerance for "reaching" the wireEndPoint or comparing XYZ points
            // Based on view scale for better adaptability, falling back to a default if scale is not available.
            double targetReachTolerance = 0.5; // Default to 0.5 feet
            if (uidoc.ActiveView != null && uidoc.ActiveView.Scale > 0)
            {
                // Roughly 1/10th of a foot or less, scaled. Adjust as needed.
                targetReachTolerance = 1.0 / uidoc.ActiveView.Scale * 0.1;
                if (targetReachTolerance < 0.001) targetReachTolerance = 0.001; // Minimum tolerance
            }
            XYZComparer.Instance.Tolerance = targetReachTolerance; // Set the tolerance for XYZ comparisons

            // 1. Add the wire's actual start point to the path
            pathPoints.Add(wireStartPoint);

            // 2. Find the starting conduit/fitting for the path
            // This is the conduit part that is closest to the wire's start point
            Element currentConduitPart = null;
            Connector currentConduitConnector = null;
            double minDistToWireStart = double.MaxValue;

            foreach (Element conduitPart in conduitElements)
            {
                ConnectorManager innerCm = GetConnectorManager(conduitPart); // Renamed to innerCm to avoid CS0136
                if (innerCm == null) continue;

                foreach (Connector c in innerCm.Connectors)
                {
                    double dist = c.Origin.DistanceTo(wireStartPoint);
                    if (dist < minDistToWireStart)
                    {
                        minDistToWireStart = dist;
                        currentConduitPart = conduitPart;
                        currentConduitConnector = c;
                    }
                }
            }

            // If no starting conduit part is found, the path cannot be built through conduits.
            if (currentConduitPart == null || currentConduitConnector == null)
            {
                // Debug.WriteLine($"No starting conduit part closest to wire start point {wireStartPoint}. Routing directly.");
                // Fallback to direct routing if no matching conduit is found at all or close enough to start
                pathPoints.Add(wireEndPoint); // Only start and end points
                return pathPoints.Distinct(XYZComparer.Instance).ToList();
            }

            // Add the entry point into the conduit network (the closest connector on the first conduit part)
            // Ensure this point is distinct from the wireStartPoint if they are the same
            if (wireStartPoint.DistanceTo(currentConduitConnector.Origin) > XYZComparer.Instance.Tolerance)
            {
                pathPoints.Add(currentConduitConnector.Origin);
            }
            visitedElements.Add(currentConduitPart.Id);

            // 3. Traverse the conduit network (greedy search for next segment)
            XYZ lastPointAdded = currentConduitConnector.Origin; // The point from which we search for the next segment
            Element lastElementAdded = currentConduitPart;

            while (lastElementAdded != null && lastPointAdded.DistanceTo(wireEndPoint) > targetReachTolerance)
            {
                ConnectorManager cm = GetConnectorManager(lastElementAdded);
                if (cm == null) break;

                Connector exitConnector = null;
                foreach (Connector conn in cm.Connectors)
                {
                    // Check if this connector is the one representing the 'lastPointAdded' on the current element
                    if (conn.Origin.DistanceTo(lastPointAdded) < XYZComparer.Instance.Tolerance)
                    {
                        exitConnector = conn;
                        break;
                    }
                }

                if (exitConnector == null || !exitConnector.IsConnected)
                {
                    Debug.WriteLine($"No valid exit connector found from {lastElementAdded.Id} at {lastPointAdded}. Attempting gap jump.");
                    // Path is broken here, or we've reached an end. Attempt to jump to next closest conduit.
                    break;
                }

                // Attempt to find a directly connected, unvisited conduit with the same RTS_ID
                Element nextDirectConduitElement = null;
                Connector nextDirectConduitConnector = null;

                foreach (Connector connectedRef in exitConnector.AllRefs)
                {
                    if (connectedRef.Owner != null && !visitedElements.Contains(connectedRef.Owner.Id))
                    {
                        // Check if this connected element is one of our matching conduits
                        if (conduitElements.Any(e => e.Id == connectedRef.Owner.Id))
                        {
                            nextDirectConduitElement = connectedRef.Owner;
                            nextDirectConduitConnector = connectedRef;
                            break;
                        }
                    }
                }

                if (nextDirectConduitElement != null)
                {
                    // Found a direct connection within the matching conduits
                    if (nextDirectConduitConnector.Origin.DistanceTo(lastPointAdded) > XYZComparer.Instance.Tolerance)
                    {
                        pathPoints.Add(nextDirectConduitConnector.Origin);
                    }
                    visitedElements.Add(nextDirectConduitElement.Id);
                    lastElementAdded = nextDirectConduitElement;
                    lastPointAdded = nextDirectConduitConnector.Origin;
                }
                else // No direct physical connection found within matching RTS_ID conduits (GAP DETECTED)
                {
                    Debug.WriteLine($"Gap detected from {lastElementAdded.Id}. Searching for next closest unvisited conduit.");

                    Element nextClosestConduitElement = null;
                    Connector nextClosestConduitConnector = null;
                    double minDistToNextClosest = double.MaxValue;

                    foreach (Element conduitPart in conduitElements)
                    {
                        if (!visitedElements.Contains(conduitPart.Id)) // Only consider unvisited conduits
                        {
                            ConnectorManager innerCm2 = GetConnectorManager(conduitPart); // Renamed again for clarity
                            if (innerCm2 == null) continue;

                            foreach (Connector c in innerCm2.Connectors)
                            {
                                double dist = c.Origin.DistanceTo(lastPointAdded);
                                if (dist < minDistToNextClosest)
                                {
                                    minDistToNextClosest = dist;
                                    nextClosestConduitElement = conduitPart;
                                    nextClosestConduitConnector = c;
                                }
                            }
                        }
                    }

                    if (nextClosestConduitElement != null && nextClosestConduitConnector != null)
                    {
                        // Found a conduit to jump to. Add its entry point to the path.
                        if (nextClosestConduitConnector.Origin.DistanceTo(lastPointAdded) > XYZComparer.Instance.Tolerance)
                        {
                            pathPoints.Add(nextClosestConduitConnector.Origin); // This is the "jump" point
                        }
                        visitedElements.Add(nextClosestConduitElement.Id);
                        lastElementAdded = nextClosestConduitElement;
                        lastPointAdded = nextClosestConduitConnector.Origin;
                        Debug.WriteLine($"Jumped to conduit {nextClosestConduitElement.Id} at {lastPointAdded}.");
                    }
                    else
                    {
                        // No more unvisited matching conduits to jump to. Stop conduit-guided path.
                        Debug.WriteLine("No more unvisited matching conduits to jump to. Ending conduit-guided path.");
                        break;
                    }
                }
            }

            // 4. Add the wire's actual end point to the path
            // This ensures the wire segment explicitly ends where the wire is supposed to end
            if (pathPoints.Last().DistanceTo(wireEndPoint) > XYZComparer.Instance.Tolerance)
            {
                pathPoints.Add(wireEndPoint);
            }

            // Remove duplicates (can happen if multiple connectors are at the same point, common with fittings)
            pathPoints = pathPoints.Distinct(XYZComparer.Instance).ToList();

            // Final validation: Ensure the path starts at wireStartPoint and ends at wireEndPoint
            // Re-check after distinct as points might shift or be removed
            if (pathPoints.Count < 2 || wireStartPoint.DistanceTo(pathPoints[0]) > XYZComparer.Instance.Tolerance || wireEndPoint.DistanceTo(pathPoints.Last()) > XYZComparer.Instance.Tolerance)
            {
                // This path is not valid or complete between the wire's actual endpoints and the conduit run
                Debug.WriteLine($"Generated path does not correctly connect wire endpoints via conduits for {wireStartPoint} to {wireEndPoint}. Path invalid.");
                return new List<XYZ>();
            }

            return pathPoints;
        }

        /// <summary>
        /// Helper to get ConnectorManager from an Element.
        /// </summary>
        private ConnectorManager GetConnectorManager(Element elem)
        {
            if (elem is Conduit conduit) return conduit.ConnectorManager;
            // Correctly handle ConduitFitting - it's a FamilyInstance, and MEPModel holds ConnectorManager
            if (elem is FamilyInstance fitting && fitting.MEPModel?.ConnectorManager != null) return fitting.MEPModel.ConnectorManager;
            return null;
        }

        // Helper class for Distinct() on XYZ points
        private class XYZComparer : IEqualityComparer<XYZ>
        {
            public static readonly XYZComparer Instance = new XYZComparer();
            // A small tolerance for comparing XYZ points. Set dynamically based on view scale.
            public double Tolerance { get; set; } = 1e-6; // Default value, will be updated by BuildConduitPath

            public bool Equals(XYZ x, XYZ y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(x, null) || ReferenceEquals(y, null)) return false;

                // Use DistanceTo for robust XYZ comparison with tolerance
                return x.DistanceTo(y) < Tolerance;
            }

            public int GetHashCode(XYZ obj)
            {
                if (obj == null) return 0;
                // A simple hash code that groups points within tolerance.
                // The Equals method will handle true equality check for Distinct().
                // Multiplying by 1000 and casting to int effectively quantizes the points.
                return ((int)(obj.X / Tolerance)).GetHashCode() ^
                       ((int)(obj.Y / Tolerance)).GetHashCode() ^
                       ((int)(obj.Z / Tolerance)).GetHashCode();
            }
        }

        #region Data Classes (Duplicated for self-containment, ensure consistency across scripts)
        // Note: These classes are duplicated across PC_Extensible.cs, PC_WireData.cs, and RT_WireRoute.cs.
        // For larger projects, consider moving them to a shared utility assembly.
        public class CableData
        {
            public string CableReference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string CableType { get; set; }
            public string CableCode { get; set; }
            public string CableConfiguration { get; set; }
            public string Cores { get; set; }
            public string ConductorActive { get; set; }
            public string Insulation { get; set; }
            public string ConductorEarth { get; set; }
            public string SeparateEarthForMulticore { get; set; }
            public string CableLength { get; set; }
            public string TotalCableRunWeight { get; set; }
            public string NominalOverallDiameter { get; set; }
            public string NumberOfActiveCables { get; set; }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; }
            public string AsNsz3008CableDeratingFactor { get; set; }
        }
        #endregion
    }
}
