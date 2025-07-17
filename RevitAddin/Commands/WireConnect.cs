//
// --- FILE: WireConnect.cs ---
//
// File: WireConnect.cs
//
// Namespace: RTS.Commands
//
// Class: WireConnectClass
//
// Function: This Revit external command looks for unconnected wire ends on the active view
//           and connects them to the closest lighting fixture or lighting device within
//           a 2000mm XY-plane tolerance and a 3000mm Z-axis tolerance. On completion, it
//           provides a brief report to the user on the number of connections made and
//           to which category of element.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 17, 2025: Initial creation of the WireConnect command.
// - July 17, 2025: Corrected BuiltInCategory for wires (OST_ElectricalWires).
// - July 17, 2025: Rolled back preprocessor directives.
// - July 17, 2025: Corrected BuiltInCategory for wires again (OST_Wire).
// - July 17, 2025: Added check for ConnectorType.Conduit on target connectors and improved error logging for failed connections.
// - July 17, 2025: Modified distance calculation to consider 2D (XY) tolerance and a separate Z-axis tolerance.
// - July 17, 2025: Corrected 'ConnectorType.Conduit' error by using 'Domain.DomainConduit' for filtering.
// - July 17, 2025: Removed erroneous 'Domain.DomainConduit' check to fix compilation error.
// - July 17, 2025: Included the count of unconnected wire ends in the final report.
// - July 17, 2025: Enhanced diagnostics by adding more debug output and collecting specific connection failure messages.
// - July 17, 2025: Integrated detailed failed connection messages directly into the user-facing report.
// - July 17, 2025: Added logic to connect wires to other wires if no higher-priority lighting fixture/device is found within tolerance.
// - July 17, 2025: Fixed CS0266 error by using FilteredElementCollector.Excluding() instead of LINQ Where() for element exclusion.
// - July 17, 2025: Removed wire-to-wire connection feature; now only connects to lighting fixtures/devices.
// - July 17, 2025: Increased Z-axis tolerance to 3000mm.
// - July 17, 2025: Increased XY-plane tolerance to 1000mm.
// - July 17, 2025: Increased XY-plane tolerance to 2000mm.
//
// ---

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RTS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WireConnectClass : IExternalCommand
    {
        /// <summary>
        /// Calculates the 2D distance between two XYZ points (ignoring Z-coordinate).
        /// </summary>
        /// <param name="p1">First point.</param>
        /// <param name="p2">Second point.</param>
        /// <returns>The 2D distance.</returns>
        private double Calculate2DDistance(XYZ p1, XYZ p2)
        {
            double dx = p1.X - p2.X;
            double dy = p1.Y - p2.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the application and document objects
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Define the connection tolerance for XY plane (2000mm)
            // Revit's internal units are feet, so convert 2000mm to feet.
            // 1 foot = 304.8 mm
            double xyToleranceFeet = 2000.0 / 304.8; // Increased XY tolerance to 2000mm

            // Define the connection tolerance for Z-axis (3000mm, approx 3 meters)
            double zToleranceFeet = 3000.0 / 304.8;

            // Counters for the report
            int connectionsMade = 0;
            int unconnectedWireEndsCount = 0;
            Dictionary<string, int> connectionsByCategory = new Dictionary<string, int>();
            List<string> failedConnectionMessages = new List<string>(); // To collect specific error messages

            // Start a new transaction for making changes to the model
            using (Transaction trans = new Transaction(doc, "Connect Wires to Fixtures"))
            {
                try
                {
                    trans.Start();

                    // Get the active view
                    View activeView = doc.ActiveView;
                    if (activeView == null)
                    {
                        message = "No active view found. Please open a view.";
                        trans.RollBack();
                        return Result.Failed;
                    }

                    // 1. Collect all Wire elements in the active view
                    FilteredElementCollector wireCollector = new FilteredElementCollector(doc, activeView.Id)
                        .OfCategory(BuiltInCategory.OST_Wire)
                        .WhereElementIsNotElementType();

                    List<Wire> wiresInView = wireCollector.Cast<Wire>().ToList();
                    System.Diagnostics.Debug.WriteLine($"Found {wiresInView.Count} wire elements in the active view.");


                    // 2. Collect all Lighting Fixture and Lighting Device elements in the active view
                    // Create filters for Lighting Fixtures and Lighting Devices
                    ElementFilter lightingFixtureFilter = new ElementCategoryFilter(BuiltInCategory.OST_LightingFixtures);
                    ElementFilter lightingDeviceFilter = new ElementCategoryFilter(BuiltInCategory.OST_LightingDevices);

                    // Combine the filters using an OR filter
                    LogicalOrFilter lightingElementsFilter = new LogicalOrFilter(lightingFixtureFilter, lightingDeviceFilter);

                    FilteredElementCollector lightingElementCollector = new FilteredElementCollector(doc, activeView.Id)
                        .WherePasses(lightingElementsFilter)
                        .WhereElementIsNotElementType();

                    List<FamilyInstance> lightingElements = lightingElementCollector.Cast<FamilyInstance>().ToList();
                    System.Diagnostics.Debug.WriteLine($"Found {lightingElements.Count} lighting fixtures/devices in the active view.");

                    // Iterate through each wire to find unconnected ends
                    foreach (Wire wire in wiresInView)
                    {
                        ConnectorSet connectors = wire.ConnectorManager.Connectors;

                        foreach (Connector wireConnector in connectors)
                        {
                            if (!wireConnector.IsConnected)
                            {
                                unconnectedWireEndsCount++;
                                System.Diagnostics.Debug.WriteLine($"Unconnected wire end found at: X={wireConnector.Origin.X:F2}, Y={wireConnector.Origin.Y:F2}, Z={wireConnector.Origin.Z:F2} (Wire ID: {wire.Id})");

                                XYZ wireEndLocation = wireConnector.Origin;

                                FamilyInstance closestTargetElement = null; // Target is now explicitly a FamilyInstance
                                Connector targetConnectorToConnect = null;
                                double min2DDistance = xyToleranceFeet; // Initialize with XY tolerance

                                // Find the closest Lighting Fixture or Device
                                foreach (FamilyInstance lightingElement in lightingElements)
                                {
                                    if (lightingElement.MEPModel == null || lightingElement.MEPModel.ConnectorManager == null)
                                        continue;

                                    ConnectorSet targetConnectors = lightingElement.MEPModel.ConnectorManager.Connectors;

                                    foreach (Connector targetConnector in targetConnectors)
                                    {
                                        // Ensure the target connector is electrical and free.
                                        if (targetConnector.Domain == Domain.DomainElectrical &&
                                            !targetConnector.IsConnected)
                                        {
                                            double current2DDistance = Calculate2DDistance(wireEndLocation, targetConnector.Origin);
                                            double zDifference = Math.Abs(wireEndLocation.Z - targetConnector.Origin.Z);

                                            if (current2DDistance < min2DDistance && zDifference < zToleranceFeet)
                                            {
                                                min2DDistance = current2DDistance;
                                                closestTargetElement = lightingElement; // Assign FamilyInstance
                                                targetConnectorToConnect = targetConnector;
                                            }
                                        }
                                    }
                                }

                                // If a suitable closest lighting element and connector were found within tolerance
                                if (closestTargetElement != null && targetConnectorToConnect != null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Attempting to connect wire end (ID: {wire.Id}) at {wireEndLocation.ToString()} to connector (ID: {targetConnectorToConnect.Id}) on fixture '{closestTargetElement.Name}' (ID: {closestTargetElement.Id}) at {targetConnectorToConnect.Origin.ToString()}.");
                                    System.Diagnostics.Debug.WriteLine($"    2D Distance: {min2DDistance * 304.8:F0}mm, Z Difference: {Math.Abs(wireEndLocation.Z - targetConnectorToConnect.Origin.Z) * 304.8:F0}mm");
                                    System.Diagnostics.Debug.WriteLine($"    Target Connector Details: Domain={targetConnectorToConnect.Domain}, IsConnected={targetConnectorToConnect.IsConnected}, ConnectorType={targetConnectorToConnect.ConnectorType}");

                                    try
                                    {
                                        wireConnector.ConnectTo(targetConnectorToConnect);
                                        connectionsMade++;
                                        System.Diagnostics.Debug.WriteLine($"  SUCCESS: Connected wire end (ID: {wire.Id}) to '{closestTargetElement.Name}' (ID: {closestTargetElement.Id}).");

                                        string categoryName = closestTargetElement.Category.Name;
                                        if (connectionsByCategory.ContainsKey(categoryName))
                                        {
                                            connectionsByCategory[categoryName]++;
                                        }
                                        else
                                        {
                                            connectionsByCategory.Add(categoryName, 1);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        string errorMessage = $"  FAILED to connect wire end (ID: {wire.Id}) at {wireEndLocation.ToString()} to '{closestTargetElement.Name}' (ID: {closestTargetElement.Id}) at {targetConnectorToConnect.Origin.ToString()}. Error: {ex.Message}";
                                        System.Diagnostics.Debug.WriteLine(errorMessage);
                                        failedConnectionMessages.Add(errorMessage);
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"  No suitable lighting fixture/device found within tolerance for wire end at {wireEndLocation.ToString()} (Wire ID: {wire.Id}).");
                                }
                            }
                        }
                    }

                    trans.Commit();

                    // Generate the report
                    string report = $"Wire Connection Report:\nTotal connections made: {connectionsMade}\n";
                    report += $"Total unconnected wire ends found: {unconnectedWireEndsCount}\n\n";

                    if (connectionsMade > 0)
                    {
                        report += "Connections by Category:\n";
                        foreach (var entry in connectionsByCategory)
                        {
                            report += $"  {entry.Key}: {entry.Value}\n";
                        }
                    }
                    else
                    {
                        report += $"No new connections were made within the specified XY tolerance ({xyToleranceFeet * 304.8:F0}mm) and Z tolerance ({zToleranceFeet * 304.8:F0}mm).\n";
                    }

                    if (failedConnectionMessages.Any())
                    {
                        report += "\n--- Failed Connection Details ---\n";
                        foreach (string failMsg in failedConnectionMessages)
                        {
                            report += $"{failMsg}\n";
                        }
                        report += "---------------------------------\n";
                    }

                    TaskDialog.Show("Wire Connect Report", report);

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    // Roll back the transaction if any error occurs
                    if (trans.GetStatus() == TransactionStatus.Started)
                    {
                        trans.RollBack();
                    }
                    message = $"An unhandled error occurred during wire connection: {ex.Message}\n{ex.StackTrace}";
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }
            }
        }
    }
}
