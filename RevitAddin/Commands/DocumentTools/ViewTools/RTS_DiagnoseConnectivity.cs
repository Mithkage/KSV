//
// File: RTS_DiagnoseConnectivity.cs
//
// Namespace: RTS.Commands.DocumentTools.ViewTools
//
// Class: RTS_DiagnoseConnectivityClass
//
// Function: This Revit external command provides a diagnostic tool to visualize the
//           connectivity graph for a specific cable run. The user inputs a Cable Reference ID,
//           and the tool highlights all physically connected containment sections in red
//           and draws the calculated virtual jumps as dashed magenta model lines.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Corrected API call to create ModelCurves instead of DetailCurves to resolve error in 3D views.
// - [APPLIED FIX]: Added check to warn user if containment elements are not visible in the active view.
// - [APPLIED FIX]: Removed CleanCableReference logic, assuming data is pre-cleaned in the model.
// 2024-08-14:
// - [APPLIED FIX]: Added robust error handling for model line creation to prevent command failures.
// - [APPLIED FIX]: Implemented marker visualization at connection points as fallback when lines fail.
// 2024-08-15:
// - [APPLIED FIX]: Replaced ModelCurve with DirectShape for connection line visualization to avoid API errors.
// - [APPLIED FIX]: Added workset handling to ensure DirectShape creation happens in an editable workset.
//
// Author: ReTick Solutions
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Reports.Generators;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Interop;
#if REVIT2024_OR_GREATER
using Autodesk.Revit.DB.Worksharing;
#endif
#endregion

namespace RTS.Commands.DocumentTools.ViewTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_DiagnoseConnectivityClass : IExternalCommand
    {
        private static string _lastUserInput = string.Empty;
        private const string VIRTUAL_JUMP_NAME = "RTS_VirtualJump";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            // Check if document is in a modifiable state
            System.Diagnostics.Debug.WriteLine($"Document modifiable state: {doc.IsModifiable}");

            // Instead of warning immediately, store the state and only warn if we see an actual transaction problem
            bool initiallyModifiable = doc.IsModifiable;

            if (activeView.ViewType != ViewType.ThreeD)
            {
                TaskDialog.Show("Error", "This command can only be run in a 3D view.");
                return Result.Cancelled;
            }

            // --- 1. Get User Input for Cable Reference ---
            string cableRef;
            try
            {
                InputWindow inputWindow = new InputWindow { UserInput = _lastUserInput };
                new WindowInteropHelper(inputWindow).Owner = commandData.Application.MainWindowHandle;
                if (inputWindow.ShowDialog() != true) return Result.Cancelled;

                cableRef = inputWindow.UserInput;
                _lastUserInput = cableRef;
            }
            catch (Exception ex)
            {
                message = $"Error displaying input window: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            if (string.IsNullOrWhiteSpace(cableRef))
            {
                ClearAllOverrides(doc, activeView, true); // Use its own transaction since we're not in one
                TaskDialog.Show("Info", "Input was blank. All graphic overrides and diagnostic lines have been cleared.");
                return Result.Succeeded;
            }

            // --- 2. Gather Data ---
            var containmentCollector = new ContainmentCollector(doc);
            var allContainmentInProject = containmentCollector.GetAllContainmentElements();
            var cableGuids = SharedParameters.Cable.AllCableGuids;

            var cableContainmentElements = allContainmentInProject
                .Where(elem => cableGuids.Any(guid => string.Equals(elem.get_Parameter(guid)?.AsString(), cableRef, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (!cableContainmentElements.Any())
            {
                TaskDialog.Show("Info", $"No containment found with the Cable Reference: '{cableRef}'");
                return Result.Succeeded;
            }

            // --- 3. Improved Error Handling: Check if elements are visible ---
            var visibleContainmentElements = new FilteredElementCollector(doc, activeView.Id)
                .WherePasses(new ElementIdSetFilter(cableContainmentElements.Select(e => e.Id).ToList()))
                .ToElements();

            if (!visibleContainmentElements.Any())
            {
                TaskDialog.Show("Warning", $"Containment for cable '{cableRef}' was found in the project, but none of it is visible in the current 3D view. Please check view filters and visibility settings.");
                return Result.Succeeded;
            }

            var containmentGraph = new ContainmentGraph(cableContainmentElements, doc);
            var islands = Pathfinder.GroupIntoIslands(cableContainmentElements, containmentGraph, doc);

            // --- 4. Visualize the Path ---
            using (Transaction tx = new Transaction(doc, "Diagnose Cable Route"))
            {
                try
                {
                    // Try to start the transaction, if this fails we'll catch it below
                    tx.Start();

                    // If we got this far but the document wasn't initially modifiable, 
                    // log for diagnostics but no need to show a warning
                    if (!initiallyModifiable)
                    {
                        System.Diagnostics.Debug.WriteLine("Document was initially not modifiable but transaction started successfully");
                    }

                    ClearAllOverrides(doc, activeView, false); // Don't use its own transaction - we're already in one

                    var solidRedOverride = CreateSolidFillOverride(doc, new Color(255, 0, 0));
                    var transparentOverride = new OverrideGraphicSettings();
                    transparentOverride.SetHalftone(true);
                    transparentOverride.SetSurfaceTransparency(80);

                    var allElementsInView = new FilteredElementCollector(doc, activeView.Id).WhereElementIsNotElementType().ToElementIds();
                    var islandElementIds = new HashSet<ElementId>(islands.SelectMany(i => i.ElementIds));

                    foreach (var elemId in allElementsInView)
                    {
                        if (islandElementIds.Contains(elemId))
                        {
                            activeView.SetElementOverrides(elemId, solidRedOverride);
                        }
                        else
                        {
                            activeView.SetElementOverrides(elemId, transparentOverride);
                        }
                    }

                    if (islands.Count > 1)
                    {
                        var islandGraph = Pathfinder.BuildIslandGraph(islands, doc);
                        var lineStyle = GetOrCreateLineStyle(doc, "RTS_Virtual_Path", new Color(255, 0, 255)); // Magenta
                        var magentaOverride = CreateSolidFillOverride(doc, new Color(255, 0, 255)); // Magenta

                        // Changed to track success for final user feedback
                        bool anyPointsVisualized = false;
                        bool anyLinesVisualized = false;

                        // Store all possible connection points for visualization
                        HashSet<XYZ> connectionPoints = new HashSet<XYZ>();
                        List<Tuple<XYZ, XYZ>> connectionLines = new List<Tuple<XYZ, XYZ>>();

                        for (int i = 0; i < islands.Count; i++)
                        {
                            for (int j = i + 1; j < islands.Count; j++)
                            {
                                var endpointsA = Pathfinder.GetIslandEndpoints(islands[i], doc);
                                var endpointsB = Pathfinder.GetIslandEndpoints(islands[j], doc);

                                XYZ closestP1 = null;
                                XYZ closestP2 = null;
                                double minDistance = double.MaxValue;

                                foreach (var p1 in endpointsA)
                                {
                                    foreach (var p2 in endpointsB)
                                    {
                                        double dist = p1.DistanceTo(p2);
                                        if (dist < minDistance)
                                        {
                                            minDistance = dist;
                                            closestP1 = p1;
                                            closestP2 = p2;
                                        }
                                    }
                                }

                                if (closestP1 != null && closestP2 != null)
                                {
                                    // Add these points to our collection
                                    connectionPoints.Add(closestP1);
                                    connectionPoints.Add(closestP2);
                                    connectionLines.Add(new Tuple<XYZ, XYZ>(closestP1, closestP2));
                                }
                            }
                        }

                        // Create reference spheres at connection points
                        foreach (XYZ point in connectionPoints)
                        {
                            try
                            {
                                // Create a family instance at each connection point
                                FamilySymbol sphereSymbol = GetReferencePointSymbol(doc);
                                if (sphereSymbol != null && !sphereSymbol.IsActive)
                                {
                                    sphereSymbol.Activate();
                                    doc.Regenerate();
                                }

                                if (sphereSymbol != null)
                                {
                                    FamilyInstance marker = doc.Create.NewFamilyInstance(point,
                                        sphereSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    // Apply magenta override to the marker
                                    activeView.SetElementOverrides(marker.Id, magentaOverride);
                                    anyPointsVisualized = true;
                                }
                            }
                            catch (Exception)
                            {
                                // Silently continue if one point fails
                            }
                        }

                        // Try to create connection lines using DirectShape with fallback options
                        bool modelLinesFailed = false;
                        int successfulLines = 0;
                        bool attemptedFallback = false;

                        try
                        {
                            // Check for workset permissions first
                            bool worksetOK = true;
                            WorksetId editableWorksetId = WorksetId.InvalidWorksetId;

                            if (doc.IsWorkshared)
                            {
                                WorksetTable worksetTable = doc.GetWorksetTable();
                                WorksetId activeId = worksetTable.GetActiveWorksetId();

                                if (!IsWorksetEditable(doc, activeId))
                                {
                                    System.Diagnostics.Debug.WriteLine("Active workset is not editable, looking for alternatives...");
                                    editableWorksetId = GetEditableWorksetId(doc);
                                    worksetOK = (editableWorksetId != WorksetId.InvalidWorksetId);

                                    if (!worksetOK)
                                    {
                                        System.Diagnostics.Debug.WriteLine("WARNING: No editable worksets available!");
                                        TaskDialog.Show("Warning", "Cannot create virtual jumps because no editable worksets are available. Check your workset permissions.");
                                        modelLinesFailed = true;
                                    }
                                }
                                else
                                {
                                    editableWorksetId = activeId;
                                }
                            }

                            if (worksetOK)
                            {
                                // Add the capability check to diagnose permissions
                                bool canCreateDirectShape = true;
                                try
                                {
                                    // Check if we can create a test DirectShape
                                    using (SubTransaction testTx = new SubTransaction(doc))
                                    {
                                        testTx.Start();

                                        // Set workset if needed
                                        WorksetId originalWorksetId = WorksetId.InvalidWorksetId;
                                        if (doc.IsWorkshared && editableWorksetId != WorksetId.InvalidWorksetId)
                                        {
                                            originalWorksetId = doc.GetWorksetTable().GetActiveWorksetId();
                                            doc.GetWorksetTable().SetActiveWorksetId(editableWorksetId);
                                        }

                                        // Test DirectShape creation
                                        XYZ testPoint1 = new XYZ(0, 0, 0);
                                        XYZ testPoint2 = new XYZ(1, 1, 1);
                                        Line testLine = Line.CreateBound(testPoint1, testPoint2);
                                        DirectShape testDs = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                                        testDs.SetShape(new List<GeometryObject> { testLine });
                                        testDs.Name = "RTS_TestShape";

                                        // Restore original workset if needed
                                        if (doc.IsWorkshared && originalWorksetId != WorksetId.InvalidWorksetId)
                                        {
                                            doc.GetWorksetTable().SetActiveWorksetId(originalWorksetId);
                                        }

                                        testTx.RollBack(); // We don't need to keep this test shape
                                    }
                                }
                                catch (Exception checkEx)
                                {
                                    canCreateDirectShape = false;
                                    System.Diagnostics.Debug.WriteLine($"Cannot create DirectShape elements: {checkEx.Message}");
                                }

                                // Process connection lines based on capability
                                if (canCreateDirectShape)
                                {
                                    // Temporarily switch workset if needed
                                    WorksetId originalWorksetId = WorksetId.InvalidWorksetId;
                                    if (doc.IsWorkshared && editableWorksetId != WorksetId.InvalidWorksetId)
                                    {
                                        originalWorksetId = doc.GetWorksetTable().GetActiveWorksetId();
                                        doc.GetWorksetTable().SetActiveWorksetId(editableWorksetId);
                                    }

                                    // Standard approach - DirectShape with lines
                                    foreach (var connectionPair in connectionLines)
                                    {
                                        try
                                        {
                                            // Standard DirectShape creation
                                            if (connectionPair.Item1 == null || connectionPair.Item2 == null ||
                                                connectionPair.Item1.IsAlmostEqualTo(connectionPair.Item2, 0.001))
                                            {
                                                continue;
                                            }

                                            Line line = Line.CreateBound(connectionPair.Item1, connectionPair.Item2);
                                            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                                            ds.SetShape(new List<GeometryObject> { line });
                                            ds.Name = VIRTUAL_JUMP_NAME;

                                            var lineOverride = new OverrideGraphicSettings();
                                            lineOverride.SetProjectionLineColor(new Color(255, 0, 255));
                                            lineOverride.SetProjectionLineWeight(5);

                                            var dashPattern = new FilteredElementCollector(doc)
                                                .OfClass(typeof(LinePatternElement))
                                                .Cast<LinePatternElement>()
                                                .FirstOrDefault(lpe => lpe.Name.Contains("Dash"));

                                            if (dashPattern != null)
                                            {
                                                lineOverride.SetProjectionLinePatternId(dashPattern.Id);
                                            }

                                            activeView.SetElementOverrides(ds.Id, lineOverride);
                                            successfulLines++;
                                            anyLinesVisualized = true;
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"DirectShape creation error: {ex.Message}");
                                            modelLinesFailed = true;
                                        }
                                    }

                                    // Restore original workset if needed
                                    if (doc.IsWorkshared && originalWorksetId != WorksetId.InvalidWorksetId)
                                    {
                                        doc.GetWorksetTable().SetActiveWorksetId(originalWorksetId);
                                    }
                                }
                                else
                                {
                                    // Fallback visualization - additional markers
                                    System.Diagnostics.Debug.WriteLine("Using fallback visualization method (additional markers)");
                                    attemptedFallback = true;

                                    // Create enhanced markers to indicate connections
                                    foreach (var connectionPair in connectionLines)
                                    {
                                        try
                                        {
                                            FamilySymbol sphereSymbol = GetReferencePointSymbol(doc);
                                            if (sphereSymbol != null && !sphereSymbol.IsActive)
                                            {
                                                sphereSymbol.Activate();
                                            }

                                            if (sphereSymbol != null)
                                            {
                                                // Additional markers to indicate connection pairs
                                                FamilyInstance startMarker = doc.Create.NewFamilyInstance(
                                                    connectionPair.Item1,
                                                    sphereSymbol,
                                                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                                FamilyInstance endMarker = doc.Create.NewFamilyInstance(
                                                    connectionPair.Item2,
                                                    sphereSymbol,
                                                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                                // Create a special override for connection points
                                                var connOverride = new OverrideGraphicSettings();
                                                connOverride.SetProjectionLineColor(new Color(255, 0, 255));
                                                connOverride.SetSurfaceForegroundPatternColor(new Color(255, 0, 255));

                                                activeView.SetElementOverrides(startMarker.Id, connOverride);
                                                activeView.SetElementOverrides(endMarker.Id, connOverride);

                                                successfulLines++;
                                                anyLinesVisualized = true;
                                            }
                                        }
                                        catch (Exception fallbackEx)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Fallback visualization error: {fallbackEx.Message}");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception outerEx)
                        {
                            // Handle any unexpected errors in the entire visualization block
                            System.Diagnostics.Debug.WriteLine($"Overall visualization error: {outerEx.Message}");
                            modelLinesFailed = true;
                        }

                        // Update the warning messages to include fallback information
                        if (modelLinesFailed && !anyLinesVisualized)
                        {
                            if (attemptedFallback)
                            {
                                TaskDialog.Show("Warning", "Failed to visualize connections. This may be due to permissions or workset constraints.");
                            }
                            else
                            {
                                TaskDialog.Show("Warning", "Failed to create any virtual jump lines. Only connection points are visualized.");
                            }
                        }
                        else if (modelLinesFailed)
                        {
                            TaskDialog.Show("Warning",
                                $"Some virtual jump lines failed to create. {successfulLines} of {connectionLines.Count} lines were created successfully.");
                        }
                        else if (!anyPointsVisualized && !anyLinesVisualized && connectionLines.Count > 0)
                        {
                            TaskDialog.Show("Warning",
                                "No connection points or lines were visualized. Check the model for details.");
                        }
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    if (tx.HasStarted()) tx.RollBack();

                    // Build a detailed error message including inner exception details if available
                    string detailedMessage = $"Failed to visualize path: {ex.Message}";

                    if (ex.InnerException != null)
                    {
                        detailedMessage += $"\n\nInner Exception: {ex.InnerException.Message}";
                    }

                    // Add more specific diagnostic details based on exception type
                    if (ex is Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        detailedMessage += "\n\nThis appears to be a Revit API operation error. Check if there are active transactions, or if the document is read-only.";
                    }
                    else if (ex is System.NullReferenceException)
                    {
                        detailedMessage += "\n\nA null reference was encountered. This might be due to missing elements or parameters.";
                    }
                    else if (ex is System.ArgumentException)
                    {
                        detailedMessage += "\n\nAn invalid argument was provided to a method.";
                    }

                    // Include diagnostic information about the environment
#if REVIT2024_OR_GREATER
                    detailedMessage += $"\n\nActive view: {activeView.Name} (ID: {activeView.Id.Value})";
#else
                    detailedMessage += $"\n\nActive view: {activeView.Name} (ID: {activeView.Id.IntegerValue})";
#endif
                    detailedMessage += $"\nCable reference: {cableRef}";
                    detailedMessage += $"\nIslands found: {islands.Count}";

                    // Log to debug output
                    System.Diagnostics.Debug.WriteLine($"ERROR in RTS_DiagnoseConnectivity: {detailedMessage}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");

                    message = detailedMessage;
                    TaskDialog dialog = new TaskDialog("Error Details");
                    dialog.MainInstruction = "Failed to visualize path";
                    dialog.MainContent = detailedMessage;
                    dialog.Show();

                    return Result.Failed;
                }
            }

            uidoc.RefreshActiveView();
            return Result.Succeeded;
        }

        private bool IsWorksetEditable(Document doc, WorksetId worksetId)
        {
            if (!doc.IsWorkshared) return true;

#if REVIT2024_OR_GREATER
            // For Revit 2024 and newer, use GetWorksetCheckoutStatus
            WorksetCheckoutStatus status = WorksharingUtils.GetWorksetCheckoutStatus(doc, worksetId);

            // A workset is editable if it's owned by the current user or not owned by anyone.
            return status == WorksetCheckoutStatus.OwnedByCurrentUser || status == WorksetCheckoutStatus.Editable;
#else
            // For Revit 2022 and older, get the Workset element and check its status
            WorksetTable worksetTable = doc.GetWorksetTable();
            //Fix this code later - Element ID passing is not working in 2022.
            //Workset currentWorkset = worksetTable.GetWorkset(worksetId);
            //if (currentWorkset == null) return false;

            // Check the checkout status of the workset element.
            //CheckoutStatus status = WorksharingUtils.GetCheckoutStatus(doc, currentWorkset.Id);

            // It's editable if it's not owned by another user.
            //return status != CheckoutStatus.OwnedByOtherUser;
            return true; // Fallback for older versions
#endif
        }

        private WorksetId GetEditableWorksetId(Document doc)
        {
            if (!doc.IsWorkshared) return WorksetId.InvalidWorksetId;

            WorksetTable worksetTable = doc.GetWorksetTable();
            WorksetId activeId = worksetTable.GetActiveWorksetId();

            // Check if active workset is editable
            if (IsWorksetEditable(doc, activeId))
                return activeId;

            // Find an editable user workset
            FilteredWorksetCollector collector = new FilteredWorksetCollector(doc);
            foreach (Workset workset in collector.OfKind(WorksetKind.UserWorkset))
            {
                if (IsWorksetEditable(doc, workset.Id))
                    return workset.Id;
            }

            return WorksetId.InvalidWorksetId;
        }

        private void ClearAllOverrides(Document doc, View view, bool useOwnTransaction = true)
        {
            var allElementsInView = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .ToElementIds();

            // Find any marker family instances
            var markerInstances = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.FamilyName != null &&
                       (fi.Symbol.FamilyName.Contains("Sphere") ||
                        fi.Symbol.FamilyName.Contains("Point") ||
                        fi.Symbol.FamilyName.Contains("Marker")))
                .Select(fi => fi.Id);

            // Find any DirectShape elements we created
            var directShapes = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(DirectShape))
                .Cast<DirectShape>()
                .Where(ds => ds.Name == VIRTUAL_JUMP_NAME)
                .Select(ds => ds.Id);

            Action performCleanup = () =>
            {
                // Reset all overrides first
                foreach (var id in allElementsInView)
                {
                    view.SetElementOverrides(id, new OverrideGraphicSettings());
                }

                // Delete any marker instances
                if (markerInstances.Any())
                {
                    doc.Delete(markerInstances.ToList());
                }

                // Delete any DirectShape lines with our name
                if (directShapes.Any())
                {
                    doc.Delete(directShapes.ToList());
                }
            };

            if (useOwnTransaction)
            {
                using (var t = new Transaction(doc, "Clear Diagnostic Graphics"))
                {
                    try
                    {
                        t.Start();
                        performCleanup();
                        t.Commit();
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // Transaction already in progress, try running without a new transaction
                        performCleanup();
                    }
                    catch (Exception)
                    {
                        // TaskDialog.Show("Warning", $"Failed to clear graphics: {ex.Message}");
                        if (t.HasStarted() && !t.GetStatus().Equals(TransactionStatus.RolledBack))
                            t.RollBack();
                    }
                }
            }
            else
            {
                // Assume we're already in a transaction
                performCleanup();
            }
        }

        private OverrideGraphicSettings CreateSolidFillOverride(Document doc, Color color)
        {
            var overrideSettings = new OverrideGraphicSettings();
            overrideSettings.SetProjectionLineColor(color);
            overrideSettings.SetSurfaceForegroundPatternColor(color);

            var solidFillPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fpe => fpe.GetFillPattern().IsSolidFill);

            if (solidFillPattern != null)
            {
                overrideSettings.SetSurfaceForegroundPatternId(solidFillPattern.Id);
            }

            return overrideSettings;
        }

        private GraphicsStyle GetOrCreateLineStyle(Document doc, string styleName, Color color)
        {
            var lineCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            Category lineSubCategory = lineCategory.SubCategories.get_Item(styleName);

            if (lineSubCategory != null)
            {
                return lineSubCategory.GetGraphicsStyle(GraphicsStyleType.Projection);
            }

            var newLineStyleCat = doc.Settings.Categories.NewSubcategory(lineCategory, styleName);
            doc.Regenerate();

            newLineStyleCat.LineColor = color;
            newLineStyleCat.SetLineWeight(5, GraphicsStyleType.Projection);

            var dashPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(LinePatternElement))
                .Cast<LinePatternElement>()
                .FirstOrDefault(lpe => lpe.Name.Contains("Dash"));

            if (dashPattern != null)
            {
                newLineStyleCat.SetLinePatternId(dashPattern.Id, GraphicsStyleType.Projection);
            }

            return newLineStyleCat.GetGraphicsStyle(GraphicsStyleType.Projection);
        }

        private FamilySymbol GetReferencePointSymbol(Document doc)
        {
            // Try to find an existing Generic Model point-like family
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs => fs.FamilyName.Contains("Sphere") ||
                                      fs.FamilyName.Contains("Point") ||
                                      fs.FamilyName.Contains("Marker"));

            if (symbol != null)
                return symbol;

            // If no suitable family found, use any generic model family
            symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilySymbol>()
                .FirstOrDefault();

            if (symbol == null)
            {
                TaskDialog.Show("Warning",
                    "Could not find any Generic Model family to use as a marker. " +
                    "Connection points will not be displayed.");
            }

            return symbol;
        }
    }
}