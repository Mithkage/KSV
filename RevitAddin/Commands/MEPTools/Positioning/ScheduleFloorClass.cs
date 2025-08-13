// --- FILE: ScheduleFloorClass.cs ---
//
// Namespace: RTS.Commands
//
// Class: ScheduleFloorClass
//
// Function: This Revit 2024 external command updates the "Schedule Level" parameter of all
//           lighting fixtures in the host model. It identifies the architectural link model
//           specified in the "Profile Settings" for "Floors" and "Structural Slabs". For each fixture, 
//           it finds the closest floor or slab directly beneath it in the relevant linked model. 
//           If no element is directly below, it searches for the nearest one horizontally.
//           The script then finds the host level with the same name as the floor's level
//           and updates the fixture's schedule level parameter accordingly, without changing
//           the fixture's physical location.
//
// Author: Kyle Vorster
// Company: ReTick Solutions Pty Ltd
//
// Log:
// - 2025-08-13: Corrected access to ProfileSettings and Revit Link Instance ID retrieval.
// - 2025-08-13: Updated to use separate profile settings for "Floors" and "Structural Slabs".
// - July 31, 2025: Refactored summary dictionary to a class-level field for consistency.
// - July 31, 2025: Added a fast room check to pre-filter elements, falling back to ray casting for outliers.
// - July 30, 2025: Integrated the ProgressBarWindow for user feedback and cancellation.
//

#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using RTS.UI;
using RTS.Utilities;
#endregion

namespace RTS.Commands.MEPTools.Positioning
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ScheduleFloorClass : IExternalCommand
    {
        // --- Class-level field for summary reporting ---
        private Dictionary<string, int> _updateSummary;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            _updateSummary = new Dictionary<string, int>();

            // --- 1. Retrieve Profile Settings ---
            var settings = RTS_RevitUtils.GetProfileSettings(doc);
            if (settings == null || settings.ModelCategoryMappings == null)
            {
                message = "Profile Settings could not be loaded or are incomplete. Please configure them first.";
                TaskDialog.Show("Error", message);
                return Result.Cancelled;
            }


            // --- 2. Get Links and Collect Floors/Slabs ---
            var allFloors = new List<Floor>();
            var linkDataMap = new Dictionary<string, (RevitLinkInstance, Document)>();

            // Get Architectural Floors
            var archFloorsMapping = settings.ModelCategoryMappings.FirstOrDefault(m => m.CategoryName == "Floors");
            RevitLinkInstance archLinkInstance = null;
            Document archDoc = null;

            if (archFloorsMapping != null && archFloorsMapping.SelectedLink != "<None>")
            {
                string archLinkName = RTS_RevitUtils.ParseLinkName(archFloorsMapping.SelectedLink);
                var (instance, diagMsg) = RTS_RevitUtils.GetLinkInstance(doc, archLinkName);
                if (instance != null)
                {
                    archLinkInstance = instance;
                    archDoc = instance.GetLinkDocument();
                    if (archDoc != null)
                    {
                        linkDataMap[archLinkName] = (instance, archDoc);
                        var archFloors = new FilteredElementCollector(archDoc)
                            .OfClass(typeof(Floor)).WhereElementIsNotElementType().Cast<Floor>()
                            .Where(f => f.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 0)
                            .ToList();
                        allFloors.AddRange(archFloors);
                    }
                }
            }

            // Get Structural Slabs
            var structSlabsMapping = settings.ModelCategoryMappings.FirstOrDefault(m => m.CategoryName == "Structural Slabs");
            if (structSlabsMapping != null && structSlabsMapping.SelectedLink != "<None>")
            {
                string structLinkName = RTS_RevitUtils.ParseLinkName(structSlabsMapping.SelectedLink);
                if (linkDataMap.ContainsKey(structLinkName))
                {
                    // Link already processed, just add structural floors from it
                    var structDoc = linkDataMap[structLinkName].Item2;
                    var structSlabs = new FilteredElementCollector(structDoc)
                        .OfClass(typeof(Floor)).WhereElementIsNotElementType().Cast<Floor>()
                        .Where(f => f.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                        .ToList();
                    allFloors.AddRange(structSlabs);
                }
                else
                {
                    var (instance, diagMsg) = RTS_RevitUtils.GetLinkInstance(doc, structLinkName);
                    if (instance != null)
                    {
                        var structDoc = instance.GetLinkDocument();
                        if (structDoc != null)
                        {
                            linkDataMap[structLinkName] = (instance, structDoc);
                            var structSlabs = new FilteredElementCollector(structDoc)
                                .OfClass(typeof(Floor)).WhereElementIsNotElementType().Cast<Floor>()
                                .Where(f => f.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                                .ToList();
                            allFloors.AddRange(structSlabs);
                        }
                    }
                }
            }

            if (!allFloors.Any())
            {
                TaskDialog.Show("Warning", "No floors or structural slabs were found in the configured linked models.");
                return Result.Succeeded;
            }

            // --- 3. Prompt user to select an element and collect all elements of that category ---
            Reference selectedRef;
            try
            {
                selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, "Select an element to determine the category to process.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            if (selectedRef == null)
            {
                return Result.Cancelled;
            }

            Element selectedElement = doc.GetElement(selectedRef.ElementId);
            Category selectedCategory = selectedElement.Category;

            if (selectedCategory == null)
            {
                message = "The selected element does not have a valid category.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // --- 3a. Validate selected element ---
            if (selectedElement.Location as LocationPoint == null)
            {
                message = $"The selected element category '{selectedCategory.Name}' is not supported because its elements do not have a point-based location.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            Parameter scheduleLevelParam = selectedElement.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
            if (scheduleLevelParam == null)
            {
                message = $"The selected element category '{selectedCategory.Name}' is not supported because its elements do not have a 'Schedule Level' parameter.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // --- 3b. Collect all elements of the selected category ---
            var elementsToProcess = new FilteredElementCollector(doc)
                .OfCategoryId(selectedCategory.Id)
                .WhereElementIsNotElementType()
                .ToList()
                .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                .ToList();

            var hostLevels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToDictionary(l => l.Name, l => l);

            View3D view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view == null)
            {
                message = "A 3D view is required for ray casting. Please create one.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
            ReferenceIntersector intersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Floors), FindReferenceTarget.Element, view)
            {
                FindReferencesInRevitLinks = true
            };

            Floor cachedFloor = null;
            XYZ lastFixtureLocation = null;
            const double cacheProximityThreshold = 15.0;

            var progressBar = new ProgressBarWindow();
            progressBar.Show();
            int processedCount = 0;

            // --- 4. Group elements by room (using architectural link) ---
            var elementsByRoom = new Dictionary<ElementId, List<Element>>();
            var elementsNotInRooms = new List<Element>();
            if (archDoc != null && archLinkInstance != null)
            {
                Transform linkTransformInverse = archLinkInstance.GetTotalTransform().Inverse;
                foreach (var elem in elementsToProcess)
                {
                    LocationPoint locPoint = elem.Location as LocationPoint;
                    if (locPoint == null) { elementsNotInRooms.Add(elem); continue; }

                    Room room = archDoc.GetRoomAtPoint(linkTransformInverse.OfPoint(locPoint.Point));
                    if (room != null)
                    {
                        if (!elementsByRoom.ContainsKey(room.Id))
                        {
                            elementsByRoom[room.Id] = new List<Element>();
                        }
                        elementsByRoom[room.Id].Add(elem);
                    }
                    else
                    {
                        elementsNotInRooms.Add(elem);
                    }
                }
            }
            else
            {
                elementsNotInRooms.AddRange(elementsToProcess); // No arch link, process all as "not in room"
            }


            // --- 5. Start Transaction and Process Fixtures ---
            using (Transaction trans = new Transaction(doc, $"Update Schedule Levels for {selectedCategory.Name}"))
            {
                try
                {
                    trans.Start();
                    int roomCount = 0;
                    int totalRooms = elementsByRoom.Count;

                    // Process elements grouped by room
                    foreach (var roomGroup in elementsByRoom)
                    {
                        roomCount++;
                        Room room = archDoc.GetElement(roomGroup.Key) as Room;
                        progressBar.UpdateRoomStatus(room?.Name ?? "Unnamed Room", roomCount, totalRooms);

                        foreach (var fixture in roomGroup.Value)
                        {
                            if (progressBar.IsCancellationPending) { trans.RollBack(); progressBar.Close(); return Result.Cancelled; }
                            processedCount++;
                            progressBar.UpdateProgress(processedCount, elementsToProcess.Count);
                            ProcessFixture(fixture, doc, archDoc, archLinkInstance, hostLevels, allFloors, intersector, ref cachedFloor, ref lastFixtureLocation, cacheProximityThreshold);
                        }
                    }

                    // Process elements that were not in any room
                    if (elementsNotInRooms.Any())
                    {
                        progressBar.UpdateRoomStatus("Processing elements not in rooms...", 0, 0);
                        foreach (var fixture in elementsNotInRooms)
                        {
                            if (progressBar.IsCancellationPending) { trans.RollBack(); progressBar.Close(); return Result.Cancelled; }
                            processedCount++;
                            progressBar.UpdateProgress(processedCount, elementsToProcess.Count);
                            ProcessFixture(fixture, doc, archDoc, archLinkInstance, hostLevels, allFloors, intersector, ref cachedFloor, ref lastFixtureLocation, cacheProximityThreshold);
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = "An unexpected error occurred: " + ex.Message;
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }
                finally
                {
                    progressBar.Close();
                }
            }

            // --- 6. Display Summary Report ---
            if (_updateSummary.Any())
            {
                StringBuilder summaryReport = new StringBuilder();
                summaryReport.AppendLine($"Successfully updated schedule levels for the following {selectedCategory.Name} types:");
                summaryReport.AppendLine();
                foreach (var entry in _updateSummary)
                {
                    summaryReport.AppendLine($"  - {entry.Key}: {entry.Value} element(s) updated.");
                }
                TaskDialog.Show("Update Complete", summaryReport.ToString());
            }
            else
            {
                TaskDialog.Show("No Changes", $"No {selectedCategory.Name} elements required a schedule level update.");
            }

            return Result.Succeeded;
        }

        #region Helper Methods

        private void ProcessFixture(Element fixture, Document doc, Document archDoc, RevitLinkInstance archLinkInstance, Dictionary<string, Level> hostLevels, List<Floor> allFloors, ReferenceIntersector intersector, ref Floor cachedFloor, ref XYZ lastFixtureLocation, double cacheProximityThreshold)
        {
            LocationPoint locPoint = fixture.Location as LocationPoint;
            if (locPoint == null) return;

            XYZ fixtureLocation = locPoint.Point;
            Level targetLevel = null;
            bool needsRayCast = true;

            // --- OPTIMIZATION: Fast Room Check First (only if arch link is available) ---
            if (archDoc != null && archLinkInstance != null)
            {
                Room room = archDoc.GetRoomAtPoint(archLinkInstance.GetTotalTransform().Inverse.OfPoint(fixtureLocation));
                if (room != null)
                {
                    Level roomLevel = archDoc.GetElement(room.LevelId) as Level;
                    Parameter scheduleLevelParam = fixture.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                    if (roomLevel != null && scheduleLevelParam != null && scheduleLevelParam.HasValue)
                    {
                        Level currentScheduleLevel = doc.GetElement(scheduleLevelParam.AsElementId()) as Level;
                        if (currentScheduleLevel != null && currentScheduleLevel.Name == roomLevel.Name)
                        {
                            needsRayCast = false; // Already correct, skip the expensive check
                        }
                    }
                }
            }

            // --- Fallback to Ray Casting if needed ---
            if (needsRayCast)
            {
                Floor closestFloor = null;

                if (cachedFloor != null && lastFixtureLocation != null && fixtureLocation.DistanceTo(lastFixtureLocation) < cacheProximityThreshold)
                {
                    if (ValidateCachedFloor(fixtureLocation, cachedFloor, intersector))
                    {
                        closestFloor = cachedFloor;
                    }
                }

                if (closestFloor == null)
                {
                    closestFloor = FindClosestFloor(fixtureLocation, allFloors, intersector, doc);
                }

                lastFixtureLocation = fixtureLocation;
                cachedFloor = closestFloor;

                if (closestFloor != null)
                {
                    Document floorDoc = closestFloor.Document;
                    Level linkedLevel = floorDoc.GetElement(closestFloor.LevelId) as Level;
                    if (linkedLevel != null && hostLevels.TryGetValue(linkedLevel.Name, out Level hostLevel))
                    {
                        targetLevel = hostLevel;
                    }
                }
            }

            if (targetLevel != null)
            {
                Parameter scheduleLevelParam = fixture.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
                {
                    if (scheduleLevelParam.AsElementId() != targetLevel.Id)
                    {
                        scheduleLevelParam.Set(targetLevel.Id);
                        string fixtureTypeName = fixture.Name;
                        if (!_updateSummary.ContainsKey(fixtureTypeName))
                        {
                            _updateSummary[fixtureTypeName] = 0;
                        }
                        _updateSummary[fixtureTypeName]++;
                    }
                }
            }
        }

        private Floor FindClosestFloor(XYZ point, List<Floor> allFloors, ReferenceIntersector intersector, Document hostDoc)
        {
            ReferenceWithContext refWithContext = intersector.FindNearest(point, XYZ.BasisZ.Negate());

            if (refWithContext != null)
            {
                Reference reference = refWithContext.GetReference();
                ElementId linkedElementId = reference.LinkedElementId;
                if (linkedElementId != ElementId.InvalidElementId)
                {
                    RevitLinkInstance linkInstance = hostDoc.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance != null)
                    {
                        Document linkDoc = linkInstance.GetLinkDocument();
                        return linkDoc?.GetElement(linkedElementId) as Floor;
                    }
                }
            }

            // Fallback to horizontal search if ray cast fails
            Floor nearestHorizontalFloor = null;
            double minHorizontalDistance = double.MaxValue;

            var linkInstances = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                .ToDictionary(li => li.GetLinkDocument()?.Title, li => li);

            foreach (var floor in allFloors)
            {
                Document floorDoc = floor.Document;
                if (!linkInstances.TryGetValue(floorDoc.Title, out RevitLinkInstance linkInstance)) continue;

                Transform linkTransform = linkInstance.GetTotalTransform();
                GeometryElement geoElem = floor.get_Geometry(new Options());

                foreach (GeometryObject geoObj in geoElem)
                {
                    Solid solid = geoObj as Solid;
                    if (solid == null || solid.Faces.Size == 0) continue;

                    Solid transformedSolid = SolidUtils.CreateTransformed(solid, linkTransform);
                    try
                    {
                        foreach (Face face in transformedSolid.Faces)
                        {
                            IntersectionResult result = face.Project(point);
                            if (result == null) continue;

                            XYZ closestPoint = result.XYZPoint;
                            double horizontalDistance = new XYZ(point.X - closestPoint.X, point.Y - closestPoint.Y, 0).GetLength();

                            if (horizontalDistance < minHorizontalDistance)
                            {
                                minHorizontalDistance = horizontalDistance;
                                nearestHorizontalFloor = floor;
                            }
                        }
                    }
                    catch { /* Ignore problematic faces */ }
                }
            }
            return nearestHorizontalFloor;
        }

        private bool ValidateCachedFloor(XYZ point, Floor cachedFloor, ReferenceIntersector intersector)
        {
            if (cachedFloor == null) return false;

            ReferenceWithContext refWithContext = intersector.FindNearest(point, XYZ.BasisZ.Negate());

            if (refWithContext != null)
            {
                Reference reference = refWithContext.GetReference();
                return reference.LinkedElementId == cachedFloor.Id;
            }
            return false;
        }
        #endregion
    }
}