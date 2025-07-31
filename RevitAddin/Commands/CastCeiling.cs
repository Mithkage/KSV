// --- FILE: CastCeiling.cs ---
//
// Namespace: RTS.Commands
//
// Class: CastCeilingClass
//
// Function: This Revit 2024 external command adjusts the vertical position of MEP elements.
//           After a user selects an initial element, the script finds all other elements of the
//           same type. For each element, it ray casts upwards to find the nearest ceiling or slab
//           from the architectural links specified in the "Profile Settings". It then moves the
//           element to align its Reference Level with the closest surface.
//
// Author: Kyle Vorster
// Company: ReTick Solutions Pty Ltd
//
// Log:
// - July 31, 2025: Added preprocessor directives for Revit 2022/2024 compatibility to remove warnings.
// - July 31, 2025: Corrected method signatures to resolve variable scope compiler errors.
// - July 31, 2025: Added user option for Fast vs. Accurate processing.
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
using RTS.UI; // Required for the ProgressBarWindow
#endregion

namespace RTS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CastCeilingClass : IExternalCommand
    {
        // --- Class-level fields for summary reporting ---
        private Dictionary<string, int> _movedElementsSummary;
        private Dictionary<string, int> _failedElementsSummary;
        private List<long> _outlierElements;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Initialize the summary collections for this run
            _movedElementsSummary = new Dictionary<string, int>();
            _failedElementsSummary = new Dictionary<string, int>();
            _outlierElements = new List<long>();

            // --- 1. Retrieve Profile Settings and Links ---
            var settings = RTS_RevitUtils.GetProfileSettings(doc);
            if (settings == null)
            {
                message = "Profile Settings could not be loaded. Please configure them first.";
                TaskDialog.Show("Error", message);
                return Result.Cancelled;
            }

            var (ceilingLink, slabLink, diagnosticMessage) = RTS_RevitUtils.GetLinkInstances(doc, settings);

            // --- 2. User Selections ---
            IList<Reference> selectedRefs;
            try
            {
                selectedRefs = uiDoc.Selection.PickObjects(ObjectType.Element, new MepElementSelectionFilter(), "Select one or more electrical or MEP elements to process.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            if (selectedRefs == null || selectedRefs.Count == 0) return Result.Cancelled;

            // Ask for surface type
            TaskDialog surfaceDialog = new TaskDialog("Select Surface Types");
            surfaceDialog.MainInstruction = "Which surfaces should the elements be cast to?";
            surfaceDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Ceilings Only");
            surfaceDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Slabs (Floors) Only");
            surfaceDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Both Ceilings and Slabs");
            surfaceDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            TaskDialogResult surfaceResult = surfaceDialog.Show();

            if (surfaceResult == TaskDialogResult.CommandLink1) { if (ceilingLink == null) { TaskDialog.Show("Error", "The 'Ceilings' link is not defined in Profile Settings."); return Result.Cancelled; } slabLink = null; }
            else if (surfaceResult == TaskDialogResult.CommandLink2) { if (slabLink == null) { TaskDialog.Show("Error", "The 'Slabs' link is not defined in Profile Settings."); return Result.Cancelled; } ceilingLink = null; }
            else if (surfaceResult == TaskDialogResult.CommandLink3) { if (ceilingLink == null && slabLink == null) { TaskDialog.Show("Error", "Neither the 'Ceilings' nor the 'Slabs' link is defined in Profile Settings."); return Result.Cancelled; } }
            else { return Result.Cancelled; }

            // --- 2a. Check if the relevant links are Room Bounding ---
            if (!CheckRoomBoundingLinks(doc, ceilingLink, slabLink, settings))
            {
                return Result.Cancelled; // User chose not to proceed
            }

            // Ask for processing method
            TaskDialog methodDialog = new TaskDialog("Select Processing Method");
            methodDialog.MainInstruction = "Choose a processing method";
            methodDialog.MainContent = "Fast Processing is quicker but assumes flat ceilings within each room. Accurate Processing is slower but correctly handles sloped or stepped ceilings.";
            methodDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Fast Processing (Recommended for simple models)");
            methodDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Accurate Processing (Recommended for complex ceilings)");
            methodDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            TaskDialogResult methodResult = methodDialog.Show();

            bool useFastProcessing = true;
            if (methodResult == TaskDialogResult.CommandLink1) { useFastProcessing = true; }
            else if (methodResult == TaskDialogResult.CommandLink2) { useFastProcessing = false; }
            else { return Result.Cancelled; }


            // --- 3. Collect and Group Elements ---
            var selectedTypeIds = selectedRefs.Select(r => doc.GetElement(r.ElementId).GetTypeId()).Where(id => id != ElementId.InvalidElementId).Distinct().ToList();
            if (!selectedTypeIds.Any()) { message = "The selected elements do not have valid types."; TaskDialog.Show("Error", message); return Result.Failed; }

            var elementsToProcess = new FilteredElementCollector(doc).WhereElementIsNotElementType().Where(e => selectedTypeIds.Contains(e.GetTypeId())).ToList();
            if (!elementsToProcess.Any()) { TaskDialog.Show("Information", "No elements of the selected types were found to process."); return Result.Succeeded; }

            // --- 4. Setup for Ray Casting and Grouping ---
            View3D view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view == null) { message = "A 3D view is required. Please create one."; TaskDialog.Show("Error", message); return Result.Failed; }

            var ceilingIntersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Ceilings), FindReferenceTarget.Face, view) { FindReferencesInRevitLinks = true };
            var slabIntersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Floors), FindReferenceTarget.Face, view) { FindReferencesInRevitLinks = true };

            const double maxMoveDistance = 6.56168; // 2 meters in feet

            var elementsByRoom = new Dictionary<ElementId, List<Element>>();
            var elementsNotInRooms = new List<Element>();
            RevitLinkInstance roomLink = ceilingLink ?? slabLink;
            Document roomLinkDoc = roomLink?.GetLinkDocument();
            Transform linkTransform = roomLink?.GetTotalTransform();
            Transform linkTransformInverse = linkTransform?.Inverse;

            foreach (var elem in elementsToProcess)
            {
                LocationPoint locPoint = elem.Location as LocationPoint;
                if (locPoint == null || roomLinkDoc == null) { elementsNotInRooms.Add(elem); continue; }
                Room room = roomLinkDoc.GetRoomAtPoint(linkTransformInverse.OfPoint(locPoint.Point));
                if (room != null)
                {
                    if (!elementsByRoom.ContainsKey(room.Id)) elementsByRoom[room.Id] = new List<Element>();
                    elementsByRoom[room.Id].Add(elem);
                }
                else { elementsNotInRooms.Add(elem); }
            }

            // --- 5. Process Elements ---
            var progressBar = new ProgressBarWindow();
            progressBar.Show();
            int processedCount = 0;

            using (Transaction trans = new Transaction(doc, "Cast Elements to Ceiling or Slab"))
            {
                try
                {
                    trans.Start();
                    int roomCount = 0;
                    int totalRooms = elementsByRoom.Count;

                    foreach (var roomGroup in elementsByRoom)
                    {
                        if (progressBar.IsCancellationPending) { trans.RollBack(); progressBar.Close(); return Result.Cancelled; }
                        roomCount++;
                        Room room = roomLinkDoc.GetElement(roomGroup.Key) as Room;
                        progressBar.UpdateRoomStatus(room?.Name ?? "Unnamed Room", roomCount, totalRooms);

                        var roomElements = roomGroup.Value;

                        if (useFastProcessing)
                        {
                            LocationPoint roomCenter = room.Location as LocationPoint;
                            if (roomCenter == null) { ProcessIndividually(roomElements, ceilingIntersector, slabIntersector, ceilingLink, slabLink, maxMoveDistance, doc, progressBar, ref processedCount, elementsToProcess.Count); continue; }

                            ReferenceWithContext hit = FindClosestSurface(linkTransform.OfPoint(roomCenter.Point), ceilingIntersector, slabIntersector, ceilingLink, slabLink);
                            if (hit != null)
                            {
                                double targetZ = hit.GetReference().GlobalPoint.Z;
                                foreach (var elem in roomElements)
                                {
                                    ProcessElement(elem, targetZ, maxMoveDistance, doc);
                                    processedCount++;
                                    progressBar.UpdateProgress(processedCount, elementsToProcess.Count);
                                }
                            }
                            else { ProcessIndividually(roomElements, ceilingIntersector, slabIntersector, ceilingLink, slabLink, maxMoveDistance, doc, progressBar, ref processedCount, elementsToProcess.Count); }
                        }
                        else
                        {
                            BoundingBoxXYZ roomBB = room.get_BoundingBox(null);
                            if (roomBB == null) { ProcessIndividually(roomElements, ceilingIntersector, slabIntersector, ceilingLink, slabLink, maxMoveDistance, doc, progressBar, ref processedCount, elementsToProcess.Count); continue; }

                            roomBB.Transform = linkTransform;
                            XYZ p1 = new XYZ(roomBB.Min.X, roomBB.Min.Y, roomBB.Min.Z);
                            XYZ p2 = new XYZ(roomBB.Max.X, roomBB.Max.Y, roomBB.Min.Z);

                            ReferenceWithContext hit1 = FindClosestSurface(p1, ceilingIntersector, slabIntersector, ceilingLink, slabLink);
                            ReferenceWithContext hit2 = FindClosestSurface(p2, ceilingIntersector, slabIntersector, ceilingLink, slabLink);

                            if (hit1 != null && hit2 != null && Math.Abs(hit1.GetReference().GlobalPoint.Z - hit2.GetReference().GlobalPoint.Z) < 0.01)
                            {
                                double targetZ = hit1.GetReference().GlobalPoint.Z;
                                foreach (var elem in roomElements)
                                {
                                    ProcessElement(elem, targetZ, maxMoveDistance, doc);
                                    processedCount++;
                                    progressBar.UpdateProgress(processedCount, elementsToProcess.Count);
                                }
                            }
                            else { ProcessIndividually(roomElements, ceilingIntersector, slabIntersector, ceilingLink, slabLink, maxMoveDistance, doc, progressBar, ref processedCount, elementsToProcess.Count); }
                        }
                    }

                    if (elementsNotInRooms.Any())
                    {
                        progressBar.UpdateRoomStatus("Processing elements not in rooms...", 0, 0);
                        var sortedNotInRooms = elementsNotInRooms
                            .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                            .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                            .ThenBy(e => (e.Location as LocationPoint)?.Point.Z ?? 0)
                            .ToList();

                        ProcessIndividually(sortedNotInRooms, ceilingIntersector, slabIntersector, ceilingLink, slabLink, maxMoveDistance, doc, progressBar, ref processedCount, elementsToProcess.Count);
                    }

                    trans.Commit();
                }
                catch (Exception ex) { trans.RollBack(); message = "An error occurred: " + ex.Message; TaskDialog.Show("Error", message); return Result.Failed; }
                finally { progressBar.Close(); }
            }

            DisplaySummary();
            return Result.Succeeded;
        }

        #region Processing Methods

        private void ProcessElement(Element elem, double targetZ, double maxMove, Document doc)
        {
            double currentRefZ = GetElementReferenceElevation(elem, doc);
            if (currentRefZ == double.MinValue) { LogFailure(elem.Name); return; }

            double moveDistance = targetZ - currentRefZ;
            if (Math.Abs(moveDistance) > maxMove)
            {
#if REVIT2024_OR_GREATER
                _outlierElements.Add(elem.Id.Value);
#else
                _outlierElements.Add(elem.Id.IntegerValue);
#endif
                return;
            }

            ElementTransformUtils.MoveElement(doc, elem.Id, new XYZ(0, 0, moveDistance));
            LogSuccess(elem.Name);
        }

        private void ProcessIndividually(List<Element> elements, ReferenceIntersector cInt, ReferenceIntersector sInt, RevitLinkInstance cLink, RevitLinkInstance sLink, double maxMove, Document doc, ProgressBarWindow progressBar, ref int processedCount, int totalCount)
        {
            ReferenceWithContext cachedHit = null;
            XYZ lastElementLocation = null;
            const double cacheProximityThreshold = 15.0;

            foreach (var elem in elements)
            {
                if (progressBar.IsCancellationPending) return;

                LocationPoint locPoint = elem.Location as LocationPoint;
                if (locPoint == null) { LogFailure(elem.Name); continue; }

                XYZ currentLocation = locPoint.Point;
                ReferenceWithContext targetHit = null;
                XYZ targetPoint = null;

                if (cachedHit != null && lastElementLocation != null && currentLocation.DistanceTo(lastElementLocation) < cacheProximityThreshold)
                {
                    if (ValidateCacheByBoundingBox(currentLocation, cachedHit, doc))
                    {
                        targetHit = cachedHit;
                        targetPoint = new XYZ(currentLocation.X, currentLocation.Y, cachedHit.GetReference().GlobalPoint.Z);
                    }
                }

                if (targetHit == null)
                {
                    targetHit = FindClosestSurface(currentLocation, cInt, sInt, cLink, sLink);
                    if (targetHit != null)
                    {
                        targetPoint = targetHit.GetReference().GlobalPoint;
                    }
                }

                lastElementLocation = currentLocation;
                cachedHit = targetHit;

                if (targetHit == null) { LogFailure(elem.Name); continue; }

                ProcessElement(elem, targetPoint.Z, maxMove, doc);
                processedCount++;
                progressBar.UpdateProgress(processedCount, totalCount);
            }
        }

        #endregion

        #region Helper Methods

        private double GetElementReferenceElevation(Element elem, Document doc)
        {
            if (elem is FamilyInstance fi)
            {
                Level level = doc.GetElement(fi.LevelId) as Level;
                Parameter elevationParam = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                if (level != null && elevationParam != null && elevationParam.HasValue)
                {
                    return level.Elevation + elevationParam.AsDouble();
                }
            }
            LocationPoint locPoint = elem.Location as LocationPoint;
            return locPoint?.Point.Z ?? double.MinValue;
        }

        private ReferenceWithContext FindClosestSurface(XYZ point, ReferenceIntersector ceilingIntersector, ReferenceIntersector slabIntersector, RevitLinkInstance ceilingLink, RevitLinkInstance slabLink)
        {
            XYZ startPoint = point - new XYZ(0, 0, 3.28084);
            ReferenceWithContext ceilingHit = null, slabHit = null;
            if (ceilingLink != null)
            {
                ceilingHit = ceilingIntersector.FindNearest(startPoint, XYZ.BasisZ);
                if (ceilingHit != null && ceilingHit.GetReference().ElementId != ceilingLink.Id) ceilingHit = null;
            }
            if (slabLink != null)
            {
                slabHit = slabIntersector.FindNearest(startPoint, XYZ.BasisZ);
                if (slabHit != null && slabHit.GetReference().ElementId != slabLink.Id) slabHit = null;
            }
            if (ceilingHit != null && slabHit != null) return (ceilingHit.Proximity < slabHit.Proximity) ? ceilingHit : slabHit;
            return ceilingHit ?? slabHit;
        }

        private bool ValidateCacheByBoundingBox(XYZ point, ReferenceWithContext cachedHit, Document doc)
        {
            if (cachedHit == null) return false;

            var linkInstance = doc.GetElement(cachedHit.GetReference().ElementId) as RevitLinkInstance;
            if (linkInstance == null) return false;

            var linkDoc = linkInstance.GetLinkDocument();
            if (linkDoc == null) return false;

            var linkedElement = linkDoc.GetElement(cachedHit.GetReference().LinkedElementId);
            if (linkedElement == null) return false;

            BoundingBoxXYZ bb = linkedElement.get_BoundingBox(null);
            if (bb == null) return false;

            Transform transform = linkInstance.GetTotalTransform();
            bb.Transform = transform;

            return point.X >= bb.Min.X && point.X <= bb.Max.X && point.Y >= bb.Min.Y && point.Y <= bb.Max.Y;
        }

        private bool CheckRoomBoundingLinks(Document doc, RevitLinkInstance ceilingLink, RevitLinkInstance slabLink, ProfileSettings settings)
        {
            List<string> nonBoundingLinks = new List<string>();

            if (ceilingLink != null)
            {
                var linkType = doc.GetElement(ceilingLink.GetTypeId()) as RevitLinkType;
                if (linkType != null && linkType.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).AsInteger() == 0)
                {
                    nonBoundingLinks.Add($"Ceilings link: '{RTS_RevitUtils.ParseLinkName(settings.CeilingsLink)}'");
                }
            }
            if (slabLink != null)
            {
                var linkType = doc.GetElement(slabLink.GetTypeId()) as RevitLinkType;
                if (linkType != null && linkType.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).AsInteger() == 0)
                {
                    nonBoundingLinks.Add($"Slabs link: '{RTS_RevitUtils.ParseLinkName(settings.SlabsLink)}'");
                }
            }

            if (nonBoundingLinks.Any())
            {
                TaskDialog warningDialog = new TaskDialog("Room Bounding Warning");
                warningDialog.MainInstruction = "One or more links are not Room Bounding";
                warningDialog.MainContent = "The following link(s) are not set to 'Room Bounding'. This will prevent the high-speed room processing from working correctly, and the script will be slower.\n\n"
                                            + string.Join("\n", nonBoundingLinks)
                                            + "\n\nDo you want to continue?";
                warningDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                warningDialog.DefaultButton = TaskDialogResult.No;

                return warningDialog.Show() == TaskDialogResult.Yes;
            }

            return true;
        }

        private void LogSuccess(string typeName)
        {
            if (!_movedElementsSummary.ContainsKey(typeName)) _movedElementsSummary[typeName] = 0;
            _movedElementsSummary[typeName]++;
        }

        private void LogFailure(string typeName)
        {
            if (!_failedElementsSummary.ContainsKey(typeName)) _failedElementsSummary[typeName] = 0;
            _failedElementsSummary[typeName]++;
        }

        private void DisplaySummary()
        {
            StringBuilder sb = new StringBuilder();
            if (_movedElementsSummary.Any())
            {
                sb.AppendLine("Operation complete. The following elements were moved:");
                foreach (var entry in _movedElementsSummary) sb.AppendLine($"  - {entry.Key}: {entry.Value} element(s)");
            }
            else sb.AppendLine("No elements were moved.");

            if (_failedElementsSummary.Any())
            {
                sb.AppendLine("\nThe following elements could not be re-located (no ceiling or slab found):");
                foreach (var entry in _failedElementsSummary) sb.AppendLine($"  - {entry.Key}: {entry.Value} element(s)");
            }

            if (_outlierElements.Any())
            {
                sb.AppendLine("\nThe following elements were not moved as the required correction exceeded 2 meters. Please review these outliers:");
                sb.AppendLine("Element IDs: " + string.Join(", ", _outlierElements));
            }

            TaskDialog.Show("Summary", sb.ToString());
        }

        #endregion

        #region Selection Filter
        public class MepElementSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category == null) return false;

#if REVIT2024_OR_GREATER
                long catId = elem.Category.Id.Value;
#else
                long catId = elem.Category.Id.IntegerValue;
#endif

                return catId == (long)BuiltInCategory.OST_LightingFixtures || catId == (long)BuiltInCategory.OST_ElectricalFixtures || catId == (long)BuiltInCategory.OST_ElectricalEquipment || catId == (long)BuiltInCategory.OST_CommunicationDevices || catId == (long)BuiltInCategory.OST_FireAlarmDevices || catId == (long)BuiltInCategory.OST_SecurityDevices || catId == (long)BuiltInCategory.OST_MechanicalEquipment || catId == (long)BuiltInCategory.OST_Sprinklers || catId == (long)BuiltInCategory.OST_PlumbingFixtures;
            }
            public bool AllowReference(Reference reference, XYZ position) { return false; }
        }
        #endregion
    }
}
