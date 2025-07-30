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
// - July 30, 2025: Added a user prompt to select which surfaces to check against (ceilings, slabs, or both).
// - July 30, 2025: Updated logic to move the element's Reference Level to the target surface.
// - July 30, 2025: Added a 3-meter vertical move threshold to prevent moving outlier elements.
//

#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // --- 1. Retrieve Profile Settings and Links using the Utility Class ---
            var settings = RTS_RevitUtils.GetProfileSettings(doc);
            if (settings == null)
            {
                message = "Profile Settings could not be loaded. Please configure them first.";
                TaskDialog.Show("Error", message);
                return Result.Cancelled;
            }

            var (ceilingLink, slabLink, diagnosticMessage) = RTS_RevitUtils.GetLinkInstances(doc, settings);

            // --- 2. User Selection of an Element ---
            Element selectedElement;
            try
            {
                Reference selectedRef = uiDoc.Selection.PickObject(ObjectType.Element, new MepElementSelectionFilter(), "Select an electrical or MEP element to process.");
                selectedElement = doc.GetElement(selectedRef.ElementId);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            if (selectedElement == null) return Result.Cancelled;

            // --- 2a. Ask user for search preference ---
            TaskDialog mainDialog = new TaskDialog("Select Surface Types");
            mainDialog.MainInstruction = "Which surfaces should the elements be cast to?";
            mainDialog.MainContent = "Choose whether to search for ceilings, slabs (floors), or both. The script will use the links defined in your Profile Settings.";
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Ceilings Only");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Slabs (Floors) Only");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Both Ceilings and Slabs");
            mainDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            mainDialog.DefaultButton = TaskDialogResult.CommandLink3;

            TaskDialogResult tResult = mainDialog.Show();

            if (tResult == TaskDialogResult.CommandLink1) // Ceilings Only
            {
                if (ceilingLink == null)
                {
                    TaskDialog.Show("Error", "The 'Ceilings' link is not defined in Profile Settings. Please set it before running this command.");
                    return Result.Cancelled;
                }
                slabLink = null; // Nullify slab link to ignore it in the search
            }
            else if (tResult == TaskDialogResult.CommandLink2) // Slabs Only
            {
                if (slabLink == null)
                {
                    TaskDialog.Show("Error", "The 'Slabs' link is not defined in Profile Settings. Please set it before running this command.");
                    return Result.Cancelled;
                }
                ceilingLink = null; // Nullify ceiling link to ignore it in the search
            }
            else if (tResult == TaskDialogResult.CommandLink3) // Both
            {
                if (ceilingLink == null && slabLink == null)
                {
                    TaskDialog.Show("Error", "Neither the 'Ceilings' nor the 'Slabs' link is defined in Profile Settings.");
                    return Result.Cancelled;
                }
            }
            else // User cancelled
            {
                return Result.Cancelled;
            }


            // --- 3. Collect All Elements of the Same Type ---
            ElementId typeId = selectedElement.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                message = "The selected element does not have a valid type.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            var elementsToProcess = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e => e.GetTypeId() == typeId)
                .ToList()
                .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                .ToList();

            if (!elementsToProcess.Any())
            {
                TaskDialog.Show("Information", "No other elements of the selected type were found to process.");
                return Result.Succeeded;
            }

            // --- 4. Setup for Ray Casting ---
            View3D view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view == null)
            {
                message = "A 3D view is required to perform this operation. Please create one if it doesn't exist.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            var ceilingIntersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Ceilings), FindReferenceTarget.Face, view) { FindReferencesInRevitLinks = true };
            var slabIntersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Floors), FindReferenceTarget.Face, view) { FindReferencesInRevitLinks = true };

            var movedElementsSummary = new Dictionary<string, int>();
            var failedElementsSummary = new Dictionary<string, int>();
            var outlierElements = new List<long>();

            ReferenceWithContext cachedHit = null;
            XYZ lastElementLocation = null;
            const double cacheProximityThreshold = 15.0;
            const double maxMoveDistance = 9.84252; // 3 meters in feet

            // --- 5. Initialize and Show Progress Bar ---
            var progressBar = new ProgressBarWindow();
            progressBar.Show();
            int processedCount = 0;

            // --- 6. Process Elements ---
            using (Transaction trans = new Transaction(doc, "Cast Elements to Ceiling or Slab"))
            {
                try
                {
                    trans.Start();

                    foreach (var elem in elementsToProcess)
                    {
                        if (progressBar.IsCancellationPending)
                        {
                            trans.RollBack();
                            progressBar.Close();
                            return Result.Cancelled;
                        }

                        processedCount++;
                        progressBar.UpdateProgress(processedCount, elementsToProcess.Count);

                        XYZ currentLocation = (elem.Location as LocationPoint)?.Point;
                        if (currentLocation == null)
                        {
                            LogFailure(failedElementsSummary, elem.Name);
                            continue;
                        }

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
                            targetHit = FindClosestSurface(currentLocation, ceilingIntersector, slabIntersector, ceilingLink, slabLink);
                            if (targetHit != null)
                            {
                                targetPoint = targetHit.GetReference().GlobalPoint;
                            }
                        }

                        lastElementLocation = currentLocation;
                        cachedHit = targetHit;

                        if (targetHit == null)
                        {
                            LogFailure(failedElementsSummary, elem.Name);
                            continue;
                        }

                        double currentReferenceElevation = GetElementReferenceElevation(elem, doc);
                        if (currentReferenceElevation == double.MinValue)
                        {
                            LogFailure(failedElementsSummary, elem.Name);
                            continue;
                        }

                        XYZ moveVector = new XYZ(0, 0, targetPoint.Z - currentReferenceElevation);

                        if (Math.Abs(moveVector.Z) > maxMoveDistance)
                        {
                            outlierElements.Add(elem.Id.IntegerValue);
                            continue;
                        }

                        ElementTransformUtils.MoveElement(doc, elem.Id, moveVector);

                        LogSuccess(movedElementsSummary, elem.Name);
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = "An error occurred during the operation: " + ex.Message;
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }
                finally
                {
                    progressBar.Close();
                }
            }

            // --- 7. Display Summary Report ---
            DisplaySummary(movedElementsSummary, failedElementsSummary, outlierElements);
            return Result.Succeeded;
        }

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
            if (locPoint != null)
            {
                return locPoint.Point.Z;
            }

            return double.MinValue;
        }

        private ReferenceWithContext FindClosestSurface(XYZ point, ReferenceIntersector ceilingIntersector, ReferenceIntersector slabIntersector, RevitLinkInstance ceilingLink, RevitLinkInstance slabLink)
        {
            XYZ startPoint = point - new XYZ(0, 0, 3.28084);

            ReferenceWithContext ceilingHit = null;
            if (ceilingLink != null)
            {
                ceilingHit = ceilingIntersector.FindNearest(startPoint, XYZ.BasisZ);
                if (ceilingHit != null && ceilingHit.GetReference().ElementId != ceilingLink.Id)
                {
                    ceilingHit = null;
                }
            }

            ReferenceWithContext slabHit = null;
            if (slabLink != null)
            {
                slabHit = slabIntersector.FindNearest(startPoint, XYZ.BasisZ);
                if (slabHit != null && slabHit.GetReference().ElementId != slabLink.Id)
                {
                    slabHit = null;
                }
            }

            if (ceilingHit != null && slabHit != null)
            {
                return (ceilingHit.Proximity < slabHit.Proximity) ? ceilingHit : slabHit;
            }

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

        private void LogSuccess(Dictionary<string, int> summary, string typeName)
        {
            if (!summary.ContainsKey(typeName)) summary[typeName] = 0;
            summary[typeName]++;
        }

        private void LogFailure(Dictionary<string, int> summary, string typeName)
        {
            if (!summary.ContainsKey(typeName)) summary[typeName] = 0;
            summary[typeName]++;
        }

        private void DisplaySummary(Dictionary<string, int> movedSummary, Dictionary<string, int> failedSummary, List<long> outlierIds)
        {
            StringBuilder sb = new StringBuilder();
            if (movedSummary.Any())
            {
                sb.AppendLine("Operation complete. The following elements were moved:");
                foreach (var entry in movedSummary)
                {
                    sb.AppendLine($"  - {entry.Key}: {entry.Value} element(s)");
                }
            }
            else
            {
                sb.AppendLine("No elements were moved.");
            }

            if (failedSummary.Any())
            {
                sb.AppendLine("\nThe following elements could not be re-located (no ceiling or slab found):");
                foreach (var entry in failedSummary)
                {
                    sb.AppendLine($"  - {entry.Key}: {entry.Value} element(s)");
                }
            }

            if (outlierIds.Any())
            {
                sb.AppendLine("\nThe following elements were not moved as the required correction exceeded 3 meters. Please review these outliers:");
                sb.AppendLine("Element IDs: " + string.Join(", ", outlierIds));
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

                var catId = elem.Category.Id.IntegerValue;

                return catId == (long)BuiltInCategory.OST_LightingFixtures ||
                       catId == (long)BuiltInCategory.OST_ElectricalFixtures ||
                       catId == (long)BuiltInCategory.OST_ElectricalEquipment ||
                       catId == (long)BuiltInCategory.OST_CommunicationDevices ||
                       catId == (long)BuiltInCategory.OST_FireAlarmDevices ||
                       catId == (long)BuiltInCategory.OST_SecurityDevices ||
                       catId == (long)BuiltInCategory.OST_MechanicalEquipment ||
                       catId == (long)BuiltInCategory.OST_Sprinklers ||
                       catId == (long)BuiltInCategory.OST_PlumbingFixtures;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
        #endregion
    }
}
