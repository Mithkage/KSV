// --- FILE: ScheduleFloorClass.cs ---
//
// Namespace: RTS.Commands
//
// Class: ScheduleFloorClass
//
// Function: This Revit 2024 external command updates the "Schedule Level" parameter of all
//           lighting fixtures in the host model. It identifies the architectural link model
//           specified in the "Profile Settings" for "Floors". For each fixture, it finds the
//           closest floor directly beneath it in the linked model. If no floor is directly
//           below (e.g., in a void), it searches for the nearest floor horizontally.
//           The script then finds the host level with the same name as the floor's level
//           and updates the fixture's schedule level parameter accordingly, without changing
//           the fixture's physical location.
//
// Author: Your Name/Company
//
// Log:
// - July 30, 2025: Integrated the ProgressBarWindow for user feedback and cancellation.
// - July 30, 2025: Added sorting of elements by location to maximize cache effectiveness.
// - July 30, 2025: Optimized performance by creating the 3D view and ReferenceIntersector only once.
//

#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
    public class ScheduleFloorClass : IExternalCommand
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

            string targetLinkFileName = RTS_RevitUtils.ParseLinkName(settings.FloorsLink);
            var (linkInstance, diagnosticMessage) = RTS_RevitUtils.GetLinkInstance(doc, targetLinkFileName);

            if (linkInstance == null)
            {
                message = "The required source link for 'Floors' could not be found.\n\n" + diagnosticMessage;
                TaskDialog.Show("Link Not Found", message);
                return Result.Failed;
            }

            Document archDoc = linkInstance.GetLinkDocument();
            if (archDoc == null)
            {
                TaskDialog.Show("Error", $"Could not access the document for the link '{targetLinkFileName}'.");
                return Result.Failed;
            }

            // --- 2. Collect Necessary Elements ---
            var lightingFixtures = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_LightingFixtures)
                .WhereElementIsNotElementType()
                .ToList()
                .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                .ToList();

            if (!lightingFixtures.Any())
            {
                TaskDialog.Show("Information", "No lighting fixtures were found in the project to process.");
                return Result.Succeeded;
            }

            var linkedFloors = new FilteredElementCollector(archDoc)
                .OfClass(typeof(Floor))
                .WhereElementIsNotElementType()
                .Cast<Floor>()
                .ToList();

            if (!linkedFloors.Any())
            {
                TaskDialog.Show("Warning", $"No floors were found in the linked model '{targetLinkFileName}'. No fixtures can be updated.");
                return Result.Succeeded;
            }

            var hostLevels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToDictionary(l => l.Name, l => l);

            View3D view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view == null)
            {
                message = "A 3D view is required to perform this operation. Please create one if it doesn't exist.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
            ReferenceIntersector intersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_Floors), FindReferenceTarget.Element, view)
            {
                FindReferencesInRevitLinks = true
            };

            Floor cachedFloor = null;
            XYZ lastFixtureLocation = null;
            const double cacheProximityThreshold = 15.0; // 15 feet

            var updateSummary = new Dictionary<string, int>();

            // --- 3. Initialize and Show Progress Bar ---
            var progressBar = new ProgressBarWindow();
            progressBar.Show();
            int processedCount = 0;

            // --- 4. Start Transaction and Process Fixtures ---
            using (Transaction trans = new Transaction(doc, "Update Lighting Fixture Schedule Levels"))
            {
                try
                {
                    trans.Start();

                    foreach (var fixture in lightingFixtures)
                    {
                        // Check for cancellation request from the progress bar window
                        if (progressBar.IsCancellationPending)
                        {
                            trans.RollBack();
                            progressBar.Close();
                            return Result.Cancelled;
                        }

                        processedCount++;
                        progressBar.UpdateProgress(processedCount, lightingFixtures.Count);

                        LocationPoint locPoint = fixture.Location as LocationPoint;
                        if (locPoint == null) continue;

                        XYZ fixtureLocation = locPoint.Point;
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
                            closestFloor = FindClosestFloor(fixtureLocation, linkedFloors, linkInstance, intersector, archDoc);
                        }

                        lastFixtureLocation = fixtureLocation;
                        cachedFloor = closestFloor;

                        if (closestFloor == null) continue;

                        Level linkedLevel = archDoc.GetElement(closestFloor.LevelId) as Level;
                        if (linkedLevel == null) continue;

                        if (hostLevels.TryGetValue(linkedLevel.Name, out Level hostLevel))
                        {
                            Parameter scheduleLevelParam = fixture.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                            if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
                            {
                                if (scheduleLevelParam.AsElementId() != hostLevel.Id)
                                {
                                    scheduleLevelParam.Set(hostLevel.Id);
                                    string fixtureTypeName = fixture.Name;
                                    if (!updateSummary.ContainsKey(fixtureTypeName))
                                    {
                                        updateSummary[fixtureTypeName] = 0;
                                    }
                                    updateSummary[fixtureTypeName]++;
                                }
                            }
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

            // --- 5. Display Summary Report ---
            if (updateSummary.Any())
            {
                StringBuilder summaryReport = new StringBuilder();
                summaryReport.AppendLine("Successfully updated schedule levels for the following fixture types:");
                summaryReport.AppendLine();
                foreach (var entry in updateSummary)
                {
                    summaryReport.AppendLine($"  - {entry.Key}: {entry.Value} fixture(s) updated.");
                }
                TaskDialog.Show("Update Complete", summaryReport.ToString());
            }
            else
            {
                TaskDialog.Show("No Changes", "No lighting fixtures required a schedule level update.");
            }

            return Result.Succeeded;
        }

        #region Helper Methods

        private Floor FindClosestFloor(XYZ point, List<Floor> allFloors, RevitLinkInstance linkInstance, ReferenceIntersector intersector, Document linkDoc)
        {
            Transform linkTransform = linkInstance.GetTotalTransform();

            ReferenceWithContext refWithContext = intersector.FindNearest(point, XYZ.BasisZ.Negate());

            if (refWithContext != null)
            {
                Reference reference = refWithContext.GetReference();
                ElementId linkedElementId = reference.LinkedElementId;
                if (linkedElementId != ElementId.InvalidElementId)
                {
                    return linkDoc.GetElement(linkedElementId) as Floor;
                }
            }

            double searchRadius = 50.0;
            var outline = new Outline(point - new XYZ(searchRadius, searchRadius, 0), point + new XYZ(searchRadius, searchRadius, 1000));
            var bbFilter = new BoundingBoxIntersectsFilter(outline);

            var candidateFloorIds = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(Floor))
                .WherePasses(bbFilter)
                .ToElementIds();

            var candidateFloors = allFloors.Where(f => candidateFloorIds.Contains(f.Id)).ToList();
            if (!candidateFloors.Any()) return null;

            Floor nearestHorizontalFloor = null;
            double minHorizontalDistance = double.MaxValue;

            foreach (var floor in candidateFloors)
            {
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
                    catch
                    {
                        // Ignore problematic faces
                    }
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
