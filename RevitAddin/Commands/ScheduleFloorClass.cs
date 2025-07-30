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
// - July 30, 2025: Limited element processing to 20 for testing purposes.
// - July 30, 2025: Added sorting of elements by location to maximize cache effectiveness.
// - July 30, 2025: Optimized performance by creating the 3D view and ReferenceIntersector only once.
// - July 30, 2025: Implemented a more robust, multi-step link finding method with enhanced diagnostics.
//

#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
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

            // --- 1. Retrieve Profile Settings to find the Architectural Link ---
            var settings = GetProfileSettings(doc);
            if (settings == null)
            {
                message = "Profile Settings could not be loaded. Please configure them first.";
                TaskDialog.Show("Error", message);
                return Result.Cancelled;
            }

            string targetLinkFileName = ParseLinkName(settings.FloorsLink);
            var (linkInstance, diagnosticMessage) = GetLinkInstance(doc, targetLinkFileName);

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
            // --- FOR TESTING ONLY: Limit to the first 20 elements. Remove .Take(20) for production. ---
            var lightingFixtures = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_LightingFixtures)
                .WhereElementIsNotElementType()
                .ToList()
                .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                .Take(20)
                .ToList();

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

            // --- OPTIMIZATION: Get 3D view and create ReferenceIntersector ONCE ---
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

            // --- OPTIMIZATION: Caching variables ---
            Floor cachedFloor = null;
            XYZ lastFixtureLocation = null;
            const double cacheProximityThreshold = 15.0; // 15 feet

            var updateSummary = new Dictionary<string, int>();

            // --- 3. Start Transaction and Process Fixtures ---
            using (Transaction trans = new Transaction(doc, "Update Lighting Fixture Schedule Levels"))
            {
                try
                {
                    trans.Start();

                    foreach (var fixture in lightingFixtures)
                    {
                        LocationPoint locPoint = fixture.Location as LocationPoint;
                        if (locPoint == null) continue;

                        XYZ fixtureLocation = locPoint.Point;
                        Floor closestFloor = null;

                        // --- OPTIMIZATION: Caching Strategy ---
                        if (cachedFloor != null && lastFixtureLocation != null && fixtureLocation.DistanceTo(lastFixtureLocation) < cacheProximityThreshold)
                        {
                            // If this fixture is close to the last one, test the cached floor first.
                            if (IsPointAboveFloor(fixtureLocation, cachedFloor, linkInstance))
                            {
                                closestFloor = cachedFloor;
                            }
                        }

                        // If cache didn't work, perform the full search
                        if (closestFloor == null)
                        {
                            closestFloor = FindClosestFloor(fixtureLocation, linkedFloors, linkInstance, intersector, archDoc);
                        }

                        // Update cache for next iteration
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
            }

            // --- 4. Display Summary Report ---
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

        private ProfileSettings GetProfileSettings(Document doc)
        {
            var settingsList = RecallDataFromExtensibleStorage<ProfileSettings>(doc, ProfileSettings.SettingsSchemaGuid, ProfileSettings.SettingsSchemaName, ProfileSettings.SettingsFieldName, ProfileSettings.SettingsDataStorageElementName);
            return settingsList.FirstOrDefault();
        }

        private (RevitLinkInstance linkInstance, string diagnosticMessage) GetLinkInstance(Document doc, string linkFileName)
        {
            var (allLinks, availableNames, problematicNames) = GetAllProjectLinks(doc);

            RevitLinkInstance linkInstance = null;
            if (!string.IsNullOrEmpty(linkFileName) && allLinks.ContainsKey(linkFileName))
            {
                linkInstance = allLinks[linkFileName];
            }

            string diagnostics = "";
            if (linkInstance == null && !string.IsNullOrEmpty(linkFileName))
            {
                diagnostics += $"Link Not Found in Settings: '{linkFileName}'\n";
                diagnostics += "\nAvailable links found in project:\n" + (availableNames.Any() ? " - " + string.Join("\n - ", availableNames) : "None");
                if (problematicNames.Any())
                {
                    diagnostics += "\n\nCould not read the following links (may be unloaded or corrupt):\n" + " - " + string.Join("\n - ", problematicNames);
                }
            }

            return (linkInstance, diagnostics);
        }

        private (Dictionary<string, RevitLinkInstance> instances, List<string> availableNames, List<string> problematicNames) GetAllProjectLinks(Document doc)
        {
            var instances = new Dictionary<string, RevitLinkInstance>(StringComparer.OrdinalIgnoreCase);
            var availableNames = new List<string>();
            var problematicNames = new List<string>();

            var allLinkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var linkInstance in allLinkInstances)
            {
                try
                {
                    var linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                    if (linkType == null)
                    {
                        problematicNames.Add($"Instance ID {linkInstance.Id} has no valid link type.");
                        continue;
                    }

                    string fileName = null;
                    try
                    {
                        var externalRef = linkType.GetExternalFileReference();
                        if (externalRef != null && externalRef.GetPath() != null)
                        {
                            string visiblePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(externalRef.GetPath());
                            fileName = Path.GetFileName(visiblePath);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // Fallback to using the linkType.Name.
                    }

                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = linkType.Name.Split(':').FirstOrDefault()?.Trim();
                    }

                    if (!string.IsNullOrEmpty(fileName) && !instances.ContainsKey(fileName))
                    {
                        instances.Add(fileName, linkInstance);
                        availableNames.Add(fileName);
                    }
                    else if (string.IsNullOrEmpty(fileName))
                    {
                        problematicNames.Add($"Instance ID {linkInstance.Id} (Type: {linkType.Name}) could not be resolved to a filename.");
                    }
                }
                catch (Exception ex)
                {
                    problematicNames.Add($"Instance ID {linkInstance.Id} failed with error: {ex.Message}");
                }
            }
            return (instances, availableNames, problematicNames);
        }

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

            // --- OPTIMIZATION: Bounding Box Filtering ---
            double searchRadius = 50.0; // 50 feet search radius
            var outline = new Outline(point - new XYZ(searchRadius, searchRadius, 0), point + new XYZ(searchRadius, searchRadius, 1000));
            var bbFilter = new BoundingBoxIntersectsFilter(outline);

            // Get the IDs of floors in the linked model that intersect the search box
            var candidateFloorIds = new FilteredElementCollector(linkDoc)
                .OfClass(typeof(Floor))
                .WherePasses(bbFilter)
                .ToElementIds();

            // Create a filtered list of floor elements to check
            var candidateFloors = allFloors.Where(f => candidateFloorIds.Contains(f.Id)).ToList();
            if (!candidateFloors.Any()) return null; // No floors found in the search radius

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

        private bool IsPointAboveFloor(XYZ point, Floor floor, RevitLinkInstance linkInstance)
        {
            if (floor == null) return false;

            Transform linkTransform = linkInstance.GetTotalTransform();
            GeometryElement geoElem = floor.get_Geometry(new Options());

            foreach (GeometryObject geoObj in geoElem)
            {
                Solid solid = geoObj as Solid;
                if (solid == null || solid.Faces.Size == 0) continue;

                Solid transformedSolid = SolidUtils.CreateTransformed(solid, linkTransform);
                try
                {
                    // Check if the point is inside the solid's horizontal boundaries
                    // This is a simplified check but effective for this caching purpose
                    BoundingBoxXYZ bb = transformedSolid.GetBoundingBox();
                    if (point.X >= bb.Min.X && point.X <= bb.Max.X &&
                        point.Y >= bb.Min.Y && point.Y <= bb.Max.Y)
                    {
                        return true;
                    }
                }
                catch { /* Ignore problematic solids */ }
            }
            return false;
        }

        private string ParseLinkName(string formattedName)
        {
            if (string.IsNullOrEmpty(formattedName)) return null;
            var match = Regex.Match(formattedName, @"\] - (.*)");
            return match.Success ? match.Groups[1].Value.Trim() : formattedName;
        }

        #endregion

        #region Extensible Storage Classes and Helpers
        public class ProfileSettings
        {
            public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
            public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
            public const string SettingsFieldName = "ProfileSettingsJson";
            public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";
            public string FloorsLink { get; set; }
        }

        public List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) where T : new()
        {
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema == null) return new List<T>();
            var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            if (dataStorage == null) return new List<T>();
            Entity entity = dataStorage.GetEntity(schema);
            if (!entity.IsValid()) return new List<T>();
            string jsonString = entity.Get<string>(schema.GetField(fieldName));
            if (string.IsNullOrEmpty(jsonString)) return new List<T>();
            try { return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>(); }
            catch { return new List<T>(); }
        }
        #endregion
    }
}
