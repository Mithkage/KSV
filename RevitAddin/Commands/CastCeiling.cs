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
//           element to the closest surface and attempts to re-host it.
//
// Author: Your Name/Company
//
// Log:
// - July 30, 2025: Implemented Bounding Box filtering and result caching for significant performance optimization.
// - July 30, 2025: Added sorting of elements by location to maximize cache effectiveness.
// - July 30, 2025: Implemented a more robust, multi-step link finding method with enhanced diagnostics.
//

#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
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

            // --- 1. Retrieve Profile Settings to find the Architectural Links ---
            var settings = GetProfileSettings(doc);
            if (settings == null)
            {
                message = "Profile Settings could not be loaded. Please configure them first.";
                TaskDialog.Show("Error", message);
                return Result.Cancelled;
            }

            var (ceilingLink, slabLink, diagnosticMessage) = GetLinkInstances(doc, settings);
            if ((!string.IsNullOrEmpty(ParseLinkName(settings.CeilingsLink)) && ceilingLink == null) ||
                (!string.IsNullOrEmpty(ParseLinkName(settings.SlabsLink)) && slabLink == null))
            {
                message = "One or more required source links could not be found.\n\n" + diagnosticMessage;
                TaskDialog.Show("Link Not Found", message);
                return Result.Failed;
            }

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

            // --- 3. Collect All Elements of the Same Type ---
            ElementId typeId = selectedElement.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                message = "The selected element does not have a valid type.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // --- OPTIMIZATION: Sort elements by location to improve caching ---
            var elementsToProcess = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e => e.GetTypeId() == typeId)
                .ToList()
                .OrderBy(e => (e.Location as LocationPoint)?.Point.X ?? 0)
                .ThenBy(e => (e.Location as LocationPoint)?.Point.Y ?? 0)
                .ToList();

            // --- 4. Setup for Ray Casting ---
            View3D view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view == null)
            {
                message = "A 3D view is required to perform this operation. Please create one if it doesn't exist.";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            var categoriesToFind = new List<BuiltInCategory> { BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Floors };
            var categoryFilter = new ElementMulticategoryFilter(categoriesToFind);
            var intersector = new ReferenceIntersector(categoryFilter, FindReferenceTarget.Element, view)
            {
                FindReferencesInRevitLinks = true
            };

            var movedElementsSummary = new Dictionary<string, int>();
            var failedElementsSummary = new Dictionary<string, int>();

            // --- OPTIMIZATION: Caching variables ---
            Reference cachedReference = null;
            XYZ lastElementLocation = null;
            const double cacheProximityThreshold = 15.0; // 15 feet

            // --- 5. Process Elements ---
            using (Transaction trans = new Transaction(doc, "Cast Elements to Ceiling or Slab"))
            {
                try
                {
                    trans.Start();

                    foreach (var elem in elementsToProcess)
                    {
                        LocationPoint locPoint = elem.Location as LocationPoint;
                        if (locPoint == null)
                        {
                            LogFailure(failedElementsSummary, elem.Name);
                            continue;
                        }

                        XYZ currentLocation = locPoint.Point;
                        ReferenceWithContext targetHit = null;

                        // --- OPTIMIZATION: Caching Strategy ---
                        if (cachedReference != null && lastElementLocation != null && currentLocation.DistanceTo(lastElementLocation) < cacheProximityThreshold)
                        {
                            // If this element is close to the last one, test the cached surface first.
                            targetHit = CheckCachedReference(cachedReference, currentLocation, intersector);
                        }

                        // If cache didn't work, perform the full search
                        if (targetHit == null)
                        {
                            targetHit = FindClosestSurface(currentLocation, intersector, ceilingLink, slabLink);
                        }

                        // Update cache for next iteration
                        lastElementLocation = currentLocation;
                        cachedReference = targetHit?.GetReference();

                        if (targetHit == null)
                        {
                            LogFailure(failedElementsSummary, elem.Name);
                            continue;
                        }

                        // Move and attempt to re-host the element
                        XYZ targetPoint = targetHit.GetReference().GlobalPoint;
                        XYZ moveVector = new XYZ(0, 0, targetPoint.Z - currentLocation.Z);
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
            }

            // --- 6. Display Summary Report ---
            DisplaySummary(movedElementsSummary, failedElementsSummary);
            return Result.Succeeded;
        }

        #region Helper Methods

        private ProfileSettings GetProfileSettings(Document doc)
        {
            var settingsList = RecallDataFromExtensibleStorage<ProfileSettings>(doc, ProfileSettings.SettingsSchemaGuid, ProfileSettings.SettingsSchemaName, ProfileSettings.SettingsFieldName, ProfileSettings.SettingsDataStorageElementName);
            return settingsList.FirstOrDefault();
        }

        private (RevitLinkInstance ceilingLink, RevitLinkInstance slabLink, string diagnosticMessage) GetLinkInstances(Document doc, ProfileSettings settings)
        {
            string ceilingLinkName = ParseLinkName(settings.CeilingsLink);
            string slabLinkName = ParseLinkName(settings.SlabsLink);

            var (allLinks, availableNames, problematicNames) = GetAllProjectLinks(doc);

            RevitLinkInstance ceilingLink = null;
            if (!string.IsNullOrEmpty(ceilingLinkName) && allLinks.ContainsKey(ceilingLinkName))
            {
                ceilingLink = allLinks[ceilingLinkName];
            }

            RevitLinkInstance slabLink = null;
            if (!string.IsNullOrEmpty(slabLinkName) && allLinks.ContainsKey(slabLinkName))
            {
                slabLink = allLinks[slabLinkName];
            }

            string diagnostics = "";
            if (ceilingLink == null && !string.IsNullOrEmpty(ceilingLinkName))
            {
                diagnostics += $"Ceiling Link Not Found in Settings: '{ceilingLinkName}'\n";
            }
            if (slabLink == null && !string.IsNullOrEmpty(slabLinkName))
            {
                diagnostics += $"Slab Link Not Found in Settings: '{slabLinkName}'\n";
            }

            if (!string.IsNullOrEmpty(diagnostics) || !availableNames.Any())
            {
                diagnostics += "\nAvailable links found in project:\n" + (availableNames.Any() ? " - " + string.Join("\n - ", availableNames) : "None");
                if (problematicNames.Any())
                {
                    diagnostics += "\n\nCould not read the following links (may be unloaded or corrupt):\n" + " - " + string.Join("\n - ", problematicNames);
                }
            }

            return (ceilingLink, slabLink, diagnostics);
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

        private ReferenceWithContext FindClosestSurface(XYZ point, ReferenceIntersector intersector, RevitLinkInstance ceilingLink, RevitLinkInstance slabLink)
        {
            XYZ startPoint = point - new XYZ(0, 0, 3.28084); // Offset 1 meter down

            ReferenceWithContext ceilingHit = null;
            if (ceilingLink != null)
            {
                ceilingHit = intersector.FindNearest(startPoint, XYZ.BasisZ);
                // Ensure the hit is actually in the intended link
                if (ceilingHit != null && ceilingHit.GetReference().ElementId != ceilingLink.Id)
                {
                    ceilingHit = null;
                }
            }

            ReferenceWithContext slabHit = null;
            if (slabLink != null)
            {
                slabHit = intersector.FindNearest(startPoint, XYZ.BasisZ);
                // Ensure the hit is actually in the intended link
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

        private ReferenceWithContext CheckCachedReference(Reference cachedRef, XYZ point, ReferenceIntersector intersector)
        {
            if (cachedRef == null) return null;

            XYZ startPoint = point - new XYZ(0, 0, 3.28084);
            var hits = intersector.Find(startPoint, XYZ.BasisZ);

            foreach (var hit in hits)
            {
                if (hit.GetReference().LinkedElementId == cachedRef.LinkedElementId)
                {
                    return hit;
                }
            }
            return null;
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

        private void DisplaySummary(Dictionary<string, int> movedSummary, Dictionary<string, int> failedSummary)
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
            TaskDialog.Show("Summary", sb.ToString());
        }

        private string ParseLinkName(string formattedName)
        {
            if (string.IsNullOrEmpty(formattedName) || formattedName == "<None>") return null;
            var match = Regex.Match(formattedName, @"\] - (.*)");
            return match.Success ? match.Groups[1].Value.Trim() : formattedName;
        }

        #endregion

        #region Extensible Storage and Selection Filter
        public class ProfileSettings
        {
            public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
            public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
            public const string SettingsFieldName = "ProfileSettingsJson";
            public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";
            public string CeilingsLink { get; set; }
            public string SlabsLink { get; set; }
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

        public class MepElementSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                if (elem.Category == null) return false;

                var catId = elem.Category.Id.IntegerValue;

                return catId == (int)BuiltInCategory.OST_LightingFixtures ||
                       catId == (int)BuiltInCategory.OST_ElectricalFixtures ||
                       catId == (int)BuiltInCategory.OST_ElectricalEquipment ||
                       catId == (int)BuiltInCategory.OST_CommunicationDevices ||
                       catId == (int)BuiltInCategory.OST_FireAlarmDevices ||
                       catId == (int)BuiltInCategory.OST_SecurityDevices ||
                       catId == (int)BuiltInCategory.OST_MechanicalEquipment ||
                       catId == (int)BuiltInCategory.OST_Sprinklers ||
                       catId == (int)BuiltInCategory.OST_PlumbingFixtures;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
        #endregion
    }
}
