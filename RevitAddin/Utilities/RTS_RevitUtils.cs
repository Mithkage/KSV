//
// --- FILE: RTS_RevitUtils.cs ---
//
// Description:
// A static utility class containing shared helper methods for the RTS Revit add-in.
// This includes functions for accessing extensible storage, finding and managing
// Revit links, and other common operations to avoid code duplication across commands.
//
// Change Log:
// - July 30, 2025: Added SaveDataToExtensibleStorage method and ReviztoLinkRecord class.
// - July 30, 2025: Initial creation. Extracted shared methods from CastCeilingClass
//                  and ScheduleFloorClass.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
#endregion

namespace RTS.Utilities
{
    #region Data Models for Extensible Storage

    /// <summary>
    /// Data model for storing the profile settings from the Profile Settings window.
    /// </summary>
    public class ProfileSettings
    {
        public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
        public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
        public const string SettingsFieldName = "ProfileSettingsJson";
        public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";

        public string GridsLink { get; set; }
        public string LevelsLink { get; set; }
        public string WallsLink { get; set; }
        public string FloorsLink { get; set; }
        public string CeilingsLink { get; set; }
        public string SlabsLink { get; set; }
    }

    /// <summary>
    /// Data model for storing Revizto link information.
    /// </summary>
    public class ReviztoLinkRecord
    {
        public static readonly Guid SchemaGuid = new Guid("A1B2C3D4-E5F6-47A8-9B0C-1234567890AB");
        public const string SchemaName = "RTS_ReviztoLinkSchema";
        public const string FieldName = "ReviztoLinkJson";
        public const string DataStorageName = "RTS_Revizto_Link_Storage";
        public string LinkName { get; set; }
        public string FilePath { get; set; }
        public string Description { get; set; }
        public string LastModified { get; set; }
    }

    #endregion

    /// <summary>
    /// A static utility class containing shared helper methods for the RTS Revit add-in.
    /// </summary>
    public static class RTS_RevitUtils
    {
        private const string VendorId = "ReTick_Solutions";

        #region Link Management

        /// <summary>
        /// Finds specific RevitLinkInstances based on names provided in the ProfileSettings.
        /// </summary>
        public static (RevitLinkInstance ceilingLink, RevitLinkInstance slabLink, string diagnosticMessage) GetLinkInstances(Document doc, ProfileSettings settings)
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

        /// <summary>
        /// Finds a single RevitLinkInstance by its filename.
        /// </summary>
        public static (RevitLinkInstance linkInstance, string diagnosticMessage) GetLinkInstance(Document doc, string linkFileName)
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

        /// <summary>
        /// Robustly scans the project and returns all valid, loaded Revit Link Instances and their names.
        /// </summary>
        public static (Dictionary<string, RevitLinkInstance> instances, List<string> availableNames, List<string> problematicNames) GetAllProjectLinks(Document doc)
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
                        // Fallback to using the linkType.Name for cloud models.
                    }

                    if (string.IsNullOrEmpty(fileName))
                    {
                        // RevitLinkType.Name is often in the format "FileName.rvt : OtherInfo". We take the part before the colon.
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

        /// <summary>
        /// Parses a link file name from the format "[Discipline] - FileName.rvt".
        /// </summary>
        public static string ParseLinkName(string formattedName)
        {
            if (string.IsNullOrEmpty(formattedName) || formattedName == "<None>") return null;
            var match = Regex.Match(formattedName, @"\] - (.*)");
            return match.Success ? match.Groups[1].Value.Trim() : formattedName;
        }

        #endregion

        #region Extensible Storage

        /// <summary>
        /// Retrieves the saved ProfileSettings from the project's extensible storage.
        /// </summary>
        public static ProfileSettings GetProfileSettings(Document doc)
        {
            var settingsList = RecallDataFromExtensibleStorage<ProfileSettings>(doc, ProfileSettings.SettingsSchemaGuid, ProfileSettings.SettingsSchemaName, ProfileSettings.SettingsFieldName, ProfileSettings.SettingsDataStorageElementName);
            return settingsList.FirstOrDefault();
        }

        /// <summary>
        /// A generic method to save a list of objects to a JSON string in Extensible Storage.
        /// </summary>
        public static void SaveDataToExtensibleStorage<T>(Document doc, List<T> dataList, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName)
        {
            Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName);
            DataStorage dataStorage = GetOrCreateDataStorage(doc, dataStorageElementName);
            string jsonString = JsonSerializer.Serialize(dataList, new JsonSerializerOptions { WriteIndented = true });
            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(fieldName), jsonString);
            dataStorage.SetEntity(entity);
        }

        /// <summary>
        /// A generic method to retrieve a list of objects from a JSON string stored in Extensible Storage.
        /// </summary>
        public static List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) where T : new()
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

            try
            {
                return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>();
            }
            catch
            {
                return new List<T>();
            }
        }

        private static Schema GetOrCreateSchema(Guid schemaGuid, string schemaName, string fieldName)
        {
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid);
                schemaBuilder.SetSchemaName(schemaName);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor);
                schemaBuilder.SetVendorId(VendorId);
                schemaBuilder.AddSimpleField(fieldName, typeof(string));
                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        private static DataStorage GetOrCreateDataStorage(Document doc, string dataStorageElementName)
        {
            var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            if (dataStorage == null)
            {
                dataStorage = DataStorage.Create(doc);
                dataStorage.Name = dataStorageElementName;
            }
            return dataStorage;
        }

        #endregion
    }
}
