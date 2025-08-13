// --- FILE: RTS_RevitUtils.cs (UPDATED) ---
//
// File: RTS_RevitUtils.cs
// Namespace: RTS.Utilities
//
// This file contains utility methods for the Revit Add-in.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Linq;
using System.Text.Json;
using RTS.UI; // Use the UI namespace for the correct constants
#endregion

namespace RTS.Utilities
{
    public static class RTS_RevitUtils
    {
        /// <summary>
        /// Retrieves and deserializes profile settings from Extensible Storage.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <returns>A deserialized ProfileSettings object, or a new instance if not found.</returns>
        public static ProfileSettings GetProfileSettings(Document doc)
        {
            // Use constants from the UI class to ensure consistency
            Schema schema = Schema.Lookup(ProfileSettingsWindow.SettingsSchemaGuid);
            if (schema == null)
            {
                return new ProfileSettings();
            }

            var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            var dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == ProfileSettingsWindow.SettingsDataStorageElementName);
            if (dataStorage == null)
            {
                return new ProfileSettings();
            }

            var entity = dataStorage.GetEntity(schema);
            if (!entity.IsValid())
            {
                return new ProfileSettings();
            }

            string json = entity.Get<string>(schema.GetField(ProfileSettingsWindow.SettingsFieldName));
            if (string.IsNullOrEmpty(json))
            {
                return new ProfileSettings();
            }

            try
            {
                return JsonSerializer.Deserialize<ProfileSettings>(json) ?? new ProfileSettings();
            }
            catch
            {
                return new ProfileSettings();
            }
        }

        /// <summary>
        /// Parses the link name from the format "[Discipline] - LinkName".
        /// </summary>
        public static string ParseLinkName(string selectedLink)
        {
            if (string.IsNullOrEmpty(selectedLink) || !selectedLink.Contains("] - "))
            {
                return selectedLink;
            }
            return selectedLink.Split(new[] { "] - " }, StringSplitOptions.None).Last();
        }

        /// <summary>
        /// Gets a RevitLinkInstance by its name.
        /// </summary>
        public static (RevitLinkInstance, string) GetLinkInstance(Document doc, string linkName)
        {
            if (string.IsNullOrEmpty(linkName) || linkName == "<None>")
            {
                return (null, "No link specified.");
            }

            var linkInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .FirstOrDefault(li => li.Name.Contains(linkName));

            if (linkInstance == null)
            {
                return (null, $"Link instance '{linkName}' not found.");
            }

            return (linkInstance, "Success");
        }
    }

    // This legacy class is renamed to avoid ambiguity.
    public class LegacyProfileSettings { }
}