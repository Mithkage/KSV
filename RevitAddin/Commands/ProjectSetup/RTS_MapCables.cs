//
// File: RTS_MapCables.cs
//
// Namespace: RTS.Commands.ProjectSetup
//
// Class: RTS_MapCablesClass
//
// Function: This file defines a Revit external command to generate and apply parameter
//           mapping files. It now integrates with the centralized ScheduleManager to provide
//           granular control over mapping for each standard schedule type.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Integrated the new MapCablesWindow UI, removing the old placeholder classes.
// - [APPLIED FIX]: Refactored to use the centralized ScheduleManager for all schedule definitions.
// - [APPLIED FIX]: Implemented a new UI to allow selection of specific schedules for template generation or mapping.
//
#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using RtsScheduleDef = RTS.Utilities.ScheduleDefinition;
#endregion

namespace RTS.Commands.ProjectSetup
{
    public class CsvMappingEntry
    {
        public string RTSGuid { get; set; }
        public string RTSName { get; set; }
        public string DataType { get; set; }
        public string ClientGuid { get; set; }
        public string ClientParameterName { get; set; }
        public string CsvNotes { get; set; }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RTS_MapCablesClass : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // --- 1. Populate UI with schedules from the central manager ---
            var scheduleItems = ScheduleManager.StandardSchedules
                .Select(s => new ScheduleMappingItem { Name = s.Name, IsSelected = false })
                .ToList();

            var selectionWindow = new MapCablesWindow(scheduleItems);
            bool? dialogResult = selectionWindow.ShowDialog();

            if (dialogResult != true)
            {
                return Result.Cancelled;
            }

            var selectedSchedules = selectionWindow.Schedules.Where(s => s.IsSelected).ToList();
            if (!selectedSchedules.Any())
            {
                TaskDialog.Show("Information", "No schedules were selected.");
                return Result.Succeeded;
            }

            if (selectionWindow.GenerateTemplates)
            {
                return HandleTemplateGeneration(selectedSchedules);
            }
            else if (selectionWindow.MapFromFiles)
            {
                return HandleMapping(commandData.Application.ActiveUIDocument, selectedSchedules);
            }

            return Result.Succeeded;
        }

        private Result HandleTemplateGeneration(List<ScheduleMappingItem> selectedSchedules)
        {
            string folderPath;
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder to save the mapping templates";
                if (folderDialog.ShowDialog() != DialogResult.OK) return Result.Cancelled;
                folderPath = folderDialog.SelectedPath;
            }

            int generatedCount = 0;
            foreach (var item in selectedSchedules)
            {
                var scheduleDef = ScheduleManager.StandardSchedules.FirstOrDefault(s => s.Name == item.Name);
                if (scheduleDef != null)
                {
                    GenerateSingleTemplate(scheduleDef, folderPath);
                    generatedCount++;
                }
            }

            TaskDialog.Show("Success", $"{generatedCount} mapping templates were generated successfully in:\n{folderPath}");
            return Result.Succeeded;
        }

        private void GenerateSingleTemplate(RtsScheduleDef scheduleDef, string folderPath)
        {
            var csvContent = new StringBuilder();
            csvContent.AppendLine("RTS GUID,RTS NAME,DATATYPE,CLIENT GUID,CLIENT PARAMETER NAME,NOTES");

            foreach (var guid in scheduleDef.RequiredSharedParameterGuids)
            {
                var paramDef = SharedParameterData.MySharedParameters.FirstOrDefault(p => p.Guid == guid);
                if (paramDef != null)
                {
                    csvContent.AppendLine($"{paramDef.Guid},{paramDef.Name},{paramDef.DataType},,,");
                }
            }

            string filePath = Path.Combine(folderPath, $"{scheduleDef.Name}_Mapping.csv");
            File.WriteAllText(filePath, csvContent.ToString());
        }

        private Result HandleMapping(UIDocument uidoc, List<ScheduleMappingItem> selectedSchedules)
        {
            string[] csvFilePaths;
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv";
                openFileDialog.Title = "Select the CSV mapping files to import";
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() != DialogResult.OK) return Result.Cancelled;
                csvFilePaths = openFileDialog.FileNames;
            }

            var report = new StringBuilder("Mapping Process Summary:\n\n");
            int totalParamsUpdated = 0;

            foreach (string filePath in csvFilePaths)
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                var scheduleDef = ScheduleManager.StandardSchedules.FirstOrDefault(s => fileName.StartsWith(s.Name));

                if (scheduleDef == null)
                {
                    report.AppendLine($"- Skipped '{fileName}': Could not match to a standard schedule.");
                    continue;
                }

                List<CsvMappingEntry> mappingEntries = ParseMappingFile(filePath);
                if (!mappingEntries.Any())
                {
                    report.AppendLine($"- Skipped '{fileName}': No valid mapping entries found.");
                    continue;
                }

                int updatedCount = MapParametersOnElements(uidoc.Document, mappingEntries, scheduleDef.Category);
                totalParamsUpdated += updatedCount;
                report.AppendLine($"- Processed '{fileName}': Updated {updatedCount} parameters for the '{scheduleDef.Category}' category.");
            }

            TaskDialog.Show("Mapping Complete", report.ToString());
            return Result.Succeeded;
        }

        private List<CsvMappingEntry> ParseMappingFile(string filePath)
        {
            var entries = new List<CsvMappingEntry>();
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] parts = lines[i].Split(',');
                    if (parts.Length >= 6)
                    {
                        if (!string.IsNullOrEmpty(parts[3].Trim()) || !string.IsNullOrEmpty(parts[4].Trim()))
                        {
                            entries.Add(new CsvMappingEntry
                            {
                                RTSGuid = parts[0].Trim(),
                                RTSName = parts[1].Trim(),
                                DataType = parts[2].Trim(),
                                ClientGuid = parts[3].Trim(),
                                ClientParameterName = parts[4].Trim(),
                                CsvNotes = parts[5].Trim()
                            });
                        }
                    }
                }
            }
            catch { /* Suppress errors, return empty list on failure */ }
            return entries;
        }

        private int MapParametersOnElements(Document doc, List<CsvMappingEntry> mappingEntries, BuiltInCategory category)
        {
            var elementsToProcess = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToList();

            if (!elementsToProcess.Any()) return 0;

            int updatedCount = 0;
            using (var t = new Transaction(doc, $"Map Parameters for {category}"))
            {
                t.Start();
                foreach (Element element in elementsToProcess)
                {
                    foreach (CsvMappingEntry entry in mappingEntries)
                    {
                        Parameter sourceParam = GetParameterByNameOrGuid(element, entry.ClientGuid, entry.ClientParameterName);
                        if (sourceParam == null || !sourceParam.HasValue) continue;

                        Parameter targetParam = GetParameterByNameOrGuid(element, entry.RTSGuid, entry.RTSName);
                        if (targetParam != null && !targetParam.IsReadOnly)
                        {
                            if (GetParameterValueAsString(targetParam) != GetParameterValueAsString(sourceParam))
                            {
                                SetParameterValue(targetParam, sourceParam);
                                updatedCount++;
                            }
                        }
                    }
                }
                t.Commit();
            }
            return updatedCount;
        }

        private Parameter GetParameterByNameOrGuid(Element element, string guidString, string nameString)
        {
            if (!string.IsNullOrEmpty(guidString) && Guid.TryParse(guidString, out Guid guid))
            {
                Parameter param = element.get_Parameter(guid);
                if (param != null) return param;
            }
            if (!string.IsNullOrEmpty(nameString))
            {
                return element.LookupParameter(nameString);
            }
            return null;
        }

        private string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue) return string.Empty;
            return param.AsValueString() ?? param.AsString() ?? "";
        }

        private void SetParameterValue(Parameter target, Parameter source)
        {
            switch (source.StorageType)
            {
                case StorageType.String: target.Set(source.AsString()); break;
                case StorageType.Double: target.Set(source.AsDouble()); break;
                case StorageType.Integer: target.Set(source.AsInteger()); break;
                case StorageType.ElementId: target.Set(source.AsElementId()); break;
            }
        }
    }
}
