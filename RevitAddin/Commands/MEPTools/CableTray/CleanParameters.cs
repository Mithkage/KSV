// File: CleanParameters.cs
//
// Namespace: RTS.Commands.MEPTools.CableTray
//
// Class: CleanParametersClass
//
// Function: This Revit external command iterates through all Cable Trays, Conduits,
//           and their respective fittings in the active model. It applies a standardized
//           cleaning logic to the values of all "RTS_Cable_XX" shared parameters on these
//           elements to ensure consistent naming conventions.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Implemented robust per-element error handling to prevent a single failure from stopping the entire process.
// - [APPLIED FIX]: Added detailed CSV reporting to log every parameter change (original vs. new value) and any skipped elements.
//
// Author: Kyle Vorster
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Utilities; // Required to access SharedParameters
#endregion

namespace RTS.Commands.MEPTools.CableTray
{
    /// <summary>
    /// A data class for logging the cleaning process.
    /// </summary>
    public class CleaningReportLog
    {
        public string ElementId { get; set; }
        public string ElementType { get; set; }
        public string ParameterName { get; set; }
        public string OriginalValue { get; set; }
        public string NewValue { get; set; }
        public string Status { get; set; }
    }

    /// <summary>
    /// A Revit external command to clean and standardize cable parameter values.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CleanParametersClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Guid> cableGuids = SharedParameters.Cable.AllCableGuids;

            var categoriesToClean = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };
            var categoryFilter = new ElementMulticategoryFilter(categoriesToClean);

            var elementsToProcess = new FilteredElementCollector(doc)
                .WherePasses(categoryFilter)
                .WhereElementIsNotElementType()
                .ToList();

            if (!elementsToProcess.Any())
            {
                TaskDialog.Show("Clean Parameters", "No relevant cable trays, conduits, or fittings were found in the project.");
                return Result.Succeeded;
            }

            var reportLogs = new List<CleaningReportLog>();
            int elementsUpdated = 0;
            int parametersCleaned = 0;

            using (var tx = new Transaction(doc, "Clean Cable Parameters"))
            {
                tx.Start();

                foreach (var elem in elementsToProcess)
                {
                    bool elementWasUpdated = false;
                    try
                    {
                        foreach (var guid in cableGuids)
                        {
                            Parameter cableParam = elem.get_Parameter(guid);

                            if (cableParam == null || !cableParam.HasValue || string.IsNullOrWhiteSpace(cableParam.AsString()))
                            {
                                continue;
                            }

                            if (cableParam.IsReadOnly)
                            {
                                reportLogs.Add(new CleaningReportLog
                                {
                                    ElementId = elem.Id.ToString(),
                                    ElementType = elem.Category.Name,
                                    ParameterName = cableParam.Definition.Name,
                                    OriginalValue = cableParam.AsString(),
                                    NewValue = "N/A",
                                    Status = "Skipped (Read-Only)"
                                });
                                continue;
                            }

                            string originalValue = cableParam.AsString();
                            string cleanedValue = CleanCableReference(originalValue);

                            if (originalValue != cleanedValue)
                            {
                                cableParam.Set(cleanedValue);
                                parametersCleaned++;
                                elementWasUpdated = true;

                                reportLogs.Add(new CleaningReportLog
                                {
                                    ElementId = elem.Id.ToString(),
                                    ElementType = elem.Category.Name,
                                    ParameterName = cableParam.Definition.Name,
                                    OriginalValue = originalValue,
                                    NewValue = cleanedValue,
                                    Status = "Cleaned"
                                });
                            }
                        }

                        if (elementWasUpdated)
                        {
                            elementsUpdated++;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error for this specific element and continue
                        reportLogs.Add(new CleaningReportLog
                        {
                            ElementId = elem.Id.ToString(),
                            ElementType = elem.Category.Name,
                            ParameterName = "N/A",
                            OriginalValue = "N/A",
                            NewValue = "N/A",
                            Status = $"Error: {ex.Message}"
                        });
                    }
                }

                tx.Commit();
            }

            // Generate and save the CSV report
            string reportPath = GenerateReportCsv(reportLogs);

            TaskDialog.Show("Process Complete",
                $"Cleaning process finished.\n\n" +
                $"{parametersCleaned} parameter values were cleaned across {elementsUpdated} elements.\n\n" +
                $"A detailed report has been saved to:\n{reportPath}");

            return Result.Succeeded;
        }

        private string GenerateReportCsv(List<CleaningReportLog> logs)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileName = $"CleanParameters_Report_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string fullPath = Path.Combine(folderPath, fileName);

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Element ID,Element Type,Parameter Name,Original Value,New Value,Status");

                foreach (var log in logs)
                {
                    sb.AppendLine($"\"{log.ElementId}\",\"{log.ElementType}\",\"{log.ParameterName}\",\"{log.OriginalValue}\",\"{log.NewValue}\",\"{log.Status}\"");
                }

                File.WriteAllText(fullPath, sb.ToString());
            }
            catch (Exception)
            {
                return "Could not save report file.";
            }

            return fullPath;
        }

        private string CleanCableReference(string cableReference)
        {
            if (string.IsNullOrEmpty(cableReference))
            {
                return cableReference;
            }

            string cleaned = cableReference.Trim();

            int openParenIndex = cleaned.IndexOf('(');
            if (openParenIndex != -1)
            {
                cleaned = cleaned.Substring(0, openParenIndex).Trim();
            }

            int firstSlashIndex = cleaned.IndexOf('/');
            if (firstSlashIndex != -1)
            {
                cleaned = cleaned.Substring(0, firstSlashIndex).Trim();
            }

            string[] parts = cleaned.Split('-');
            Regex prefixPattern = new Regex(@"^[A-Za-z]{2}\d{2}", RegexOptions.IgnoreCase);

            if (parts.Length >= 3 && prefixPattern.IsMatch(parts[0]))
            {
                if (!int.TryParse(parts[2], out _))
                {
                    return $"{parts[0]}-{parts[1]}";
                }
            }

            if (parts.Length >= 4 && prefixPattern.IsMatch(parts[0]))
            {
                if (int.TryParse(parts[2], out _) && !int.TryParse(parts[3], out _))
                {
                    return $"{parts[0]}-{parts[1]}-{parts[2]}";
                }
            }

            return cleaned;
        }
    }
}
