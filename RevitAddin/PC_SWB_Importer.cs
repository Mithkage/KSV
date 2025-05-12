// src/PC_SWBImporter.cs

using System.Text; // Required for Encoding support
using System.Text.RegularExpressions; // Required for Regex number extraction (still used if other parts need it, but not for frame size)
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms; // Needed for OpenFileDialog
using Microsoft.VisualBasic.FileIO; // Needed for TextFieldParser
using System.Globalization; // Needed for CultureInfo

namespace PC_SWB_Importer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_SWB_ImporterClass : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            var app = uiapp.Application; 
            Document doc = uidoc.Document;

            // --- 1. Get CSV File Path from User ---
            string csvFilePath = GetCsvFilePath();
            if (string.IsNullOrEmpty(csvFilePath))
            {
                message = "Operation cancelled by user.";
                return Result.Cancelled;
            }

            // --- 2. Define Parameter Mapping ---
            Dictionary<string, string> parameterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"Cable Reference", "PC_Cable Reference"},
                {"SWB From", "PC_SWB From"},
                {"SWB To", "PC_SWB To"},
                {"SWB Type", "PC_SWB Type"},
                {"SWB Load", "PC_SWB Load"},
                {"SWB Load Scope", "PC_SWB Load Scope"},
                {"SWB PF", "PC_SWB PF"},
                {"Cable Length", "PC_Cable Length"},
                {"Cable Size - Active conductors", "PC_Cable Size - Active conductors"},
                {"Cable Size - Neutral conductors", "PC_Cable Size - Neutral conductors"},
                {"Cable Size - Earthing conductor", "PC_Cable Size - Earthing conductor"},
                {"Active Conductor material", "PC_Active Conductor material"},
                {"# of Phases", "PC_# of Phases"},
                {"Cable Insulation", "PC_Cable Insulation"},
                {"Installation Method", "PC_Installation Method"},
                {"Cable Additional De-rating", "PC_Cable Additional De-rating"},
                {"Switchgear Trip Unit Type", "PC_Switchgear Trip Unit Type"},
                {"Switchgear Manufacturer", "PC_Switchgear Manufacturer"},
                {"Bus Type", "PC_Bus Type"},
                {"Bus/Chassis Rating (A)", "PC_Bus/Chassis Rating (A)"},
                {"Upstream Diversity", "PC_Upstream Diversity"},
                {"Isolator Type", "PC_Isolator Type"},
                {"Isolator Rating (A)", "PC_Isolator Rating (A)"},
                {"Protective Device Rating (A)", "PC_Protective Device Rating (A)"},
                {"Protective Device Manufacturer", "PC_Protective Device Manufacturer"},
                {"Protective Device Type", "PC_Protective Device Type"},
                {"Protective Device Model", "PC_Protective Device Model"},
                {"Protective Device OCR/Trip Unit", "PC_Protective Device OCR/Trip Unit"},
                {"Protective Device Trip Setting (A)", "PC_Protective Device Trip Setting (A)"}
            };

            // --- 3. Read and Parse CSV Data ---
            List<Dictionary<string, string>> csvData = new List<Dictionary<string, string>>();
            string[] headers = null;
            int swbToCsvIndex = -1;
            Dictionary<string, Dictionary<string, string>> csvLookup = null;

            try
            {
                using (TextFieldParser parser = new TextFieldParser(csvFilePath, Encoding.Default))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    parser.HasFieldsEnclosedInQuotes = true;

                    if (!parser.EndOfData) {
                        headers = parser.ReadFields()?.Select(h => h.Trim()).ToArray();
                        if (headers == null || headers.Length == 0) { message = "CSV header row missing or empty."; TaskDialog.Show("CSV Error", message); return Result.Failed; }
                        swbToCsvIndex = Array.FindIndex(headers, h => h.Equals("SWB To", StringComparison.OrdinalIgnoreCase));
                        if (swbToCsvIndex < 0) { message = "CSV header row does not contain the required 'SWB To' column."; TaskDialog.Show("CSV Error", message); return Result.Failed; }
                    } else { message = "CSV file appears empty (no header row found)."; TaskDialog.Show("CSV Error", message); return Result.Failed; }

                    while (!parser.EndOfData) {
                        string[] fields = parser.ReadFields();
                        if (fields != null && headers != null && fields.Length == headers.Length) {
                            var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            for (int i = 0; i < headers.Length; i++) {
                                bool isRelevantHeader = !string.IsNullOrWhiteSpace(headers[i]) &&
                                                        (i == swbToCsvIndex ||
                                                         parameterMapping.ContainsKey(headers[i]) ||
                                                         headers[i].Equals("Protective Device Model", StringComparison.OrdinalIgnoreCase)); 

                                if (isRelevantHeader) {
                                    string rawValue = fields[i] ?? string.Empty;
                                    string processedValue = rawValue;

                                    if (headers[i].Equals("Cable Reference", StringComparison.OrdinalIgnoreCase)) {
                                        processedValue = rawValue.Replace("=", "").Replace("\"", "").Trim();
                                    } else {
                                        processedValue = rawValue.Trim();
                                    }
                                    rowData[headers[i]] = processedValue;
                                }
                            }
                            if (rowData.ContainsKey("SWB To") && !string.IsNullOrWhiteSpace(rowData["SWB To"])) {
                                csvData.Add(rowData);
                            }
                        } else if (fields != null) { System.Diagnostics.Debug.WriteLine($"Skipping CSV row {parser.LineNumber}: Expected {headers?.Length ?? 0} fields, found {fields.Length}."); }
                    }
                } 

                if (!csvData.Any()) { message = "No valid data rows found in CSV. Check that the 'SWB To' column contains values."; TaskDialog.Show("CSV Warning", message); return Result.Succeeded; }

                csvLookup = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                foreach(var row in csvData) {
                    if (row.TryGetValue("SWB To", out string key) && !string.IsNullOrWhiteSpace(key)) {
                        csvLookup[key] = row;
                    }
                }
            }
            catch (IOException ioEx) { message = $"Error accessing CSV file: {ioEx.Message}\nEnsure the file is not open in another program."; TaskDialog.Show("File Access Error", message); return Result.Failed; }
            catch (Exception ex) { long lineNumber = (ex is MalformedLineException mex) ? mex.LineNumber : -1; message = $"Error reading/parsing CSV" + (lineNumber > 0 ? $" near line {lineNumber}" : "") + $": {ex.Message}"; TaskDialog.Show("CSV Error", message); return Result.Failed; }

            // --- 4. Find and Filter Detail Items ---
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> detailItemsToUpdate;
            try
            {
                detailItemsToUpdate = collector
                    .OfCategory(BuiltInCategory.OST_DetailComponents)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .ToList()
                    .Where(el =>
                    {
                        try
                        {
                            Parameter powerCadParam = el.LookupParameter("PC_PowerCAD");
                            bool powerCadIsTrue = powerCadParam != null
                                               && powerCadParam.HasValue
                                               && powerCadParam.StorageType == StorageType.Integer
                                               && powerCadParam.AsInteger() == 1;

                            if (!powerCadIsTrue) return false;

                            Parameter swbToParam = el.LookupParameter("PC_SWB To");
                            return swbToParam != null
                                && swbToParam.HasValue
                                && !string.IsNullOrWhiteSpace(swbToParam.AsString());
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Error checking parameters for Element ID {el.Id}: {ex.Message}"); return false; }
                    })
                    .ToList();
            }
            catch (Exception ex) { message = $"Error filtering elements: {ex.Message}"; TaskDialog.Show("Revit Element Error", message); return Result.Failed; }

            if (!detailItemsToUpdate.Any()) { TaskDialog.Show("Info", "No Detail Items found matching criteria (PC_PowerCAD=Yes and non-empty PC_SWB To)."); return Result.Succeeded; }


            // --- 5. Update Parameters within a Transaction ---
            int updatedCount = 0;
            int skippedCount = 0;
            List<string> errors = new List<string>();
            using (Transaction tx = new Transaction(doc)) {
                try {
                    tx.Start("Import PowerCAD CSV Data & Set Frame Size");
                    foreach (Element detailItem in detailItemsToUpdate) {
                        Parameter swbToParamRevit = detailItem.LookupParameter("PC_SWB To");
                        string originalSwbToValueRevit = swbToParamRevit?.AsString();

                        if (string.IsNullOrWhiteSpace(originalSwbToValueRevit)) { skippedCount++; continue; }

                        string lookupKey = originalSwbToValueRevit;
                        bool keyModifiedByType = false;

                        Parameter swbTypeParamRevit = detailItem.LookupParameter("PC_SWB Type");
                        string swbTypeValue = swbTypeParamRevit?.AsString();

                        if (!string.IsNullOrWhiteSpace(swbTypeValue) && swbTypeValue.Equals("S", StringComparison.OrdinalIgnoreCase))
                        {
                            keyModifiedByType = true;
                            string startMarker = "(Bus ";
                            string endMarker = ")";
                            int startIndex = originalSwbToValueRevit.IndexOf(startMarker, StringComparison.Ordinal);

                            if (startIndex != -1) {
                                int actualStartIndex = startIndex + startMarker.Length;
                                int endIndex = originalSwbToValueRevit.IndexOf(endMarker, actualStartIndex, StringComparison.Ordinal);
                                if (endIndex != -1) {
                                    lookupKey = originalSwbToValueRevit.Substring(actualStartIndex, endIndex - actualStartIndex).Trim();
                                    if (string.IsNullOrWhiteSpace(lookupKey)) {
                                        errors.Add($"Extracted empty lookup key from SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item.");
                                        skippedCount++; continue;
                                    }
                                } else {
                                    errors.Add($"End marker '{endMarker}' not found after '{startMarker}' in SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item.");
                                    skippedCount++; continue;
                                }
                            } else {
                                errors.Add($"Start marker '{startMarker}' not found in SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item.");
                                skippedCount++; continue;
                            }
                        }

                        if (csvLookup != null && csvLookup.TryGetValue(lookupKey, out var matchingCsvRow)) {
                            bool itemUpdatedThisLoop = false; // Renamed from itemUpdated to avoid scope conflict

                            foreach (var kvp in parameterMapping) {
                                string csvHeader = kvp.Key; string revitParamName = kvp.Value;
                                if (matchingCsvRow.TryGetValue(csvHeader, out string csvValue)) {
                                    string valueToSet = csvValue;

                                    if (csvHeader.Equals("Protective Device Trip Setting (A)", StringComparison.OrdinalIgnoreCase)) {
                                        if (double.TryParse(csvValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue)) {
                                            valueToSet = Math.Ceiling(numericValue).ToString("F0", CultureInfo.InvariantCulture);
                                        } else if (!string.IsNullOrWhiteSpace(csvValue)) { System.Diagnostics.Debug.WriteLine($"Could not parse CSV value '{csvValue}' for rounding ('{csvHeader}' on Elem ID {detailItem.Id}). Using original."); }
                                    }

                                    Parameter targetParam = detailItem.LookupParameter(revitParamName);
                                    if (targetParam != null && !targetParam.IsReadOnly) {
                                        try {
                                            string currentValueStr = GetParameterValueAsString(targetParam);
                                            if (currentValueStr != valueToSet) {
                                                bool success = SetParameterValueFromString(targetParam, valueToSet);
                                                if(success) { itemUpdatedThisLoop = true; } // Use renamed variable
                                                else { errors.Add($"Could not set '{valueToSet}' (processed from '{csvValue}') for '{revitParamName}' on Elem ID {detailItem.Id}. Target type: {targetParam.StorageType}"); }
                                            }
                                        } catch (Exception ex) { errors.Add($"Error setting '{revitParamName}' on Elem ID {detailItem.Id}: {ex.Message}"); }
                                    } else if (targetParam == null) { System.Diagnostics.Debug.WriteLine($"Mapped Param '{revitParamName}' not found on Elem ID {detailItem.Id}."); }
                                }
                            }

                            // --- Extract and Set Frame Size from Protective Device Model (using char.IsDigit) ---
                            string frameSizeRevitParamName = "PC_Frame Size";
                            string protectiveDeviceModelCsvKey = "Protective Device Model";
                            if (matchingCsvRow.TryGetValue(protectiveDeviceModelCsvKey, out string protectiveDeviceModelCsvValue) && !string.IsNullOrWhiteSpace(protectiveDeviceModelCsvValue))
                            {
                                string numericalValueStr = new string(protectiveDeviceModelCsvValue.Where(char.IsDigit).ToArray());

                                if (!string.IsNullOrEmpty(numericalValueStr)) 
                                {
                                    Parameter frameSizeParam = detailItem.LookupParameter(frameSizeRevitParamName);

                                    if (frameSizeParam != null && !frameSizeParam.IsReadOnly)
                                    {
                                        try
                                        {
                                            string currentFrameSizeStr = GetParameterValueAsString(frameSizeParam);
                                            if (currentFrameSizeStr != numericalValueStr)
                                            {
                                                bool success = SetParameterValueFromString(frameSizeParam, numericalValueStr);
                                                if (success) { itemUpdatedThisLoop = true; } // Use renamed variable
                                                else { errors.Add($"Could not set extracted frame size '{numericalValueStr}' for '{frameSizeRevitParamName}' (from CSV model '{protectiveDeviceModelCsvValue}') on Elem ID {detailItem.Id}. Target parameter type: {frameSizeParam.StorageType}."); }
                                            }
                                        }
                                        catch (Exception ex) { errors.Add($"Error setting '{frameSizeRevitParamName}' (value: {numericalValueStr}) on Elem ID {detailItem.Id}: {ex.Message}"); }
                                    }
                                    else if (frameSizeParam == null) { errors.Add($"Parameter '{frameSizeRevitParamName}' not found on Elem ID {detailItem.Id} for Frame Size update."); }
                                    else { errors.Add($"Parameter '{frameSizeRevitParamName}' is read-only on Elem ID {detailItem.Id} for Frame Size update."); }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"No numerical digits found in '{protectiveDeviceModelCsvKey}' value '{protectiveDeviceModelCsvValue}' for Elem ID {detailItem.Id} using char.IsDigit method.");
                                }
                            }
                            // --- END FRAME SIZE LOGIC ---

                            if (itemUpdatedThisLoop) { updatedCount++; } // Increment updatedCount if itemUpdatedThisLoop is true

                        } else {
                            skippedCount++;
                            string originalValueInfo = keyModifiedByType ? $"(original SWB_To: '{originalSwbToValueRevit}')" : "";
                            System.Diagnostics.Debug.WriteLine($"CSV lookup failed for key '{lookupKey}' {originalValueInfo} derived from Elem ID {detailItem.Id}.");
                        }
                    } // End foreach detailItem

                    if (updatedCount > 0 || errors.Any()) { // Commit if any updates made OR if there were errors (to capture partial success with errors)
                        tx.Commit();
                        if (updatedCount == 0 && errors.Any()) { message = "Potential issues encountered. No elements were updated successfully. See error details."; }
                        else if (updatedCount > 0) { message = $"Successfully updated parameters for {updatedCount} elements."; }
                        // If errors.Any() is true and updatedCount > 0, message will be the success message, errors reported separately.
                    } else {
                        tx.RollBack();
                        message = "No parameter changes were necessary based on the CSV data and existing element values.";
                    }

                } catch (Exception ex) {
                    if (tx.GetStatus() == TransactionStatus.Started) tx.RollBack();
                    message = $"Transaction error: {ex.Message}"; TaskDialog.Show("Transaction Error", message); return Result.Failed;
                }
            } // End using Transaction


            // --- 6. Report Results ---
            string summary = $"Import complete.\n\n" +
                             $"Detail Items processed (PC_PowerCAD=Yes, non-empty PC_SWB To): {detailItemsToUpdate.Count}\n" +
                             $"Detail Items with parameters updated: {updatedCount}\n" +
                             $"Detail Items skipped/no changes (No CSV match / No data change required / Error during item processing): {detailItemsToUpdate.Count - updatedCount}\n"; 


            if (!string.IsNullOrEmpty(message) && !(updatedCount > 0 && !errors.Any()) ) { // Append status message if it's not the pure success one or if there are errors
                 if (message != $"Successfully updated parameters for {updatedCount} elements.")
                 {
                    summary += $"\nStatus: {message}\n";
                 }
            }

            if (errors.Any()) {
                summary += $"\nEncountered {errors.Count} errors/warnings during parameter setting:\n";
                summary += string.Join("\n", errors.Take(15)); // Show first 15 errors in the summary dialog
                if (errors.Count > 15) summary += $"\n... ({errors.Count - 15} more)";
                TaskDialog.Show("Import Report with Errors/Warnings", summary);
            } else { TaskDialog.Show("Import Report", summary); }

            // If there were errors and absolutely no items were updated, consider it a failure.
            return (errors.Any() && updatedCount == 0 && detailItemsToUpdate.Any()) ? Result.Failed : Result.Succeeded;
        }


        // --- Helper Methods ---
        private string GetCsvFilePath()
        {
            string filePath = null;
            System.Threading.Thread t = new System.Threading.Thread(() => {
                using (OpenFileDialog openFileDialog = new OpenFileDialog()) {
                    openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1; openFileDialog.RestoreDirectory = true; openFileDialog.Title = "Select PowerCAD Data CSV File";
                    if (openFileDialog.ShowDialog() == DialogResult.OK) filePath = openFileDialog.FileName;
                }
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA); t.Start(); t.Join();
            return filePath;
        }

        private bool SetParameterValueFromString(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly) return false;

            // Optional: Handle blank values for non-string types explicitly if needed
            // if (string.IsNullOrWhiteSpace(value) && param.StorageType != StorageType.String) { /* set to default or clear */ }

            try {
                switch (param.StorageType) {
                    case StorageType.Double:
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal)) {
                            if (Math.Abs(param.AsDouble() - doubleVal) > 0.000001) 
                                param.Set(doubleVal);
                            return true;
                        }
                        break;
                    case StorageType.Integer:
                        Definition def = param.Definition;
                        ForgeTypeId specTypeId = def.GetDataType(); 
                        if (specTypeId == SpecTypeId.Boolean.YesNo) { 
                            int intValToSet = -1; 
                            if (value.Equals("Yes", StringComparison.OrdinalIgnoreCase) || value == "1" || value.ToLowerInvariant() == "true") { intValToSet = 1;}
                            else if (value.Equals("No", StringComparison.OrdinalIgnoreCase) || value == "0" || value.ToLowerInvariant() == "false") { intValToSet = 0;}
                            
                            if (intValToSet != -1) {
                                if (param.AsInteger() != intValToSet) param.Set(intValToSet);
                                return true;
                            }
                        } else { 
                            if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int intVal)) { 
                                if (param.AsInteger() != intVal) param.Set(intVal);
                                return true;
                            }
                            else if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleAsIntCheck) && doubleAsIntCheck == Math.Floor(doubleAsIntCheck)) {
                                int convertedInt = Convert.ToInt32(doubleAsIntCheck);
                                if (param.AsInteger() != convertedInt) param.Set(convertedInt);
                                return true;
                            }
                        }
                        break; 
                    case StorageType.String:
                        string valueToSet = value ?? string.Empty; // Ensure nulls are handled as empty strings if that's the intent
                        if ((param.AsString() ?? string.Empty) != valueToSet) param.Set(valueToSet); 
                        return true;
                    case StorageType.ElementId:
                        System.Diagnostics.Debug.WriteLine($"Warning: Attempting to set ElementId parameter '{param.Definition.Name}' from string. This is usually not supported directly via SetParameterValueFromString.");
                        return false; 
                    default:
                        try { param.Set(value); return true; } catch { /* Fall through to return false */ }
                        break;
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Error during param.Set() for {param.Definition.Name} with value '{value}': {ex.Message}");
                return false;
            }

            System.Diagnostics.Debug.WriteLine($"Could not parse or set value '{value}' for parameter '{param.Definition.Name}' with StorageType {param.StorageType}.");
            return false;
        }


        private string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue) return string.Empty;
            switch (param.StorageType) {
                case StorageType.Double: return param.AsDouble().ToString("G17", CultureInfo.InvariantCulture);
                case StorageType.Integer:
                    Definition def = param.Definition;
                    ForgeTypeId specTypeId = def.GetDataType();
                    if (specTypeId == SpecTypeId.Boolean.YesNo) { 
                        return param.AsInteger() == 1 ? "Yes" : "No"; 
                    }
                    return param.AsInteger().ToString(CultureInfo.InvariantCulture); 
                case StorageType.String: return param.AsString() ?? string.Empty;
                case StorageType.ElementId:
                    ElementId id = param.AsElementId();
                    if (id == null || id == ElementId.InvalidElementId) return string.Empty;
                    // UPDATED LINE: Use id.Value instead of id.IntegerValue
                    return param.AsValueString() ?? id.Value.ToString(CultureInfo.InvariantCulture); 
                default: return param.AsValueString() ?? string.Empty; 
            }
        }

    } // End class PC_SWB_ImporterClass
} // End namespace
