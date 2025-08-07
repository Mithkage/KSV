// ### PREPROCESSOR DIRECTIVES MUST BE AT THE VERY TOP ###
#if REVIT2024_OR_GREATER || REVIT2022 // Use ForgeTypeId for 2022 and newer
#define USE_FORGE_TYPE_ID
#else
    // This error will trigger if a recognized Revit version symbol isn't defined
#error "Revit compilation symbol (REVIT2024_OR_GREATER or REVIT2022) not defined in project build settings."
#endif

// Conditionally use ForgeTypeId alias
#if USE_FORGE_TYPE_ID
using ForgeTypeId = Autodesk.Revit.DB.ForgeTypeId;
#endif

// Standard using statements will go below this block
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO; // For TextFieldParser
using System.Globalization;

namespace RTS.Commands
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
            // Application app = uiapp.Application; // 'app' is declared but never used. Consider removing if not needed in full context.
            Document doc = uidoc.Document;

            string csvFilePath = GetCsvFilePath();
            if (string.IsNullOrEmpty(csvFilePath))
            {
                message = "Operation cancelled by user.";
                return Result.Cancelled;
            }

            Dictionary<string, string> parameterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"Cable Reference", "PC_Cable Reference"}, {"SWB From", "PC_SWB From"}, {"SWB To", "PC_SWB To"},
                {"SWB Type", "PC_SWB Type"}, {"SWB Load", "PC_SWB Load"}, {"SWB Load Scope", "PC_SWB Load Scope"},
                {"SWB PF", "PC_SWB PF"}, {"Cable Length", "PC_Cable Length"},
                {"Cable Size - Active conductors", "PC_Cable Size - Active conductors"},
                {"Cable Size - Neutral conductors", "PC_Cable Size - Neutral conductors"},
                {"Cable Size - Earthing conductor", "PC_Cable Size - Earthing conductor"},
                {"Active Conductor material", "PC_Active Conductor material"}, {"# of Phases", "PC_# of Phases"},
                {"Cable Insulation", "PC_Cable Insulation"}, {"Installation Method", "PC_Installation Method"},
                {"Cable Additional De-rating", "PC_Cable Additional De-rating"}, // Assuming this parameter exists
                {"Switchgear Trip Unit Type", "PC_Switchgear Trip Unit Type"},
                {"Switchgear Manufacturer", "PC_Switchgear Manufacturer"}, {"Bus Type", "PC_Bus Type"},
                {"Bus/Chassis Rating (A)", "PC_Bus/Chassis Rating (A)"}, {"Upstream Diversity", "PC_Upstream Diversity"},
                {"Isolator Type", "PC_Isolator Type"}, {"Isolator Rating (A)", "PC_Isolator Rating (A)"},
                {"Protective Device Rating (A)", "PC_Protective Device Rating (A)"}, // Already in original list
                {"Protective Device Manufacturer", "PC_Protective Device Manufacturer"},
                {"Protective Device Type", "PC_Protective Device Type"},
                {"Protective Device Model", "PC_Protective Device Model"},
                {"Protective Device OCR/Trip Unit", "PC_Protective Device OCR/Trip Unit"},
                {"Protective Device Trip Setting (A)", "PC_Protective Device Trip Setting (A)"}
            };

            List<Dictionary<string, string>> csvData = new List<Dictionary<string, string>>();
            string[] headers = null;
            int swbToCsvIndex = -1;
            Dictionary<string, Dictionary<string, string>> csvLookup = null;

            try
            {
                using (TextFieldParser parser = new TextFieldParser(csvFilePath, Encoding.Default)) // Consider Encoding.UTF8 for wider compatibility
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    parser.HasFieldsEnclosedInQuotes = true;

                    if (!parser.EndOfData)
                    {
                        headers = parser.ReadFields()?.Select(h => h.Trim()).ToArray();
                        if (headers == null || headers.Length == 0) { message = "CSV header row missing or empty."; TaskDialog.Show("CSV Error", message); return Result.Failed; }
                        swbToCsvIndex = Array.FindIndex(headers, h => h.Equals("SWB To", StringComparison.OrdinalIgnoreCase));
                        if (swbToCsvIndex < 0) { message = "CSV header row does not contain the required 'SWB To' column."; TaskDialog.Show("CSV Error", message); return Result.Failed; }
                    }
                    else { message = "CSV file appears empty (no header row found)."; TaskDialog.Show("CSV Error", message); return Result.Failed; }

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (fields != null && headers != null && fields.Length == headers.Length)
                        {
                            var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            for (int i = 0; i < headers.Length; i++)
                            {
                                bool isRelevantHeader = !string.IsNullOrWhiteSpace(headers[i]) &&
                                                        (i == swbToCsvIndex ||
                                                         parameterMapping.ContainsKey(headers[i]) ||
                                                         headers[i].Equals("Protective Device Model", StringComparison.OrdinalIgnoreCase));
                                if (isRelevantHeader)
                                {
                                    string rawFieldValue = fields[i] ?? string.Empty;
                                    string processedValue;
                                    if (headers[i].Equals("Cable Reference", StringComparison.OrdinalIgnoreCase))
                                    {
                                        processedValue = rawFieldValue.Replace("=", "").Replace("\"", "").Trim();
                                    }
                                    else
                                    {
                                        processedValue = CleanPotentialExcelFormulaOutput(rawFieldValue);
                                    }
                                    rowData[headers[i]] = processedValue;
                                }
                            }
                            if (rowData.ContainsKey("SWB To") && !string.IsNullOrWhiteSpace(rowData["SWB To"]))
                            {
                                csvData.Add(rowData);
                            }
                        }
                        else if (fields != null) { System.Diagnostics.Debug.WriteLine($"Skipping CSV row {parser.LineNumber}: Expected {headers?.Length ?? 0} fields, found {fields.Length}."); }
                    }
                }

                if (!csvData.Any()) { message = "No valid data rows found in CSV. Check that the 'SWB To' column contains values."; TaskDialog.Show("CSV Warning", message); return Result.Succeeded; }

                csvLookup = csvData
                    .Where(r => r.TryGetValue("SWB To", out string k) && !string.IsNullOrWhiteSpace(k))
                    .GroupBy(r => r["SWB To"], StringComparer.OrdinalIgnoreCase) // Handle potential duplicate SWB To keys in CSV
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase); // Take first if duplicates
            }
            catch (IOException ioEx) { message = $"Error accessing CSV file: {ioEx.Message}\nEnsure the file is not open in another program."; TaskDialog.Show("File Access Error", message); return Result.Failed; }
            catch (Exception ex) { long lineNumber = ex is MalformedLineException mex ? mex.LineNumber : -1; message = $"Error reading/parsing CSV" + (lineNumber > 0 ? $" near line {lineNumber}" : "") + $": {ex.Message}"; TaskDialog.Show("CSV Error", message); return Result.Failed; }

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
                            bool powerCadIsTrue = powerCadParam?.HasValue == true &&
                                                  powerCadParam.StorageType == StorageType.Integer &&
                                                  powerCadParam.AsInteger() == 1;
                            if (!powerCadIsTrue) return false;
                            Parameter swbToParam = el.LookupParameter("PC_SWB To");
                            return swbToParam?.HasValue == true && !string.IsNullOrWhiteSpace(swbToParam.AsString());
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Error checking parameters for Element ID {el.Id}: {ex.ToString()}"); return false; }
                    })
                    .ToList();
            }
            catch (Exception ex) { message = $"Error filtering elements: {ex.Message}"; TaskDialog.Show("Revit Element Error", message); return Result.Failed; }

            if (!detailItemsToUpdate.Any()) { TaskDialog.Show("Info", "No Detail Items found matching criteria (PC_PowerCAD=Yes and non-empty PC_SWB To)."); return Result.Succeeded; }

            int updatedCount = 0;
            List<string> errors = new List<string>();
            using (Transaction tx = new Transaction(doc))
            {
                try
                {
                    tx.Start("Import PowerCAD CSV Data & Set Frame Size");
                    foreach (Element detailItem in detailItemsToUpdate)
                    {
                        Parameter swbToParamRevit = detailItem.LookupParameter("PC_SWB To");
                        string originalSwbToValueRevit = swbToParamRevit?.AsString();
                        if (string.IsNullOrWhiteSpace(originalSwbToValueRevit)) { continue; }

                        string lookupKey = originalSwbToValueRevit;
                        // bool keyModifiedByType = false; // CS0219: Variable assigned but never used. Commented out.
                        // If needed for debug, uncomment its usage and the variable itself.
                        Parameter swbTypeParamRevit = detailItem.LookupParameter("PC_SWB Type");
                        string swbTypeValue = swbTypeParamRevit?.AsString();

                        if (!string.IsNullOrWhiteSpace(swbTypeValue) && swbTypeValue.Equals("S", StringComparison.OrdinalIgnoreCase))
                        {
                            // keyModifiedByType = true;
                            string startMarker = "(Bus ";
                            string endMarker = ")";
                            int startIndex = originalSwbToValueRevit.IndexOf(startMarker, StringComparison.Ordinal);
                            if (startIndex != -1)
                            {
                                int actualStartIndex = startIndex + startMarker.Length;
                                int endIndex = originalSwbToValueRevit.IndexOf(endMarker, actualStartIndex, StringComparison.Ordinal);
                                if (endIndex != -1)
                                {
                                    lookupKey = originalSwbToValueRevit.Substring(actualStartIndex, endIndex - actualStartIndex).Trim();
                                    if (string.IsNullOrWhiteSpace(lookupKey)) { errors.Add($"Extracted empty lookup key from SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item."); continue; }
                                }
                                else { errors.Add($"End marker '{endMarker}' not found after '{startMarker}' in SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item."); continue; }
                            }
                            else { errors.Add($"Start marker '{startMarker}' not found in SWB_To '{originalSwbToValueRevit}' for 'S' type on Elem ID {detailItem.Id}. Skipping item."); continue; }
                        }

                        if (csvLookup != null && csvLookup.TryGetValue(lookupKey, out var matchingCsvRow))
                        {
                            bool itemUpdatedThisLoop = false;
                            foreach (var kvp in parameterMapping)
                            {
                                string csvHeader = kvp.Key; string revitParamName = kvp.Value;
                                if (matchingCsvRow.TryGetValue(csvHeader, out string csvValue))
                                {
                                    string valueToSet = csvValue;
                                    if (csvHeader.Equals("Protective Device Trip Setting (A)", StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (double.TryParse(csvValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                                        { valueToSet = Math.Ceiling(numericValue).ToString("F0", CultureInfo.InvariantCulture); }
                                        else if (!string.IsNullOrWhiteSpace(csvValue))
                                        { System.Diagnostics.Debug.WriteLine($"Could not parse CSV value '{csvValue}' for rounding ('{csvHeader}' on Elem ID {detailItem.Id}). Using original cleaned value: '{valueToSet}'."); }
                                    }
                                    Parameter targetParam = detailItem.LookupParameter(revitParamName);
                                    if (targetParam != null && !targetParam.IsReadOnly)
                                    {
                                        try
                                        {
                                            string currentValueStr = GetParameterValueAsString(targetParam);
                                            if (currentValueStr != valueToSet)
                                            {
                                                if (SetParameterValueFromString(targetParam, valueToSet)) { itemUpdatedThisLoop = true; }
                                                else { errors.Add($"Could not set '{valueToSet}' (from CSV field '{csvHeader}':'{csvValue}') for '{revitParamName}' on Elem ID {detailItem.Id}. Target type: {targetParam.StorageType}"); }
                                            }
                                        }
                                        catch (Exception ex) { errors.Add($"Error setting '{revitParamName}' on Elem ID {detailItem.Id}: {ex.ToString()}"); }
                                    }
                                    else if (targetParam == null) { System.Diagnostics.Debug.WriteLine($"Mapped Param '{revitParamName}' not found on Elem ID {detailItem.Id}."); }
                                }
                            }

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
                                            if (GetParameterValueAsString(frameSizeParam) != numericalValueStr)
                                            {
                                                if (SetParameterValueFromString(frameSizeParam, numericalValueStr)) { itemUpdatedThisLoop = true; }
                                                else { errors.Add($"Could not set extracted frame size '{numericalValueStr}' for '{frameSizeRevitParamName}' (from CSV model '{protectiveDeviceModelCsvValue}') on Elem ID {detailItem.Id}. Target type: {frameSizeParam.StorageType}."); }
                                            }
                                        }
                                        catch (Exception ex) { errors.Add($"Error setting '{frameSizeRevitParamName}' (value: {numericalValueStr}) on Elem ID {detailItem.Id}: {ex.ToString()}"); }
                                    }
                                    else if (frameSizeParam == null) { errors.Add($"Parameter '{frameSizeRevitParamName}' not found on Elem ID {detailItem.Id} for Frame Size update."); }
                                    else { errors.Add($"Parameter '{frameSizeRevitParamName}' is read-only on Elem ID {detailItem.Id} for Frame Size update."); }
                                }
                                else { System.Diagnostics.Debug.WriteLine($"No numerical digits found in '{protectiveDeviceModelCsvKey}' value '{protectiveDeviceModelCsvValue}' for Elem ID {detailItem.Id} using char.IsDigit method."); }
                            }
                            if (itemUpdatedThisLoop) { updatedCount++; }
                        }
                        else
                        {
                            // string originalValueInfo = keyModifiedByType ? $"(original SWB_To: '{originalSwbToValueRevit}')" : ""; // keyModifiedByType commented out
                            string originalValueInfo = $"(original SWB_To: '{originalSwbToValueRevit}')"; // Simplified debug output
                            System.Diagnostics.Debug.WriteLine($"CSV lookup failed for key '{lookupKey}' {originalValueInfo} derived from Elem ID {detailItem.Id}.");
                        }
                    } // End foreach detailItem

                    if (updatedCount > 0 || errors.Any())
                    {
                        tx.Commit();
                        if (updatedCount == 0 && errors.Any()) { message = "Potential issues encountered. No elements were updated successfully. See error details."; }
                        else if (updatedCount > 0) { message = $"Successfully updated parameters for {updatedCount} elements."; }
                    }
                    else
                    {
                        tx.RollBack();
                        message = "No parameter changes were necessary based on the CSV data and existing element values.";
                    }
                }
                catch (Exception ex)
                {
                    if (tx.GetStatus() == TransactionStatus.Started) tx.RollBack();
                    message = $"Transaction error: {ex.ToString()}"; // Use ex.ToString() for more detail
                    TaskDialog.Show("Transaction Error", message); return Result.Failed;
                }
            } // End using Transaction

            string summary = $"Import complete.\n\n" +
                             $"Detail Items processed (PC_PowerCAD=Yes, non-empty PC_SWB To): {detailItemsToUpdate.Count}\n" +
                             $"Detail Items with parameters updated: {updatedCount}\n" +
                             $"Detail Items skipped/no changes: {detailItemsToUpdate.Count - updatedCount}\n";

            if (!string.IsNullOrEmpty(message) && !(updatedCount > 0 && !errors.Any()) && message != "No parameter changes were necessary based on the CSV data and existing element values." && message != $"Successfully updated parameters for {updatedCount} elements.")
            {
                summary += $"\nStatus: {message}\n";
            }

            if (errors.Any())
            {
                summary += $"\nEncountered {errors.Count} errors/warnings during parameter setting (see Visual Studio Debug Output for all):\n";
                summary += string.Join("\n", errors.Take(15)); // Show first 15 errors in dialog
                if (errors.Count > 15) summary += $"\n... ({errors.Count - 15} more errors/warnings logged)";
                TaskDialog.Show("Import Report with Errors/Warnings", summary);
            }
            else { TaskDialog.Show("Import Report", summary); }

            return errors.Any() && updatedCount == 0 && detailItemsToUpdate.Any() ? Result.Failed : Result.Succeeded;
        }

        private string CleanPotentialExcelFormulaOutput(string csvFieldValue)
        {
            if (string.IsNullOrWhiteSpace(csvFieldValue)) return csvFieldValue;
            string value = csvFieldValue.Trim();
            if (value.StartsWith("=\"") && value.EndsWith("\"") && value.Length >= 3) return value.Substring(2, value.Length - 3);
            if (value.StartsWith("=")) return value.Substring(1);
            return value;
        }

        private string GetCsvFilePath()
        {
            string filePath = null;
            System.Threading.Thread t = new System.Threading.Thread(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    Title = "Select PowerCAD Data CSV File"
                })
                { if (openFileDialog.ShowDialog() == DialogResult.OK) { filePath = openFileDialog.FileName; } }
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA); t.Start(); t.Join();
            return filePath;
        }

        private bool SetParameterValueFromString(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly) return false;
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.Double:
                        return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal) && param.Set(dVal);
                    case StorageType.Integer:
#if USE_FORGE_TYPE_ID // True for Revit 2022+
                        ForgeTypeId specTypeId = param.Definition.GetDataType();
#else
                        // Fallback logic for ParameterType if supporting pre-2022
                        // This will throw if USE_FORGE_TYPE_ID is not defined, as per current preprocessor.
                        throw new NotSupportedException("ParameterType check logic for pre-2022 needed here.");
#endif

                        if (specTypeId == SpecTypeId.Boolean.YesNo)
                        {
                            if (value.Equals("Yes", StringComparison.OrdinalIgnoreCase) || value == "1" || value.ToLowerInvariant() == "true") return param.Set(1);
                            if (value.Equals("No", StringComparison.OrdinalIgnoreCase) || value == "0" || value.ToLowerInvariant() == "false") return param.Set(0);
                            if (string.IsNullOrWhiteSpace(value)) return param.Set(0); // Default blank to No for Yes/No
                            return false;
                        }
                        // For general integers, try parsing as int, then as double (for cases like "123.0")
                        if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int iVal)) return param.Set(iVal);
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dblVal) && dblVal == Math.Floor(dblVal))
                            return param.Set(Convert.ToInt32(dblVal));
                        return false;
                    case StorageType.String:
                        return param.Set(value ?? string.Empty);
                    case StorageType.ElementId:
                        // *** CORRECTED BLOCK START ***
#if REVIT2024_OR_GREATER || REVIT2023
                        // For Revit 2023 and newer, ElementId constructor takes long (Int64)
                        return long.TryParse(value, out long idValLong) && param.Set(new ElementId(idValLong));
#elif REVIT2022
                        // For Revit 2022, ElementId constructor takes int (Int32)
                        return int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int idValInt) && param.Set(new ElementId(idValInt));
#else
                        // Fallback for older versions - typically uses int (Int32)
                        // Add an #error or specific handling if you don't support older versions.
                        return int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int idValInt) && param.Set(new ElementId(idValInt));
#endif
                        // *** CORRECTED BLOCK END ***
                    default:
                        // Fallback: try setting as string if Revit can convert it.
                        // This might not work for all types and can throw exceptions if conversion is not possible.
                        System.Diagnostics.Debug.WriteLine($"Attempting fallback param.Set(string) for {param.Definition.Name} with value '{value}' and StorageType {param.StorageType}");
                        return param.Set(value);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetParameterValueFromString for {param.Definition.Name} with value '{value}' (StorageType: {param.StorageType}): {ex.ToString()}");
                return false;
            }
        }

        private string GetParameterValueAsString(Parameter param)
        {
            if (param == null || !param.HasValue) return string.Empty;
            switch (param.StorageType)
            {
                case StorageType.Double:
                    return param.AsDouble().ToString("G17", CultureInfo.InvariantCulture);
                case StorageType.Integer:
#if USE_FORGE_TYPE_ID // True for Revit 2022+
                    ForgeTypeId specTypeId = param.Definition.GetDataType();
                    if (specTypeId == SpecTypeId.Boolean.YesNo)
                    { return param.AsInteger() == 1 ? "Yes" : "No"; }
#else
                    // ParameterType legacyType = param.Definition.ParameterType;
                    // if(legacyType == ParameterType.YesNo) { return param.AsInteger() == 1 ? "Yes" : "No"; }
                    throw new NotSupportedException("ParameterType check logic for pre-2022 needed here.");
#endif
                    return param.AsInteger().ToString(CultureInfo.InvariantCulture);
                case StorageType.String:
                    return param.AsString() ?? string.Empty;
                case StorageType.ElementId:
                    ElementId id = param.AsElementId();
                    if (id == null || id == ElementId.InvalidElementId) return string.Empty;

                    string idRepresentationText;
#if REVIT2024_OR_GREATER || REVIT2023
                        idRepresentationText = id.Value.ToString(CultureInfo.InvariantCulture);
#elif REVIT2022
                    idRepresentationText = id.IntegerValue.ToString(CultureInfo.InvariantCulture);
#else
                        // Fallback for older or misconfigured (should be caught by top #error if symbols aren't set)
                        idRepresentationText = id.IntegerValue.ToString(CultureInfo.InvariantCulture);
#endif
                    // AsValueString() can often give a more meaningful representation (e.g., element name)
                    return param.AsValueString() ?? idRepresentationText;
                default:
                    return param.AsValueString() ?? string.Empty;
            }
        }

    } // End class PC_SWB_ImporterClass
} // End namespace