#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

// Required for ForgeTypeId (Revit 2021+)
#if REVIT2021_OR_GREATER
using Autodesk.Revit.DB.Visual;
#endif
#endregion

// Ensure this namespace matches your project structure if needed
namespace PC_Cable_Importer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    // Ensure this class name matches how it's used in your .addin file
    public class PC_Cable_ImporterClass : IExternalCommand
    {
        // --- Constants ---
        private const string SWB_TO_PARAM_NAME = "PC_SWB To"; // Key parameter for matching Detail Items to CSV rows
        private const string POWERCAD_FILTER_PARAM_NAME = "PC_PowerCAD"; // Yes/No parameter for pre-filtering Detail Items
        private const string CSV_KEY_COLUMN_HEADER = "To"; // Header name in CSV used for matching and identifying data rows

        // --- Parameter Mapping Dictionary ---
        // Maps CSV Header Name to Revit Parameter Name
        private readonly Dictionary<string, string> _parameterMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) // Case-insensitive header matching
        {
            // CSV Header                               Revit Parameter Name (or null/empty if not mapped)
            { "Cable Reference",                        "PC_Cable Reference" },
            { "From",                                   "PC_SWB From" },
            { "To",                                     SWB_TO_PARAM_NAME }, // Used for matching, but also mapped if needed
            { "Cable Type",                             "PC_Cable Type" },
            { "Cores",                                  "PC_Cores" }, // Expecting Integer
            { "Cable Size (mm²)",                       "PC_Cable Size - Active conductors" }, // Expecting String
            { "Conductor (Active)",                     "PC_Active Conductor material" },
            { "Insulation",                             "PC_Cable Insulation" },
            { "Neutral Size (mm²)",                     "PC_Cable Size - Neutral conductors" }, // Expecting String
            { "Earth Size (mm²)",                       "PC_Cable Size - Earthing conductor" }, // Expecting String
            { "Conductor (Earth)",                      "PC_Earth Conductor material" },
            { "Separate Earth for Multicore",           "PC_Separate Earth for Multicore" }, // Expecting Yes/No (Integer 1/0)
            { "Cable Length (m)",                       "PC_Cable Length" }, // Expecting Double (Length)
            { "Accum. Voltage Drop % (Incl. FSC)",      "PC_Accum Volt Drop Incl FSC" }, // Expecting Double
            { "Prospective Fault at End of Cable (kA)", "PC_Prospective Fault at End of Cable" }, // Expecting Double
            { "Addition Cable Derating",                "PC_Cable Additional De-rating" }, // Expecting Double
            { "No. of Conduits",                        "PC_No. of Conduits" }, // Expecting Integer
            { "Conduit Size (mm)",                      "PC_Conduit Size" }, // Expecting String
            { "Design Progress",                        "PC_Design Progress" },
            { "Nominal Overall Diameter (mm)",          "PC_Nominal Overall Diameter" } // Expecting Double (Length)
            // Add or remove mappings as needed based on the previous examples and your requirements
        };

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            // *** FIX: Explicitly qualify Application type ***
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // 1. Prompt user for CSV file
            string csvFilePath = PromptForCsvFile();
            if (string.IsNullOrEmpty(csvFilePath))
            {
                return Result.Cancelled;
            }

            // 2. Read and Parse CSV Data
            List<Dictionary<string, string>> csvData = null;
            List<string> errors = new List<string>();
            try
            {
                // Using Encoding.Default for ANSI. The ParseCsv method will now explicitly remove \ufeff characters.
                csvData = ParseCsv(csvFilePath, Encoding.Default);
            }
            catch (IOException ex)
            {
                message = $"Error reading file: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = $"Error parsing CSV: {ex.Message}";
                // Show parsing errors immediately as they might prevent proceeding
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            if (csvData == null || csvData.Count == 0)
            {
                TaskDialog.Show("Warning", "No valid data rows found in the CSV file (identified by non-empty 'To' column).");
                return Result.Cancelled;
            }

            // Create a lookup dictionary for faster matching CSV Row based on "To" value
            Dictionary<string, Dictionary<string, string>> csvLookup = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in csvData)
            {
                // The key CSV_KEY_COLUMN_HEADER ("To") should exist due to parsing logic
                if (row.TryGetValue(CSV_KEY_COLUMN_HEADER, out string toValue) && !string.IsNullOrWhiteSpace(toValue))
                {
                    // Handle potential duplicate 'To' values in CSV - use the last one found or log a warning
                    if (csvLookup.ContainsKey(toValue))
                    {
                        errors.Add($"Warning: Duplicate '{CSV_KEY_COLUMN_HEADER}' value '{toValue}' found in CSV. Using the last occurrence.");
                    }
                    csvLookup[toValue] = row; // Add or overwrite
                }
            }


            // 3. Get Detail Items from Revit
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> detailItems = collector.OfCategory(BuiltInCategory.OST_DetailComponents)
                                                 .WhereElementIsNotElementType()
                                                 .ToElements();

            if (!detailItems.Any())
            {
                TaskDialog.Show("Warning", "No Detail Items found in the current document.");
                return Result.Cancelled;
            }

            // 4. Process Elements and Update Parameters
            int updatedCount = 0;
            int processedCount = 0;
            int eligibleCount = 0; // Count items passing the PC_PowerCAD filter


            using (Transaction tx = new Transaction(doc))
            {
                try
                {
                    tx.Start("Update Detail Items from CSV");

                    foreach (Element elem in detailItems)
                    {
                        processedCount++;

                        // *** PRE-FILTERING STEP ***
                        Parameter powerCadParam = elem.LookupParameter(POWERCAD_FILTER_PARAM_NAME);

                        // Check if the PC_PowerCAD parameter exists and is set to 1 (Yes)
                        if (powerCadParam != null && powerCadParam.HasValue && powerCadParam.AsInteger() == 1)
                        {
                            eligibleCount++; // This item is eligible for update

                            // Now check the matching parameter PC_SWB To
                            Parameter swbToParam = elem.LookupParameter(SWB_TO_PARAM_NAME);

                            // Check if the element has the key matching parameter and it has a value
                            if (swbToParam != null && swbToParam.HasValue && !string.IsNullOrWhiteSpace(swbToParam.AsString()))
                            {
                                string revitToValue = swbToParam.AsString();

                                // Find matching row in CSV data using the lookup
                                if (csvLookup.TryGetValue(revitToValue, out Dictionary<string, string> matchingCsvRow))
                                {
                                    bool itemUpdated = false;
                                    // Iterate through the parameter map
                                    foreach (var kvp in _parameterMap)
                                    {
                                        string csvHeader = kvp.Key;
                                        string revitParamName = kvp.Value;

                                        // Skip if no Revit parameter is mapped
                                        if (string.IsNullOrEmpty(revitParamName))
                                            continue;

                                        // Check if the CSV row contains this header
                                        if (matchingCsvRow.TryGetValue(csvHeader, out string csvValue) && csvValue != null) // Don't process nulls from parsing
                                        {
                                            Parameter targetParam = elem.LookupParameter(revitParamName);
                                            if (targetParam != null && !targetParam.IsReadOnly)
                                            {
                                                try
                                                {
                                                    // Pass the 'errors' list to log detailed conversion issues
                                                    bool success = SetParameterValue(targetParam, csvValue, errors, elem.Id.ToString(), revitToValue);
                                                    if (success) itemUpdated = true;
                                                }
                                                catch (Exception ex)
                                                {
                                                    // Log error setting specific parameter
                                                    errors.Add($"Error setting parameter '{revitParamName}' on element ID {elem.Id} (To='{revitToValue}') with CSV value '{csvValue}': {ex.Message}");
                                                }
                                            }
                                            // Optional: Log warnings for non-critical issues during update attempt
                                            // else if (targetParam == null) { errors.Add($"Warning: Parameter '{revitParamName}' not found on element ID {elem.Id}."); }
                                            // else if (targetParam.IsReadOnly) { errors.Add($"Warning: Parameter '{revitParamName}' is read-only on element ID {elem.Id}."); }
                                        }
                                    }
                                    if (itemUpdated) updatedCount++;
                                }
                                else // CSV match not found for this eligible item
                                {
                                    errors.Add($"Info: Eligible Detail Item ID {elem.Id} (PC_PowerCAD=Yes) with '{SWB_TO_PARAM_NAME}' = '{revitToValue}' not found in CSV's '{CSV_KEY_COLUMN_HEADER}' column.");
                                }
                            }
                            else // Eligible item missing the SWB_TO parameter or value
                            {
                                errors.Add($"Warning: Eligible Detail Item ID {elem.Id} (PC_PowerCAD=Yes) has missing or empty parameter '{SWB_TO_PARAM_NAME}'. Cannot match to CSV.");
                            }
                        }
                        // else: Item skipped because PC_PowerCAD is not 'Yes' or parameter missing. Do nothing.

                    } // End foreach Element loop

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    message = $"Transaction failed: {ex.Message}";
                    errors.Add($"Fatal Error during transaction: {ex.Message}");
                    TaskDialog.Show("Error", message);
                    // Optionally show detailed errors here too
                    if (errors.Any())
                    {
                        ShowErrorsDialog(errors);
                    }
                    return Result.Failed;
                }
            }

            // 5. Report Results
            StringBuilder report = new StringBuilder();
            report.AppendLine($"Iterated through {processedCount} Detail Items.");
            report.AppendLine($"Found {eligibleCount} items marked for update ({POWERCAD_FILTER_PARAM_NAME}=Yes).");
            report.AppendLine($"Successfully updated parameters on {updatedCount} of these eligible items based on CSV match.");

            if (errors.Any())
            {
                report.AppendLine($"{errors.Count} warnings/errors encountered during processing (see details).");
                TaskDialog.Show("Import Summary", report.ToString());
                ShowErrorsDialog(errors); // Show detailed errors in a separate dialog
            }
            else
            {
                report.AppendLine("Import completed successfully.");
                TaskDialog.Show("Import Summary", report.ToString());
            }


            return Result.Succeeded;
        }

        // --- Helper Methods ---

        /// <summary>
        /// Prompts the user to select a CSV file.
        /// </summary>
        /// <returns>The full path to the selected file, or null if cancelled.</returns>
        private string PromptForCsvFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.Title = "Select Powercalc SWB Export CSV File";
                openFileDialog.RestoreDirectory = true;

                // Need to run this on a separate thread if called from a modeless context,
                // but for a standard IExternalCommand, it should be fine.
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// Parses a CSV file, handling quoted fields with commas.
        /// Assumes the first non-empty line containing the key header ("To") is the header.
        /// Identifies data rows as those following the header where the "To" column is not empty.
        /// This method has been updated to remove all '\ufeff' (Byte Order Mark) characters from each line using a loop.
        /// </summary>
        /// <param name="filePath">Path to the CSV file.</param>
        /// <param name="encoding">File encoding (e.g., Encoding.Default for ANSI).</param>
        /// <returns>A list of dictionaries, where each dictionary represents a row (Header -> Value).</returns>
        private List<Dictionary<string, string>> ParseCsv(string filePath, Encoding encoding)
        {
            var data = new List<Dictionary<string, string>>();
            string[] headers = null;
            bool headerFound = false;
            int toColumnIndex = -1; // Index of the key column ("To") used to identify data rows

            using (StreamReader reader = new StreamReader(filePath, encoding)) // Use provided encoding
            {
                string line;
                int lineNum = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    lineNum++;
                    // *** MODIFICATION: Manually remove all \ufeff characters from the line using a loop ***
                    StringBuilder cleanedLineBuilder = new StringBuilder();
                    foreach (char c in line)
                    {
                        if (c != '\uFEFF')
                        {
                            cleanedLineBuilder.Append(c);
                        }
                    }
                    line = cleanedLineBuilder.ToString();

                    if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines

                    var values = ParseCsvLine(line);
                    if (values.Length == 0) continue; // Skip lines that parse to nothing

                    if (!headerFound)
                    {
                        // Try to find the header row - assuming it contains the key "To" column header
                        int tempIndex = Array.FindIndex(values, h => h.Trim().Equals(CSV_KEY_COLUMN_HEADER, StringComparison.OrdinalIgnoreCase));
                        if (tempIndex >= 0)
                        {
                            headers = values.Select(h => h.Trim()).ToArray(); // Trim headers
                            toColumnIndex = tempIndex;
                            headerFound = true;
                        }
                        continue; // Skip line if header not found yet
                    }

                    // --- Processing lines AFTER header is found ---

                    // Check if row has enough columns based on header count
                    if (values.Length != headers.Length)
                    {
                        Console.WriteLine($"Warning Line {lineNum}: Row skipped due to mismatched column count. Expected {headers.Length}, got {values.Length}. Line: {line}");
                        continue; // Skip malformed row
                    }

                    // *** Data Row Identification: Check if the 'To' column is populated ***
                    string toValueRaw = values[toColumnIndex];
                    if (!string.IsNullOrWhiteSpace(toValueRaw))
                    {
                        // This looks like a data row, process it
                        var rowDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < headers.Length; i++)
                        {
                            string value = values[i].Trim();
                            if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 1)
                            {
                                value = value.Substring(1, value.Length - 2).Replace("\"\"", "\""); // Handle escaped quotes "" inside
                            }
                            rowDict[headers[i]] = value; // Use OrdinalIgnoreCase comparer
                        }
                        data.Add(rowDict);
                    }
                    // else: Line skipped, likely metadata or invalid data row.

                } // end while loop
            } // end using reader

            if (!headerFound) throw new Exception($"CSV parsing error: Header row containing the key column '{CSV_KEY_COLUMN_HEADER}' not found.");
            // No need to check toColumnIndex < 0 as headerFound implies it was set

            return data;
        }


        /// <summary>
        /// Basic CSV line parser that handles commas within quoted fields.
        /// </summary>
        private string[] ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            StringBuilder currentField = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // Skip the second quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            fields.Add(currentField.ToString()); // Add the last field
            return fields.ToArray();
        }


        /// <summary>
        /// Sets the value of a Revit parameter, attempting type conversion and handling units.
        /// Assumes Revit 2021+ API for ForgeTypeId.
        /// </summary>
        /// <param name="param">The Revit parameter to set.</param>
        /// <param name="value">The string value from the CSV.</param>
        /// <param name="errorLog">List to add detailed errors to.</param>
        /// <param name="elementIdStr">Element ID for context in error messages.</param>
        /// <param name="elementToStr">Element 'To' value for context in error messages.</param>
        /// <returns>True if setting was successful, false otherwise.</returns>
        private bool SetParameterValue(Parameter param, string value, List<string> errorLog, string elementIdStr, string elementToStr)
        {
            if (param == null || param.IsReadOnly || value == null)
            {
                return false; // Cannot set or null value provided
            }

            // Handle empty strings
            if (string.IsNullOrWhiteSpace(value))
            {
                if (param.StorageType == StorageType.String)
                {
                    return param.Set(""); // Set empty string for Text parameters
                }
                else
                {
                    // Cannot set empty for non-string types typically. Log or ignore.
                    // errorLog?.Add($"Info: Skipped setting parameter '{param.Definition.Name}' on Element ID {elementIdStr} (To='{elementToStr}') due to empty CSV value.");
                    return false;
                }
            }

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.Double:
                        if (double.TryParse(value, out double dblValue))
                        {
                            ForgeTypeId specTypeId = param.Definition.GetDataType();
                            // Check if it's a unit-based type requiring potential conversion
                            if (UnitUtils.IsMeasurableSpec(specTypeId))
                            {
                                // --- Unit Conversion Logic (Revit 2021+) ---
                                if (specTypeId == SpecTypeId.Length)
                                {
                                    // Assuming CSV "Cable Length (m)" and "Nominal Overall Diameter (mm)"
                                    if (param.Definition.Name.Equals("PC_Cable Length", StringComparison.OrdinalIgnoreCase))
                                    {
                                        double internalValue = UnitUtils.ConvertToInternalUnits(dblValue, UnitTypeId.Meters);
                                        return param.Set(internalValue);
                                    }
                                    else if (param.Definition.Name.Equals("PC_Nominal Overall Diameter", StringComparison.OrdinalIgnoreCase))
                                    {
                                        double internalValue = UnitUtils.ConvertToInternalUnits(dblValue, UnitTypeId.Millimeters);
                                        return param.Set(internalValue);
                                    }
                                    else
                                    {
                                        // Default for other lengths: Assume meters or needs specific handling
                                        errorLog?.Add($"Warning: Unhandled Length parameter '{param.Definition.Name}' on Elem ID {elementIdStr}. Assuming meters for value '{dblValue}'.");
                                        double internalValue = UnitUtils.ConvertToInternalUnits(dblValue, UnitTypeId.Meters);
                                        return param.Set(internalValue);
                                    }
                                }
                                // Add other 'else if' blocks for Area, Volume, Angle etc. if needed
                                // else if (specTypeId == SpecTypeId.Area) { ... }
                                else
                                {
                                    // Default for other measurable specs: Assume direct setting works (e.g., unitless number used for a spec) or needs specific handling
                                    // errorLog?.Add($"Info: Setting measurable spec '{param.Definition.Name}' ({specTypeId.ToUnitTypeId()}) directly with value {dblValue} on Elem ID {elementIdStr}. Verify units if needed.");
                                    return param.Set(dblValue);
                                }
                            }
                            else // Not a measurable spec (e.g., just Number, Currency, etc.)
                            {
                                return param.Set(dblValue); // Set directly
                            }
                        }
                        else
                        {
                            errorLog?.Add($"Error: Could not parse '{value}' as Double for parameter '{param.Definition.Name}' on Element ID {elementIdStr} (To='{elementToStr}').");
                            return false; // Parsing failed
                        }

                    case StorageType.Integer:
                        ForgeTypeId intSpecTypeId = param.Definition.GetDataType();
                        bool isYesNo = intSpecTypeId == SpecTypeId.Boolean.YesNo;

                        if (isYesNo)
                        {
                            string trimmedValue = value.Trim().ToLowerInvariant();
                            if (trimmedValue == "yes" || trimmedValue == "true" || trimmedValue == "1")
                                return param.Set(1);
                            else if (trimmedValue == "no" || trimmedValue == "false" || trimmedValue == "0")
                                return param.Set(0);
                            else
                            {
                                errorLog?.Add($"Error: Could not parse '{value}' as Yes/No for parameter '{param.Definition.Name}' on Element ID {elementIdStr} (To='{elementToStr}').");
                                return false; // Cannot interpret Yes/No value
                            }
                        }
                        else if (int.TryParse(value, out int intValue)) // Regular integer
                        {
                            return param.Set(intValue);
                        }
                        else
                        {
                            errorLog?.Add($"Error: Could not parse '{value}' as Integer for parameter '{param.Definition.Name}' on Element ID {elementIdStr} (To='{elementToStr}').");
                            return false; // Parsing failed for integer
                        }

                    case StorageType.String:
                        return param.Set(value); // Already trimmed and unquoted during parsing

                    case StorageType.ElementId:
                        // Setting ElementId from CSV is generally not safe/implemented here
                        errorLog?.Add($"Warning: Skipped setting ElementId parameter '{param.Definition.Name}' on Element ID {elementIdStr}. Not supported by this script.");
                        return false;

                    default:
                        errorLog?.Add($"Warning: Skipped setting parameter '{param.Definition.Name}' on Element ID {elementIdStr}. Unknown StorageType: {param.StorageType}.");
                        return false; // Unknown storage type
                }
            }
            catch (Exception ex)
            {
                // Log the exception details for debugging
                errorLog?.Add($"Exception setting parameter '{param.Definition.Name}' on Element ID {elementIdStr} (To='{elementToStr}') with value '{value}'. Exception: {ex.Message}");
                Console.WriteLine($"Error setting parameter {param.Definition.Name}: {ex.Message}");
                return false; // Indicate failure
            }
        }

        /// <summary>
        /// Displays a scrollable dialog with error/warning messages.
        /// </summary>
        private void ShowErrorsDialog(List<string> errors)
        {
            if (errors == null || !errors.Any()) return;

            TaskDialog errorDialog = new TaskDialog("Import Issues");
            errorDialog.MainInstruction = "The following warnings or errors occurred during the import process:";
            // Join errors with new lines for display. Limit total length if necessary.
            string errorDetails = string.Join(Environment.NewLine, errors);
            const int maxLength = 2000; // Approx limit for TaskDialog
            if (errorDetails.Length > maxLength)
            {
                errorDetails = errorDetails.Substring(0, maxLength) + "... (more issues truncated)";
            }
            errorDialog.MainContent = errorDetails;
            errorDialog.CommonButtons = TaskDialogCommonButtons.Close;
            errorDialog.Show();
        }
    }
}
