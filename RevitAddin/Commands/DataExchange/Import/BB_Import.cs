// Required Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO; // For File operations
using System.Windows.Forms; // For OpenFileDialog
using System.Globalization; // For CultureInfo.InvariantCulture
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB; // Revit Database Namespace
using Autodesk.Revit.UI; // Revit User Interface Namespace
using Autodesk.Revit.UI.Selection;

// Updated namespace based on previous code snippet
namespace RTS.Commands.DataExchange.Import
{
    [Transaction(TransactionMode.Manual)] // Use Manual transaction mode
    [Regeneration(RegenerationOption.Manual)]
    public class BB_CableLengthImport : IExternalCommand // Or use directly in Macro Manager
    {
        // --- Constants ---
        private const string SUBJECT_FILTER = "Polylength Measurement";
        private const string SUBJECT_COLUMN_NAME = "Subject"; // Case-sensitive match with CSV Header
        private const string LABEL_COLUMN_NAME = "Label";     // Case-sensitive match with CSV Header
        private const string MEASUREMENT_COLUMN_NAME = "Measurement"; // Case-sensitive match with CSV Header
        private const string DETAIL_ITEM_PARAM_SWB_TO = "PC_SWB To";
        private const string DETAIL_ITEM_PARAM_CABLE_LENGTH = "PC_Cable Length";
        // Removed METERS_TO_FEET_CONVERSION_FACTOR

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            // Use fully qualified name to avoid ambiguity
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // --- 1. Prompt user for CSV file ---
            string csvFilePath = GetCsvFilePath();
            if (string.IsNullOrEmpty(csvFilePath))
            {
                message = "Operation cancelled: No CSV file selected.";
                return Result.Cancelled;
            }

            // --- 2. Process CSV Data ---
            // Dictionary stores summed measurements (in original units from CSV, e.g., Meters)
            Dictionary<string, double> labelSummedMeasurements;
            try
            {
                labelSummedMeasurements = ProcessCsvFile(csvFilePath);
                if (labelSummedMeasurements == null || !labelSummedMeasurements.Any())
                {
                     TaskDialog.Show("CSV Processing", "No valid data found in the CSV file matching the criteria (Subject='Polylength Measurement', non-blank Label).");
                     return Result.Failed; // Or Succeeded if no action is desired
                }
            }
            catch (Exception ex)
            {
                message = $"Error processing CSV file: {ex.Message}";
                TaskDialog.Show("CSV Error", message);
                return Result.Failed;
            }

            // --- 3. Find and Update Detail Items ---
            int updatedCount = 0;
            int skippedMissingParamSwb = 0;
            int skippedMissingParamLength = 0;
            int skippedNoMatch = 0;
            int settingErrors = 0; // Renamed counter

            // Use a FilteredElementCollector to get all Detail Item instances
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_DetailComponents);
            collector.WhereElementIsNotElementType(); // Get instances, not types

            // Start a Revit Transaction to make changes
            // Updated Transaction Name
            using (Transaction t = new Transaction(doc, "Update Cable Lengths Text Param from CSV"))
            {
                try
                {
                    t.Start();

                    foreach (Element elem in collector)
                    {
                        // Attempt to get the "PC_SWB To" parameter
                        Parameter swbToParam = elem.LookupParameter(DETAIL_ITEM_PARAM_SWB_TO);
                        if (swbToParam == null || !swbToParam.HasValue)
                        {
                             skippedMissingParamSwb++;
                             continue; // Skip if parameter doesn't exist or has no value
                        }

                        string swbToValue = swbToParam.AsValueString() ?? swbToParam.AsString(); // Get value as string
                         if (string.IsNullOrWhiteSpace(swbToValue))
                         {
                             skippedMissingParamSwb++; // Count items where the parameter exists but value is blank
                             continue;
                         }


                        // Check if this label exists in our processed CSV data
                        // Renamed variable storing the looked-up value
                        if (labelSummedMeasurements.TryGetValue(swbToValue, out double summedMeasurementValue))
                        {
                            // Label matches, now check for the "PC_Cable Length" parameter
                            Parameter cableLengthParam = elem.LookupParameter(DETAIL_ITEM_PARAM_CABLE_LENGTH);

                            // Check parameter exists, is not read-only, and IS A TEXT PARAMETER
                            if (cableLengthParam != null &&
                                !cableLengthParam.IsReadOnly &&
                                cableLengthParam.StorageType == StorageType.String) // Verify it's a Text parameter
                            {
                                try
                                {
                                    // **CONVERT TO STRING**
                                    // Convert the summed double value to a string.
                                    // Using InvariantCulture ensures '.' decimal separator.
                                    // Use formatting like "F2" for specific decimal places if needed:
                                    // string valueAsString = summedMeasurementValue.ToString("F2", CultureInfo.InvariantCulture);
                                    string valueAsString = summedMeasurementValue.ToString(CultureInfo.InvariantCulture);

                                    // Set the Text parameter value using the string
                                    bool resultSet = cableLengthParam.Set(valueAsString);

                                    if (resultSet)
                                    {
                                       updatedCount++;
                                    }
                                    else
                                    {
                                        // If Set() returns false, it might indicate an issue (e.g., value too long for text param?)
                                        System.Diagnostics.Debug.WriteLine($"Warning: Parameter.Set(string) returned false for element {elem.Id}, Label '{swbToValue}'. Check parameter constraints. Value='{valueAsString}'");
                                        settingErrors++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Log or report error setting parameter for this specific element
                                    System.Diagnostics.Debug.WriteLine($"Error setting '{DETAIL_ITEM_PARAM_CABLE_LENGTH}' (Text) for element {elem.Id} with label '{swbToValue}': {ex.Message}");
                                    settingErrors++;
                                }
                            }
                            else
                            {
                                // "PC_Cable Length" parameter doesn't exist, is read-only, or is not a Text parameter
                                if(cableLengthParam != null && cableLengthParam.StorageType != StorageType.String)
                                {
                                     System.Diagnostics.Debug.WriteLine($"Skipping element {elem.Id}, Label '{swbToValue}': Parameter '{DETAIL_ITEM_PARAM_CABLE_LENGTH}' is not a Text parameter (Type: {cableLengthParam.StorageType}).");
                                }
                                skippedMissingParamLength++;
                            }
                        }
                        else
                        {
                             // The value in "PC_SWB To" didn't match any Label in the processed CSV data
                             skippedNoMatch++;
                        }
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    message = $"Error during Revit element update: {ex.Message}";
                    TaskDialog.Show("Revit Update Error", message);
                    t.RollBack(); // Roll back changes on error
                    return Result.Failed;
                }
            } // End using Transaction

            // --- 4. Report Results ---
            string summary = $"Update Complete.\n\n" +
                             $"CSV File: {Path.GetFileName(csvFilePath)}\n" +
                             $"Detail Items Updated: {updatedCount}\n" +
                             $"Items Skipped (No matching Label found): {skippedNoMatch}\n" +
                             $"Items Skipped (Missing/Blank '{DETAIL_ITEM_PARAM_SWB_TO}' parameter): {skippedMissingParamSwb}\n" +
                             $"Items Skipped (Missing/Read-only/Wrong Type for '{DETAIL_ITEM_PARAM_CABLE_LENGTH}'): {skippedMissingParamLength}\n" + // Updated reason
                             (settingErrors > 0 ? $"Potential Errors during parameter setting: {settingErrors}\n" : ""); // Updated reason


            TaskDialog.Show("Cable Length Update Summary", summary);

            return Result.Succeeded;
        }

        /// <summary>
        /// Prompts the user to select a CSV file using a standard dialog.
        /// </summary>
        /// <returns>The full path to the selected CSV file, or null if cancelled.</returns>
        private string GetCsvFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                // Updated title slightly to remove unit emphasis if stored as text
                openFileDialog.Title = "Select CSV File with Measurement Data";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null; // User cancelled
        }

        /// <summary>
        /// Reads the specified CSV file, filters, aggregates data, and returns a dictionary.
        /// Assumes the first row is the header.
        /// Key: Label (string), Value: Summed Measurement (double, in original units from CSV)
        /// </summary>
        /// <param name="filePath">Full path to the CSV file.</param>
        /// <returns>Dictionary of Label -> Summed Measurement, or null on critical error.</returns>
        /// <exception cref="FileNotFoundException">Thrown if the file doesn't exist.</exception>
        /// <exception cref="IOException">Thrown on general I/O error.</exception>
        /// <exception cref="Exception">Thrown on parsing or header errors.</exception>
        private Dictionary<string, double> ProcessCsvFile(string filePath)
        {
            // Dictionary stores summed numeric values
            var labelSums = new Dictionary<string, double>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("CSV file not found.", filePath);
            }

            string[] lines = File.ReadAllLines(filePath);

            if (lines.Length < 2) // Need header + at least one data row
            {
                System.Diagnostics.Debug.WriteLine("CSV file has less than 2 lines (header + data).");
                return labelSums; // Return empty dictionary
            }

            // --- Parse Header (Row 0) ---
            string headerLine = lines[0];
            string[] headers = headerLine.Split(',');

            // Find column indices based on header names (case-insensitive)
            int subjectIndex = Array.FindIndex(headers, h => h.Trim().Equals(SUBJECT_COLUMN_NAME, StringComparison.OrdinalIgnoreCase));
            int labelIndex = Array.FindIndex(headers, h => h.Trim().Equals(LABEL_COLUMN_NAME, StringComparison.OrdinalIgnoreCase));
            int measurementIndex = Array.FindIndex(headers, h => h.Trim().Equals(MEASUREMENT_COLUMN_NAME, StringComparison.OrdinalIgnoreCase));


            if (subjectIndex == -1 || labelIndex == -1 || measurementIndex == -1)
            {
                 // Construct a more informative error message
                List<string> missingColumns = new List<string>();
                if (subjectIndex == -1) missingColumns.Add(SUBJECT_COLUMN_NAME);
                if (labelIndex == -1) missingColumns.Add(LABEL_COLUMN_NAME);
                if (measurementIndex == -1) missingColumns.Add(MEASUREMENT_COLUMN_NAME);

                throw new Exception($"CSV header (first row) is missing required column(s): '{string.Join("', '", missingColumns)}'. " +
                                    $"Please ensure the first row contains headers exactly matching these names (case-insensitive check performed). " +
                                    $"Found headers: '{string.Join(",", headers)}'");
            }


            // --- Process Data Rows (Starting from Row 1) ---
            for (int i = 1; i < lines.Length; i++) // Start from index 1 to skip header
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines

                string[] values = line.Split(',');

                // Basic check for column count consistency
                 if (values.Length <= Math.Max(subjectIndex, Math.Max(labelIndex, measurementIndex)))
                {
                     System.Diagnostics.Debug.WriteLine($"Skipping row {i + 1}: Incorrect number of columns. Expected at least {Math.Max(subjectIndex, Math.Max(labelIndex, measurementIndex)) + 1}, found {values.Length}. Line: '{line}'");
                     continue; // Skip rows that don't have enough columns based on required indices
                }


                // Trim values to remove leading/trailing whitespace
                string subject = values[subjectIndex].Trim();
                string label = values[labelIndex].Trim();
                string measurementStr = values[measurementIndex].Trim();

                // Apply Filters
                if (subject.Equals(SUBJECT_FILTER, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(label))
                {
                    // Try Parse Measurement (as double)
                    // Use InvariantCulture to handle '.' as decimal separator regardless of system settings
                    // Renamed variable for parsed value
                    if (double.TryParse(measurementStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double measurementValue))
                    {
                        // Add or Update the sum in the dictionary
                        if (labelSums.ContainsKey(label))
                        {
                            labelSums[label] += measurementValue;
                        }
                        else
                        {
                            labelSums.Add(label, measurementValue);
                        }
                    }
                    else
                    {
                         System.Diagnostics.Debug.WriteLine($"Skipping row {i + 1}: Could not parse '{measurementStr}' as a double in column '{MEASUREMENT_COLUMN_NAME}'. Line: '{line}'");
                        // Optionally log this error or inform the user later
                    }
                }
            }

            return labelSums;
        }
    }
}