//
// File: PC_Updater.cs
//
// Namespace: PC_Updater
//
// Class: PC_UpdaterClass
//
// Function: This Revit external command allows the user to select a CSV file
//           (typically a PowerCAD export) and updates the 'Cable Length' column
//           within that CSV. The update is performed by looking up the 'Cable Reference'
//           against the 'Model Generated Data' stored in Revit's extensible storage.
//           If an exact match is found, the CSV's 'Cable Length' is overwritten with
//           the corresponding 'Cable Length (m)' from the Model Generated Data.
//           The modified data is then saved to a new CSV file with a '-Updated' suffix.
//           A summary of updated and unmatched cable references is provided to the user.
//
// Author: Kyle Vorster (Modified by AI)
//
// Log:
// - July 2, 2025: Initial creation of PC_Updater script.
// - July 3, 2025: Updated to output CSV in ANSI encoding (Windows-1252) and
//                 to list all affected cable references in the summary report.
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible; // To access ModelGeneratedData schema and recall methods
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms; // For OpenFileDialog, SaveFileDialog
#endregion

namespace PC_Updater
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_UpdaterClass : IExternalCommand
    {
        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Instantiate PC_ExtensibleClass to access its public methods
            PC_ExtensibleClass pcExtensible = new PC_ExtensibleClass();

            // --- 1. Prompt user to select the input CSV file ---
            string inputCsvPath = GetInputCsvFilePath();
            if (string.IsNullOrEmpty(inputCsvPath))
            {
                message = "Operation cancelled by user (no input CSV selected).";
                return Result.Cancelled;
            }

            try
            {
                // --- 2. Parse the input CSV file ---
                // Using Encoding.Default to read the input CSV, which typically corresponds to ANSI on Windows.
                List<InputCsvRecord> inputRecords = ParseInputCsv(inputCsvPath);
                if (inputRecords == null || !inputRecords.Any())
                {
                    TaskDialog.Show("No Data", "No valid data found in the selected CSV file. No updates will be performed.");
                    return Result.Succeeded;
                }

                // --- 3. Recall Model Generated Data from Extensible Storage ---
                List<PC_ExtensibleClass.ModelGeneratedData> modelData = pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                    doc,
                    PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                    PC_ExtensibleClass.ModelGeneratedSchemaName,
                    PC_ExtensibleClass.ModelGeneratedFieldName,
                    PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                );

                // Create a dictionary for quick lookup of Model Generated Data by Cable Reference
                // Handle potential duplicate keys in Model Generated Data by taking the first one
                var modelDataLookup = modelData
                    .Where(m => !string.IsNullOrEmpty(m.CableReference))
                    .GroupBy(m => m.CableReference)
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                // --- 4. Update Cable Lengths in input records ---
                List<string> updatedCableRefs = new List<string>();
                List<string> unmatchedCableRefs = new List<string>();

                foreach (var record in inputRecords)
                {
                    if (string.IsNullOrEmpty(record.CableReference))
                    {
                        // Skip records without a Cable Reference, or handle as an unmatched case if needed
                        continue;
                    }

                    if (modelDataLookup.TryGetValue(record.CableReference, out PC_ExtensibleClass.ModelGeneratedData matchedModelData))
                    {
                        // Found a match, update the Cable Length in the input record
                        record.CableLength = matchedModelData.CableLengthM;
                        updatedCableRefs.Add(record.CableReference);
                    }
                    else
                    {
                        unmatchedCableRefs.Add(record.CableReference);
                    }
                }

                // --- 5. Reproduce the CSV with updated data ---
                string outputFileName = System.IO.Path.GetFileNameWithoutExtension(inputCsvPath) + "-Updated.csv";
                string outputCsvPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(inputCsvPath), outputFileName);

                // Write the output CSV using ANSI encoding (Windows-1252)
                WriteOutputCsv(outputCsvPath, inputRecords);

                // --- 6. Provide Summary to the user ---
                StringBuilder summary = new StringBuilder();
                summary.AppendLine("Cable Length Update Process Complete!");
                summary.AppendLine($"Updated CSV saved to: {outputCsvPath}");
                summary.AppendLine();
                summary.AppendLine($"Total Cable References in input CSV: {inputRecords.Count}");
                summary.AppendLine($"Successfully Updated: {updatedCableRefs.Count} cables");
                summary.AppendLine($"Not Matched in Model Data: {unmatchedCableRefs.Count} cables");
                summary.AppendLine();

                if (updatedCableRefs.Any())
                {
                    summary.AppendLine("Updated Cable References:");
                    foreach (var cableRef in updatedCableRefs.OrderBy(c => c)) // Order for readability
                    {
                        summary.AppendLine($"- {cableRef}");
                    }
                }
                else
                {
                    summary.AppendLine("No cable references were updated.");
                }
                summary.AppendLine();

                if (unmatchedCableRefs.Any())
                {
                    summary.AppendLine("Unmatched Cable References:");
                    foreach (var cableRef in unmatchedCableRefs.OrderBy(c => c)) // Order for readability
                    {
                        summary.AppendLine($"- {cableRef}");
                    }
                }
                else
                {
                    summary.AppendLine("All cable references found in Model Data.");
                }

                TaskDialog.Show("PCAD Updater Summary", summary.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An unexpected error occurred during PCAD Update: {ex.Message}\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Prompts the user to select an input CSV file.
        /// </summary>
        private string GetInputCsvFilePath()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select PowerCAD Export CSV File to Update";
                openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// Parses the input CSV file into a list of InputCsvRecord objects.
        /// It attempts to read the CSV using the default system encoding (typically ANSI/Windows-1252).
        /// </summary>
        private List<InputCsvRecord> ParseInputCsv(string filePath)
        {
            var records = new List<InputCsvRecord>();
            var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); // Case-insensitive header matching
            string[] headers = null;

            using (var reader = new StreamReader(filePath, Encoding.Default)) // Use Encoding.Default for reading
            {
                string line;
                bool headerFound = false;

                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines

                    List<string> values = ParseCsvLine(line);

                    if (!headerFound)
                    {
                        // Attempt to find the header line by looking for "Cable Reference"
                        if (values.Any(v => v.Trim().Equals("Cable Reference", StringComparison.OrdinalIgnoreCase)))
                        {
                            headers = values.Select(v => v.Trim()).ToArray();
                            for (int i = 0; i < headers.Length; i++)
                            {
                                headerMap[headers[i]] = i;
                            }
                            headerFound = true;
                            continue; // Skip the header line itself
                        }
                        // If header not found yet, keep reading lines until it is
                        continue;
                    }

                    // Ensure we have enough values for the expected headers
                    if (headers == null || values.Count < headers.Length)
                    {
                        // This line is malformed or not data, skip it
                        continue;
                    }

                    var record = new InputCsvRecord();
                    // Populate properties using the header map
                    record.CableReference = GetValueFromCsv(values, headerMap, "Cable Reference");
                    record.SWBFrom = GetValueFromCsv(values, headerMap, "SWB From");
                    record.SWBTo = GetValueFromCsv(values, headerMap, "SWB To");
                    record.SWBType = GetValueFromCsv(values, headerMap, "SWB Type");
                    record.SWBLoad = GetValueFromCsv(values, headerMap, "SWB Load");
                    record.SWBLoadScope = GetValueFromCsv(values, headerMap, "SWB Load Scope");
                    record.SWBPF = GetValueFromCsv(values, headerMap, "SWB PF");
                    record.CableLength = GetValueFromCsv(values, headerMap, "Cable Length"); // This is the target for update
                    record.CableSizeActiveConductors = GetValueFromCsv(values, headerMap, "Cable Size - Active conductors");
                    record.CableSizeNeutralConductors = GetValueFromCsv(values, headerMap, "Cable Size - Neutral conductors");
                    record.CableSizeEarthingConductor = GetValueFromCsv(values, headerMap, "Cable Size - Earthing conductor");
                    record.ActiveConductorMaterial = GetValueFromCsv(values, headerMap, "Active Conductor material");
                    record.NumberOfPhases = GetValueFromCsv(values, headerMap, "# of Phases");
                    record.CableType = GetValueFromCsv(values, headerMap, "Cable Type");
                    record.CableInsulation = GetValueFromCsv(values, headerMap, "Cable Insulation");
                    record.InstallationMethod = GetValueFromCsv(values, headerMap, "Installation Method");
                    record.CableAdditionalDerating = GetValueFromCsv(values, headerMap, "Cable Additional De-rating");
                    record.SwitchgearTripUnitType = GetValueFromCsv(values, headerMap, "Switchgear Trip Unit Type");
                    record.SwitchgearManufacturer = GetValueFromCsv(values, headerMap, "Switchgear Manufacturer");
                    record.BusType = GetValueFromCsv(values, headerMap, "Bus Type");
                    record.BusChassisRatingA = GetValueFromCsv(values, headerMap, "Bus/Chassis Rating (A)");
                    record.UpstreamDiversity = GetValueFromCsv(values, headerMap, "Upstream Diversity");
                    record.IsolatorType = GetValueFromCsv(values, headerMap, "Isolator Type");
                    record.IsolatorRatingA = GetValueFromCsv(values, headerMap, "Isolator Rating (A)");
                    record.ProtectiveDeviceRatingA = GetValueFromCsv(values, headerMap, "Protective Device Rating (A)");
                    record.ProtectiveDeviceManufacturer = GetValueFromCsv(values, headerMap, "Protective Device Manufacturer");
                    record.ProtectiveDeviceType = GetValueFromCsv(values, headerMap, "Protective Device Type");
                    record.ProtectiveDeviceModel = GetValueFromCsv(values, headerMap, "Protective Device Model");
                    record.ProtectiveDeviceOCRTripUnit = GetValueFromCsv(values, headerMap, "Protective Device OCR/Trip Unit");
                    record.ProtectiveDeviceTripSettingA = GetValueFromCsv(values, headerMap, "Protective Device Trip Setting (A)");

                    records.Add(record);
                }
            }
            return records;
        }

        /// <summary>
        /// Helper to parse a single CSV line, handling quoted fields and commas within quotes.
        /// </summary>
        private List<string> ParseCsvLine(string line)
        {
            var fields = new List<string>();
            var fieldBuilder = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i < line.Length - 1 && line[i + 1] == '"')
                    {
                        fieldBuilder.Append('"'); // Escaped quote
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes; // Toggle quote state
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(fieldBuilder.ToString());
                    fieldBuilder.Clear();
                }
                else
                {
                    fieldBuilder.Append(c);
                }
            }
            fields.Add(fieldBuilder.ToString()); // Add the last field
            return fields;
        }

        /// <summary>
        /// Gets a value from a CSV row given the header name and header map.
        /// Returns string.Empty if header or index is not found.
        /// </summary>
        private string GetValueFromCsv(List<string> values, Dictionary<string, int> headerMap, string headerName)
        {
            if (headerMap.TryGetValue(headerName, out int index) && index < values.Count)
            {
                return values[index].Trim();
            }
            return string.Empty;
        }

        /// <summary>
        /// Writes the list of InputCsvRecord objects back to a CSV file.
        /// The output file will be encoded in ANSI (Windows-1252).
        /// </summary>
        private void WriteOutputCsv(string filePath, List<InputCsvRecord> records)
        {
            var sb = new StringBuilder();

            // Manually define headers in the exact order as the input CSV for consistency
            // This ensures the output CSV structure matches the input, even if some headers
            // weren't explicitly mapped in the parsing step (though they should be).
            var headers = new List<string>
            {
                "Cable Reference", "SWB From", "SWB To", "SWB Type", "SWB Load", "SWB Load Scope",
                "SWB PF", "Cable Length", "Cable Size - Active conductors", "Cable Size - Neutral conductors",
                "Cable Size - Earthing conductor", "Active Conductor material", "# of Phases", "Cable Type",
                "Cable Insulation", "Installation Method", "Cable Additional De-rating", "Switchgear Trip Unit Type",
                "Switchgear Manufacturer", "Bus Type", "Bus/Chassis Rating (A)", "Upstream Diversity",
                "Isolator Type", "Isolator Rating (A)", "Protective Device Rating (A)", "Protective Device Manufacturer",
                "Protective Device Type", "Protective Device Model", "Protective Device OCR/Trip Unit",
                "Protective Device Trip Setting (A)"
            };
            sb.AppendLine(string.Join(",", headers.Select(h => EscapeCsvField(h))));

            // Write data rows
            foreach (var record in records)
            {
                var rowValues = new List<string>
                {
                    record.CableReference, record.SWBFrom, record.SWBTo, record.SWBType, record.SWBLoad,
                    record.SWBLoadScope, record.SWBPF, record.CableLength, record.CableSizeActiveConductors,
                    record.CableSizeNeutralConductors, record.CableSizeEarthingConductor, record.ActiveConductorMaterial,
                    record.NumberOfPhases, record.CableType, record.CableInsulation, record.InstallationMethod,
                    record.CableAdditionalDerating, record.SwitchgearTripUnitType, record.SwitchgearManufacturer,
                    record.BusType, record.BusChassisRatingA, record.UpstreamDiversity, record.IsolatorType,
                    record.IsolatorRatingA, record.ProtectiveDeviceRatingA, record.ProtectiveDeviceManufacturer,
                    record.ProtectiveDeviceType, record.ProtectiveDeviceModel, record.ProtectiveDeviceOCRTripUnit,
                    record.ProtectiveDeviceTripSettingA
                };
                sb.AppendLine(string.Join(",", rowValues.Select(v => EscapeCsvField(v))));
            }

            // Write the string builder content to the file using ANSI encoding (code page 1252)
            File.WriteAllText(filePath, sb.ToString(), Encoding.GetEncoding(1252));
        }

        /// <summary>
        /// Escapes a string for CSV output (handles commas, quotes, newlines).
        /// </summary>
        private string EscapeCsvField(string field)
        {
            if (field == null) return string.Empty;
            field = field.Trim();
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        /// <summary>
        /// Represents a single row of data from the input CSV file.
        /// </summary>
        private class InputCsvRecord
        {
            public string CableReference { get; set; }
            public string SWBFrom { get; set; }
            public string SWBTo { get; set; }
            public string SWBType { get; set; }
            public string SWBLoad { get; set; }
            public string SWBLoadScope { get; set; }
            public string SWBPF { get; set; }
            public string CableLength { get; set; } // This is the property that will be updated
            public string CableSizeActiveConductors { get; set; }
            public string CableSizeNeutralConductors { get; set; }
            public string CableSizeEarthingConductor { get; set; }
            public string ActiveConductorMaterial { get; set; }
            public string NumberOfPhases { get; set; }
            public string CableType { get; set; }
            public string CableInsulation { get; set; }
            public string InstallationMethod { get; set; }
            public string CableAdditionalDerating { get; set; }
            public string SwitchgearTripUnitType { get; set; }
            public string SwitchgearManufacturer { get; set; }
            public string BusType { get; set; }
            public string BusChassisRatingA { get; set; }
            public string UpstreamDiversity { get; set; }
            public string IsolatorType { get; set; }
            public string IsolatorRatingA { get; set; }
            public string ProtectiveDeviceRatingA { get; set; }
            public string ProtectiveDeviceManufacturer { get; set; }
            public string ProtectiveDeviceType { get; set; }
            public string ProtectiveDeviceModel { get; set; }
            public string ProtectiveDeviceOCRTripUnit { get; set; }
            public string ProtectiveDeviceTripSettingA { get; set; }

            // Constructor to initialize all string properties to empty string
            public InputCsvRecord()
            {
                foreach (PropertyInfo prop in typeof(InputCsvRecord).GetProperties())
                {
                    if (prop.PropertyType == typeof(string) && prop.CanWrite)
                    {
                        prop.SetValue(this, string.Empty);
                    }
                }
            }
        }
    }
}
