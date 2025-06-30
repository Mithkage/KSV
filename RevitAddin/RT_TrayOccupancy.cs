//
// File: RT_TrayOccupancy.cs
//
// Namespace: RT_TrayOccupancy
//
// Class: RT_TrayOccupancyClass
//
// Function: This Revit external command retrieves stored cable data from project extensible storage.
//           It then scans the active Revit model for Cable Trays with specific parameters,
//           calculates total cable weight and occupancy for each tray using conditional logic,
//           reports on missing cables, determines minimum tray size, and exports that data
//           to a CSV. Finally, it calculates the total run length for each unique cable
//           across all trays and fittings, exports that to a second CSV, and writes the
//           calculated data back to the Revit model.
//
// Author: Kyle Vorster (Modified by AI)
//
// Date: June 30, 2025 (Updated to use Extensible Storage)
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage; // Required for Extensible Storage
using System.Text.Json; // Required for JSON deserialization
using System.Diagnostics; // For Debug.WriteLine
#endregion

namespace RT_TrayOccupancy
{
    /// <summary>
    /// The main class for the Revit external command.
    /// Implements the IExternalCommand interface.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_TrayOccupancyClass : IExternalCommand
    {
        // Define a unique GUID for your Schema. This GUID must be truly unique for your application.
        // It should match the one in PC_Extensible.cs
        private static readonly Guid SchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336");
        private const string FieldName = "PC_DataJson"; // Field to store the JSON string (must match PC_Extensible)

        /// <summary>
        /// The main entry point for the external command. Revit calls this method when the user clicks the button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- 1. RECALL CLEANED DATA FROM EXTENSIBLE STORAGE ---
                List<CableData> cleanedData = RecallCableDataFromExtensibleStorage(doc);

                if (cleanedData == null || cleanedData.Count == 0)
                {
                    TaskDialog.Show("No Stored Data", "No valid cable data was found in the project's extensible storage. Please run the 'Process & Save Cable Data' command first to import data.");
                    return Result.Succeeded;
                }

                // --- 2. PROMPT USER FOR OUTPUT FOLDER ---
                string outputFolderPath = GetOutputFolderPath();
                if (string.IsNullOrEmpty(outputFolderPath))
                {
                    return Result.Cancelled;
                }

                // --- 3. EXPORT THE RECALLED DATA TO A NEW CSV (Optional, but good for verification) ---
                string cleanedScheduleFilePath = Path.Combine(outputFolderPath, "Cleaned_Cable_Schedule.csv");
                ExportDataToCsv(cleanedData, cleanedScheduleFilePath);

                // --- 4. PROCESS AND EXPORT CABLE TRAY DATA FROM REVIT MODEL ---
                List<TrayCableData> trayData = CollectTrayData(doc, cleanedData);
                if (trayData.Any())
                {
                    string trayDataFilePath = Path.Combine(outputFolderPath, "CableTray_Data.csv");
                    ExportTrayDataToCsv(trayData, trayDataFilePath);

                    // --- NEW STEP: UPDATE REVIT MODEL PARAMETERS ---
                    UpdateRevitParameters(doc, trayData);
                }

                // --- 5. PROCESS AND EXPORT CABLE LENGTH DATA ---
                List<CableLengthData> cableLengths = CollectCableLengthsData(doc, cleanedData);
                if (cableLengths.Any())
                {
                    string cableLengthsFilePath = Path.Combine(outputFolderPath, "Cable_Lengths.csv");
                    ExportCableLengthsToCsv(cableLengths, cableLengthsFilePath);
                }

                // --- 6. NOTIFY USER OF SUCCESS ---
                TaskDialog.Show("Success", $"Process complete. Revit model has been updated and files have been saved to:\n{outputFolderPath}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An unexpected error occurred: {ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        #region Extensible Storage Recall Method

        /// <summary>
        /// Recalls the cleaned cable data (PC_Data) from extensible storage.
        /// This method is copied from PC_Extensible.cs to ensure data consistency.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>A List of CableData, or an empty list if no data is found or an error occurs during recall.</returns>
        private List<CableData> RecallCableDataFromExtensibleStorage(Document doc)
        {
            Schema schema = Schema.Lookup(SchemaGuid); // Look up the schema by its GUID

            if (schema == null)
            {
                // Schema does not exist in this project yet, so no data is stored under this schema.
                return new List<CableData>();
            }

            // Find existing DataStorage elements associated with our schema
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                // Check if the DataStorage element contains an entity for our schema
                if (ds.GetEntity(schema) != null)
                {
                    dataStorage = ds;
                    break;
                }
            }

            if (dataStorage == null)
            {
                // No DataStorage element found containing our schema's data
                return new List<CableData>();
            }

            // Get the Entity from the DataStorage
            Entity entity = dataStorage.GetEntity(schema);

            if (!entity.IsValid())
            {
                return new List<CableData>();
            }

            // Get the JSON string from the field
            string jsonString = entity.Get<string>(schema.GetField(FieldName));

            if (string.IsNullOrEmpty(jsonString))
            {
                return new List<CableData>();
            }

            // Deserialize the JSON string back into a List<CableData>
            try
            {
                return JsonSerializer.Deserialize<List<CableData>>(jsonString) ?? new List<CableData>();
            }
            catch (JsonException ex)
            {
                TaskDialog.Show("Data Recall Error", $"Failed to deserialize stored PC_Data: {ex.Message}. The stored data might be corrupt or incompatible.");
                return new List<CableData>();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Data Recall Error", $"An unexpected error occurred during PC_Data recall: {ex.Message}");
                return new List<CableData>();
            }
        }

        #endregion

        #region File Dialog Methods
        // Removed GetSourceCsvFilePath() as it's no longer needed
        private string GetOutputFolderPath()
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a Folder to Save the Exported Files";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // Default to My Documents

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return null;
        }
        #endregion

        #region CSV Exporting
        // Removed ParseCsvLine and ParseAndProcessCsvData as CSV reading is no longer needed
        // The ProcessSplit and GetValueOrDefault methods are also no longer needed for CSV parsing
        // but are kept if other parts of the script still rely on them for string manipulation.
        // If not, they can be removed.
        private void ProcessSplit(string inputValue, out string countPart, out string sizePart)
        {
            string[] parts = inputValue.Split(new[] { '×', 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 1)
            {
                countPart = parts[0].Trim();
                sizePart = parts[1].Trim();
            }
            else
            {
                countPart = "1";
                sizePart = inputValue.Trim();
            }
        }

        private string GetValueOrDefault(string[] values, Dictionary<string, int> map, string headerName)
        {
            if (map.TryGetValue(headerName, out int index) && index < values.Length)
            {
                return values[index].Trim();
            }
            return string.Empty;
        }


        private void ExportDataToCsv(List<CableData> data, string filePath)
        {
            var sb = new StringBuilder();

            // Headers for the exported cleaned data CSV
            var headers = new List<string>
            {
                "Cable Reference", "From", "To", "Cable Type", "Cable Code", "Cable Configuration",
                "Cores", "Number of Active Cables", "Active Cable Size (mm\u00B2)", "Conductor (Active)",
                "Insulation", "Number of Neutral Cables", "Neutral Cable Size (mm\u00B2)",
                "Number of Earth Cables", "Earth Cable Size (mm\u00B2)", "Conductor (Earth)",
                "Separate Earth for Multicore", "Cable Length (m)",
                "Total Cable Run Weight (Incl. N & E) (kg)", "Cables kg per m", "Nominal Overall Diameter (mm)", "AS/NSZ 3008 Cable Derating Factor"
            };

            sb.AppendLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var line = new List<string>
                {
                    row.CableReference, row.From, row.To, row.CableType, row.CableCode,
                    row.CableConfiguration, row.Cores, row.NumberOfActiveCables, row.ActiveCableSize,
                    row.ConductorActive, row.Insulation, row.NumberOfNeutralCables, row.NeutralCableSize,
                    row.NumberOfEarthCables, row.EarthCableSize, row.ConductorEarth,
                    row.SeparateEarthForMulticore, row.CableLength, row.TotalCableRunWeight,
                    row.CablesKgPerM, row.NominalOverallDiameter, row.AsNsz3008CableDeratingFactor
                };

                // Format each value for CSV, handling commas by enclosing in quotes
                var formattedLine = line.Select(val =>
                {
                    if (val == null) return string.Empty;
                    val = val.Trim(); // Trim whitespace
                    if (val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r"))
                    {
                        // Escape double quotes by doubling them, then enclose the whole field in quotes
                        return $"\"{val.Replace("\"", "\"\"")}\"";
                    }
                    return val;
                });
                sb.AppendLine(string.Join(",", formattedLine));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
        #endregion

        #region Revit Model Data Extraction
        /// <summary>
        /// Cleans a cable reference string based on specific rules.
        /// If a '/' character is present, it removes the last part of the string before the slash, including the preceding '-'.
        /// For example, "CM12A-27-01-MSSB/L00/0003A" becomes "CM12A-27-01".
        /// </summary>
        /// <param name="cableReference">The original cable reference string.</param>
        /// <returns>The cleaned cable reference string.</returns>
        private string CleanCableReference(string cableReference)
        {
            if (string.IsNullOrEmpty(cableReference))
            {
                return cableReference;
            }

            // Trim whitespace from the beginning and end of the string initially
            string trimmedReference = cableReference.Trim();

            int slashIndex = trimmedReference.IndexOf('/');
            if (slashIndex == -1)
            {
                return trimmedReference; // No slash, return the already trimmed string.
            }

            // Get the part of the string before the slash
            string preSlashPart = trimmedReference.Substring(0, slashIndex);

            // Find the last hyphen in that part
            int lastHyphenIndex = preSlashPart.LastIndexOf('-');

            if (lastHyphenIndex == -1)
            {
                // If there's no hyphen, return the whole part before the slash, trimmed again for safety
                return preSlashPart.Trim();
            }

            // Return the part of the string before the last hyphen, trimmed again for safety
            return preSlashPart.Substring(0, lastHyphenIndex).Trim();
        }

        /// <summary>
        /// Collects data from Cable Tray elements in the Revit document.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="cableScheduleData">A list of processed cable data for lookups.</param>
        /// <returns>A list of TrayCableData objects.</returns>
        private List<TrayCableData> CollectTrayData(Document doc, List<CableData> cableScheduleData)
        {
            // Define the GUIDs for the shared parameters
            var rtsIdGuid = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
            var cableGuids = GetCableGuids();

            // Step 1: Group all cable references by RTS_ID from all trays
            var groupedCables = new Dictionary<string, HashSet<string>>();
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_CableTray)
                .WhereElementIsNotElementType();

            foreach (Element tray in collector)
            {
                Parameter rtsIdParam = tray.get_Parameter(rtsIdGuid);
                if (rtsIdParam != null && rtsIdParam.HasValue)
                {
                    string rtsId = rtsIdParam.AsString();
                    if (string.IsNullOrEmpty(rtsId)) continue;

                    if (!groupedCables.ContainsKey(rtsId))
                    {
                        groupedCables[rtsId] = new HashSet<string>();
                    }

                    foreach (var guid in cableGuids)
                    {
                        Parameter cableParam = tray.get_Parameter(guid);
                        if (cableParam != null && cableParam.HasValue)
                        {
                            string originalValue = cableParam.AsString();
                            string processedValue = CleanCableReference(originalValue);
                            if (!string.IsNullOrEmpty(processedValue))
                            {
                                groupedCables[rtsId].Add(processedValue);
                            }
                        }
                    }
                }
            }

            // Step 2: Process the grouped data to create the final list for export
            var trayDataList = new List<TrayCableData>();
            var cableDataDict = cableScheduleData
                .Where(c => !string.IsNullOrEmpty(c.CableReference))
                .ToDictionary(c => c.CableReference, c => c);

            var standardTraySizes = new List<int> { 100, 150, 300, 450, 600, 900, 1000 };

            foreach (var entry in groupedCables)
            {
                string rtsId = entry.Key;
                List<string> sortedCables = entry.Value.OrderBy(c => c).ToList();

                var trayRecord = new TrayCableData
                {
                    RtsId = rtsId,
                    CableValues = new List<string>()
                };

                double totalWeight = 0.0;
                double totalOccupancy = 0.0;
                var unmatchedCables = new List<string>();

                // Calculate weight and find unmatched cables based on the unique, sorted list
                foreach (string cableRef in sortedCables)
                {
                    if (cableDataDict.ContainsKey(cableRef))
                    {
                        var cableInfo = cableDataDict[cableRef];

                        // Calculate weight
                        if (double.TryParse(cableInfo.CablesKgPerM, out double kgPerM))
                        {
                            totalWeight += kgPerM;
                        }

                        // Calculate occupancy
                        double diameter = 0.0;
                        int numActiveCables = 0;
                        double deratingFactor = 0.0;

                        bool diameterValid = double.TryParse(cableInfo.NominalOverallDiameter, out diameter);
                        bool numCablesValid = int.TryParse(cableInfo.NumberOfActiveCables, out numActiveCables);
                        bool deratingValid = double.TryParse(cableInfo.AsNsz3008CableDeratingFactor, out deratingFactor);

                        if (diameterValid && numCablesValid)
                        {
                            double occupancyValue = 0.0;
                            bool isSdi = cableInfo.Cores != null &&
                                        (cableInfo.Cores.IndexOf("S.D.I.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         cableInfo.Cores.IndexOf("SDI", StringComparison.OrdinalIgnoreCase) >= 0);

                            // Apply new occupancy formula based on derating factor
                            if (deratingFactor == 1.0 || deratingFactor == 0.9384 || deratingFactor == 0.8968 || deratingFactor == 0.81)
                            {
                                occupancyValue = diameter * (isSdi ? numActiveCables * 3 : numActiveCables * 2);
                            }
                            else
                            {
                                occupancyValue = diameter * (isSdi ? numActiveCables * 2 : numActiveCables * 1);
                            }
                            totalOccupancy += occupancyValue;
                        }
                    }
                    else
                    {
                        unmatchedCables.Add(cableRef);
                    }
                }

                // Populate CableValues with the sorted list, padding with empty strings to fill 30 columns
                for (int i = 0; i < 30; i++) // Changed from 15 to 30
                {
                    if (i < sortedCables.Count)
                    {
                        trayRecord.CableValues.Add(sortedCables[i]);
                    }
                    else
                    {
                        trayRecord.CableValues.Add(string.Empty);
                    }
                }

                if (unmatchedCables.Any())
                {
                    trayRecord.RtsComment = "Cables Error: " + string.Join(":", unmatchedCables);
                }
                else
                {
                    trayRecord.RtsComment = string.Empty;
                }

                // Determine minimum tray size
                int minTraySize = 0;
                foreach (int size in standardTraySizes)
                {
                    if (totalOccupancy <= size)
                    {
                        minTraySize = size;
                        break;
                    }
                }

                trayRecord.CablesWeight = totalWeight.ToString("F1");
                trayRecord.TrayOccupancy = Math.Round(totalOccupancy, 1).ToString("F1");
                trayRecord.TrayMinSize = minTraySize.ToString();

                trayDataList.Add(trayRecord);
            }

            return trayDataList;
        }

        private List<CableLengthData> CollectCableLengthsData(Document doc, List<CableData> cableScheduleData)
        {
            // Dictionary to hold cable reference and its associated run data (lengths and tray types)
            var cableRunData = new Dictionary<string, CableRunInfo>();
            var cableGuids = GetCableGuids();

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting
            };
            var categoryFilter = new ElementMulticategoryFilter(categories);

            var collector = new FilteredElementCollector(doc)
                .WherePasses(categoryFilter)
                .WhereElementIsNotElementType();

            foreach (Element element in collector)
            {
                // Get the length of the cable tray or fitting in meters
                double lengthInFeet = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0;
                double lengthInMeters = lengthInFeet * 0.3048;

                if (lengthInMeters == 0) continue;

                // Get the Type Name of the tray
                string trayTypeName = doc.GetElement(element.GetTypeId())?.Name ?? "Unknown Type";

                foreach (var guid in cableGuids)
                {
                    Parameter cableParam = element.get_Parameter(guid);
                    if (cableParam != null && cableParam.HasValue)
                    {
                        string originalValue = cableParam.AsString();
                        string processedValue = CleanCableReference(originalValue);
                        if (!string.IsNullOrEmpty(processedValue))
                        {
                            if (!cableRunData.ContainsKey(processedValue))
                            {
                                cableRunData[processedValue] = new CableRunInfo();
                            }
                            cableRunData[processedValue].Lengths.Add(lengthInMeters);

                            // Only add the type name if it's a Cable Tray, not a fitting
                            if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                            {
                                cableRunData[processedValue].TrayTypeNames.Add(trayTypeName);
                            }
                        }
                    }
                }
            }

            var cableDataDict = cableScheduleData
                .Where(c => !string.IsNullOrEmpty(c.CableReference))
                .ToDictionary(c => c.CableReference, c => c);

            var exportData = new List<CableLengthData>();
            foreach (var kvp in cableRunData)
            {
                string cableRef = kvp.Key;
                double totalLength = kvp.Value.Lengths.Sum();
                double finalLength = Math.Round(totalLength, 1) + 5.0; // Add 5m buffer

                // Ensure RTS_Comment is initialized and only contains distinct non-empty tray type names, joined by ":"
                string rtsComment = string.Join(":", kvp.Value.TrayTypeNames.Where(name => !string.IsNullOrEmpty(name)).Distinct().OrderBy(name => name));


                string fromValue = "Missing from PowerCAD";
                string toValue = "Missing from PowerCAD";

                if (cableDataDict.ContainsKey(cableRef))
                {
                    var cableInfo = cableDataDict[cableRef];
                    fromValue = cableInfo.From;
                    toValue = cableInfo.To;
                }

                exportData.Add(new CableLengthData
                {
                    PC_Cable_Reference = cableRef,
                    From = fromValue,
                    To = toValue,
                    PC_Cable_Length = finalLength.ToString("F1"),
                    RTS_Comment = rtsComment
                });
            }

            return exportData.OrderBy(c => c.PC_Cable_Reference).ToList();
        }

        private void UpdateRevitParameters(Document doc, List<TrayCableData> trayDataList)
        {
            // Define the GUIDs for the parameters to be updated
            var rtsIdGuid = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
            var rtTrayOccupancyGuid = new Guid("a6f087c7-cecc-4335-831b-249cb9398abc");
            var rtCablesWeightGuid = new Guid("51d670fa-0338-42e7-ac9e-f2c44a56ffcc");
            var rtTrayMinSizeGuid = new Guid("5ed6b64c-af5c-4425-ab69-85a7fa5fdffe");

            using (Transaction tx = new Transaction(doc, "Update Cable Tray Parameters"))
            {
                tx.Start();

                // Create a dictionary of the calculated data for quick lookup
                var trayDataDict = trayDataList.ToDictionary(t => t.RtsId);

                var collector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CableTray)
                    .WhereElementIsNotElementType();

                foreach (Element tray in collector)
                {
                    Parameter rtsIdParam = tray.get_Parameter(rtsIdGuid);
                    if (rtsIdParam != null && rtsIdParam.HasValue)
                    {
                        string rtsId = rtsIdParam.AsString();
                        if (trayDataDict.ContainsKey(rtsId))
                        {
                            var data = trayDataDict[rtsId];

                            // Update RT_Tray Occupancy
                            Parameter occupancyParam = tray.get_Parameter(rtTrayOccupancyGuid);
                            if (occupancyParam != null && !occupancyParam.IsReadOnly)
                            {
                                occupancyParam.Set(data.TrayOccupancy);
                            }

                            // Update RT_Cables Weight
                            Parameter weightParam = tray.get_Parameter(rtCablesWeightGuid);
                            if (weightParam != null && !weightParam.IsReadOnly)
                            {
                                weightParam.Set(data.CablesWeight);
                            }

                            // Update RT_Tray Min Size
                            Parameter minSizeParam = tray.get_Parameter(rtTrayMinSizeGuid);
                            if (minSizeParam != null && !minSizeParam.IsReadOnly)
                            {
                                minSizeParam.Set(data.TrayMinSize);
                            }
                        }
                    }
                }

                tx.Commit();
            }
        }


        /// <summary>
        /// Exports the collected Cable Tray data to a CSV file.
        /// </summary>
        private void ExportTrayDataToCsv(List<TrayCableData> data, string filePath)
        {
            var sb = new StringBuilder();

            var headers = new List<string> { "RTS_ID" };
            for (int i = 1; i <= 30; i++) // Changed from 15 to 30
            {
                headers.Add($"Cable_{i:D2}");
            }
            // Add new headers
            headers.Add("RT_Tray Occupancy");
            headers.Add("RT_Cables Weight");
            headers.Add("RT_Tray Min Size");
            headers.Add("RTS_Comment");

            sb.AppendLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var line = new List<string> { row.RtsId };
                line.AddRange(row.CableValues);
                // Add new values to the line
                line.Add(row.TrayOccupancy);
                line.Add(row.CablesWeight);
                line.Add(row.TrayMinSize);
                line.Add(row.RtsComment);

                var formattedLine = line.Select(val =>
                {
                    if (val == null) return string.Empty;
                    val = val.Trim(); // Trim whitespace
                    if (val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r"))
                    {
                        // Escape double quotes by doubling them, then enclose the whole field in quotes
                        return $"\"{val.Replace("\"", "\"\"")}\"";
                    }
                    return val;
                });
                sb.AppendLine(string.Join(",", formattedLine));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private void ExportCableLengthsToCsv(List<CableLengthData> data, string filePath)
        {
            var sb = new StringBuilder();

            // Define headers
            var headers = new List<string> { "PC_Cable Reference", "From", "To", "PC_Cable Length", "RTS_Comment" };
            sb.AppendLine(string.Join(",", headers));

            // Append data rows
            foreach (var row in data)
            {
                var line = new List<string> { row.PC_Cable_Reference, row.From, row.To, row.PC_Cable_Length, row.RTS_Comment };
                var formattedLine = line.Select(val =>
                {
                    if (val == null) return string.Empty;
                    val = val.Trim(); // Trim whitespace
                    if (val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r"))
                    {
                        // Escape double quotes by doubling them, then enclose the whole field in quotes
                        return $"\"{val.Replace("\"", "\"\"")}\"";
                    }
                    return val;
                });
                sb.AppendLine(string.Join(",", formattedLine));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private List<Guid> GetCableGuids()
        {
            return new List<Guid>
            {
                new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), // Cable_01
                new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), // Cable_02
                new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"), // Cable_03
                new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), // Cable_04
                new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"), // Cable_05
                new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), // Cable_06
                new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"), // Cable_07
                new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), // Cable_08
                new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"), // Cable_09
                new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), // Cable_10
                new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"), // Cable_11
                new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), // Cable_12
                new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"), // Cable_13
                new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), // Cable_14
                new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"), // Cable_15
                new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), // Cable_16
                new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"), // Cable_17
                new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), // Cable_18
                new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"), // Cable_19
                new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), // Cable_20
                new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"), // Cable_21
                new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), // Cable_22
                new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"), // Cable_23
                new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), // Cable_24
                new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"), // Cable_25
                new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), // Cable_26
                new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"), // Cable_27
                new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), // Cable_28
                new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"), // Cable_29
                new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b")  // Cable_30
            };
        }
        #endregion

        #region Data Classes
        /// <summary>
        /// A data class to hold the values for a single row of processed cable data from the CSV.
        /// Public for JSON serialization/deserialization.
        /// </summary>
        public class CableData // Changed from private to public
        {
            public string CableReference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string CableType { get; set; }
            public string CableCode { get; set; }
            public string CableConfiguration { get; set; }
            public string Cores { get; set; }
            public string ConductorActive { get; set; }
            public string Insulation { get; set; }
            public string ConductorEarth { get; set; }
            public string SeparateEarthForMulticore { get; set; }
            public string CableLength { get; set; }
            public string TotalCableRunWeight { get; set; }
            public string NominalOverallDiameter { get; set; }
            public string NumberOfActiveCables { get; set; }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; }
            public string AsNsz3008CableDeratingFactor { get; set; }
        }

        /// <summary>
        /// A data class to hold data collected from a Cable Tray element in Revit.
        /// </summary>
        private class TrayCableData
        {
            public string RtsId { get; set; }
            public List<string> CableValues { get; set; }

            // New calculated/placeholder properties
            public string TrayOccupancy { get; set; }
            public string CablesWeight { get; set; }
            public string TrayMinSize { get; set; }
            public string RtsComment { get; set; }
        }

        private class CableLengthData
        {
            public string PC_Cable_Reference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string PC_Cable_Length { get; set; }
            public string RTS_Comment { get; set; }
        }

        /// <summary>
        /// A helper class to store information about a cable's run through trays.
        /// </summary>
        private class CableRunInfo
        {
            public List<double> Lengths { get; set; }
            public HashSet<string> TrayTypeNames { get; set; }

            public CableRunInfo()
            {
                Lengths = new List<double>();
                TrayTypeNames = new HashSet<string>();
            }
        }
        #endregion
    }
}