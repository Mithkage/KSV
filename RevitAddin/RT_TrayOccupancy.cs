// File: RT_TrayOccupancy.cs

//
// Namespace: RT_TrayOccupancy
//
// Class: RT_TrayOccupancyClass
//
// Function: This Revit external command now reads cable data directly from
//           extensible storage (managed by PC_Extensible). It then scans the active Revit
//           model for Cable Trays with specific parameters, calculates total cable weight
//           and occupancy for each tray using conditional logic, reports on missing cables,
//           determines minimum tray size, and exports that data to a second CSV. Finally,
//           it calculates the total run length for each unique cable across all trays and
//           fittings, exports that to a third CSV, and updates/merges this calculated data
//           into the Model Generated Data extensible storage based on Cable Reference.
//
// Author: Kyle Vorster
//
// Date: June 13, 2024 (Updated July 1, 2025 - Reads from Extensible Storage, Fixed Dictionary Key)
//       July 2, 2025 - Improved CleanCableReference logic for better parsing.
//       July 2, 2025 - Modified UpdateRevitParameters to update Cable_XX parameters on Conduits,
//                      Conduit Fittings, and Cable Tray Fittings.
//       July 2, 2025 - Resolved CS1061 error by updating method call to generic RecallDataFromExtensibleStorage<T>.
//       July 2, 2025 - Updated ExportDataToCsv to include FileName and ImportDate.
//       July 2, 2025 - Added logic to save calculated Cable Lengths data to Model Generated Data extensible storage.
//       July 2, 2025 - Excluded Variance from data written to ModelGeneratedData storage.
//       July 2, 2025 - Fixed "An item with the same key has already been added" error in CollectCableLengthsData
//                      by handling duplicate CableReference keys when creating lookup dictionary.
//       July 2, 2025 - Implemented update/merge logic for saving data to Model Generated Data extensible storage.
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions; // Added for generic prefix checking
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible; // Added reference to the PC_Extensible namespace
using System.Reflection; // For property iteration in merge logic
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
        /// <summary>
        /// The main entry point for the external command. Revit calls this method when the user clicks the button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- 1. RECALL CABLE DATA FROM EXTENSIBLE STORAGE (managed by PC_Extensible) ---
                // We instantiate PC_ExtensibleClass to access its public RecallDataFromExtensibleStorage method.
                PC_ExtensibleClass pcExtensible = new PC_ExtensibleClass();
                List<PC_ExtensibleClass.CableData> cleanedData = pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    doc,
                    PC_ExtensibleClass.PrimarySchemaGuid,         // Pass the primary schema GUID
                    PC_ExtensibleClass.PrimarySchemaName,         // Pass the primary schema name
                    PC_ExtensibleClass.PrimaryFieldName,          // Pass the primary field name
                    PC_ExtensibleClass.PrimaryDataStorageElementName // Pass the primary data storage element name
                );

                if (cleanedData == null || cleanedData.Count == 0)
                {
                    TaskDialog.Show("No Data Found", "No valid cable data was found in the project's primary extensible storage. Please run the 'PC_Extensible' command to import data first. The process will now exit.");
                    return Result.Succeeded; // Consider this a successful exit, as there's no data to process.
                }

                // --- 2. PROMPT USER FOR OUTPUT FOLDER ---
                string outputFolderPath = GetOutputFolderPath();
                if (string.IsNullOrEmpty(outputFolderPath))
                {
                    return Result.Cancelled;
                }

                // --- 3. EXPORT THE RECALLED DATA TO A NEW CSV (Optional, for verification/record keeping) ---
                // This step is kept to provide a 'Cleaned_Cable_Schedule.csv' output,
                // consistent with the previous version's behavior, but now from stored data.
                string cleanedScheduleFilePath = Path.Combine(outputFolderPath, "Cleaned_Cable_Schedule_From_Storage.csv");
                ExportDataToCsv(cleanedData, cleanedScheduleFilePath);


                // --- 4. PROCESS AND EXPORT CABLE TRAY DATA FROM REVIT MODEL ---
                List<TrayCableData> trayData = CollectTrayData(doc, cleanedData);
                if (trayData.Any())
                {
                    string trayDataFilePath = Path.Combine(outputFolderPath, "CableTray_Data.csv");
                    ExportTrayDataToCsv(trayData, trayDataFilePath);

                    // --- 5. UPDATE REVIT MODEL PARAMETERS ---
                    // Now passing the cable schedule data to UpdateRevitParameters
                    UpdateRevitParameters(doc, trayData);
                }

                // --- 6. PROCESS AND EXPORT CABLE LENGTH DATA ---
                List<CableLengthData> cableLengths = CollectCableLengthsData(doc, cleanedData);
                if (cableLengths.Any())
                {
                    string cableLengthsFilePath = Path.Combine(outputFolderPath, "Cable_Lengths.csv");
                    ExportCableLengthsToCsv(cableLengths, cableLengthsFilePath);

                    // --- 7. SAVE/MERGE CALCULATED CABLE LENGTHS DATA TO MODEL GENERATED EXTENSIBLE STORAGE ---
                    // Recall existing Model Generated Data first for merging
                    List<PC_ExtensibleClass.ModelGeneratedData> existingModelGeneratedData = pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                        doc,
                        PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                        PC_ExtensibleClass.ModelGeneratedSchemaName,
                        PC_ExtensibleClass.ModelGeneratedFieldName,
                        PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                    );

                    // Create a dictionary from existing data for efficient lookups by CableReference
                    var mergedModelDataDict = existingModelGeneratedData
                        .Where(mgd => !string.IsNullOrEmpty(mgd.CableReference))
                        .GroupBy(mgd => mgd.CableReference) // Handle potential duplicates in existing data
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    int updatedEntriesCount = 0;
                    int addedEntriesCount = 0;

                    // Iterate through newly calculated cable lengths
                    foreach (var newCalculatedCable in cableLengths)
                    {
                        if (string.IsNullOrEmpty(newCalculatedCable.PC_Cable_Reference))
                        {
                            // Skip if Cable Reference is empty in new data
                            continue;
                        }

                        // Create a new ModelGeneratedData object from the calculated data
                        var newModelDataEntry = new PC_ExtensibleClass.ModelGeneratedData
                        {
                            CableReference = newCalculatedCable.PC_Cable_Reference,
                            From = newCalculatedCable.From,
                            To = newCalculatedCable.To,
                            CableLengthM = newCalculatedCable.PC_Cable_Length,
                            Variance = string.Empty, // As per previous instruction, variance is empty
                            Comment = newCalculatedCable.RTS_Comment
                        };

                        if (mergedModelDataDict.TryGetValue(newModelDataEntry.CableReference, out PC_ExtensibleClass.ModelGeneratedData existingEntry))
                        {
                            // Match found: Update existing entry's values
                            // Iterate through properties and update if new value is not null/blank
                            PropertyInfo[] properties = typeof(PC_ExtensibleClass.ModelGeneratedData).GetProperties();
                            bool entryChanged = false;

                            foreach (PropertyInfo prop in properties)
                            {
                                if (prop.PropertyType == typeof(string) && prop.CanWrite)
                                {
                                    string newValue = (string)prop.GetValue(newModelDataEntry);
                                    string currentValue = (string)prop.GetValue(existingEntry);

                                    // Only update if the new value is not null/empty/whitespace AND it's different from the current value
                                    if (!string.IsNullOrWhiteSpace(newValue) && !string.Equals(newValue, currentValue, StringComparison.Ordinal))
                                    {
                                        prop.SetValue(existingEntry, newValue);
                                        entryChanged = true;
                                    }
                                    // If newValue is null/empty/whitespace, we retain the existing value.
                                    // If it's the same, no change needed.
                                }
                            }
                            if (entryChanged) updatedEntriesCount++;
                        }
                        else
                        {
                            // No match found: Add as a new entry
                            mergedModelDataDict.Add(newModelDataEntry.CableReference, newModelDataEntry);
                            addedEntriesCount++;
                        }
                    }

                    List<PC_ExtensibleClass.ModelGeneratedData> finalModelGeneratedData = mergedModelDataDict.Values.ToList();

                    // Save the final merged data to Model Generated Data storage
                    using (Transaction tx = new Transaction(doc, "Save Model Generated Data"))
                    {
                        tx.Start();
                        try
                        {
                            pcExtensible.SaveDataToExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                                doc,
                                finalModelGeneratedData,
                                PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                                PC_ExtensibleClass.ModelGeneratedSchemaName,
                                PC_ExtensibleClass.ModelGeneratedFieldName,
                                PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                            );
                            tx.Commit();
                            TaskDialog.Show("Model Data Saved", $"Calculated Cable Lengths data successfully saved/merged into Model Generated Data storage.\nUpdated entries: {updatedEntriesCount}\nAdded entries: {addedEntriesCount}");
                        }
                        catch (Exception ex)
                        {
                            tx.RollBack();
                            TaskDialog.Show("Model Data Save Error", $"Failed to save Model Generated Data: {ex.Message}");
                        }
                    }
                }

                // --- 8. NOTIFY USER OF OVERALL SUCCESS ---
                TaskDialog.Show("Process Complete", $"Process complete. Revit model has been updated and files have been saved to:\n{outputFolderPath}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An unexpected error occurred: {ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        #region File Dialog Methods
        // GetSourceCsvFilePath() is removed as data is now read from extensible storage.

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

        #region CSV Parsing and Processing (Modified for Export of recalled data)
        // ParseCsvLine and ParseAndProcessCsvData are removed as data is read from extensible storage.
        // ProcessSplit and GetValueOrDefault are also removed as they are no longer needed here.

        // Note: This ExportDataToCsv now takes PC_ExtensibleClass.CableData
        private void ExportDataToCsv(List<PC_ExtensibleClass.CableData> data, string filePath)
        {
            var sb = new StringBuilder();

            // Headers must match the properties of PC_ExtensibleClass.CableData, now including FileName and ImportDate
            var headers = new List<string>
            {
                "File Name", "Import Date", // New headers at the beginning
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
                    row.FileName,       // New data field
                    row.ImportDate,     // New data field
                    row.CableReference, row.From, row.To, row.CableType, row.CableCode,
                    row.CableConfiguration, row.Cores, row.NumberOfActiveCables, row.ActiveCableSize,
                    row.ConductorActive, row.Insulation, row.NumberOfNeutralCables, row.NeutralCableSize,
                    row.NumberOfEarthCables, row.EarthCableSize, row.ConductorEarth,
                    row.SeparateEarthForMulticore, row.CableLength, row.TotalCableRunWeight,
                    row.CablesKgPerM, row.NominalOverallDiameter, row.AsNsz3008CableDeratingFactor
                };

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
        /// Cleans a cable reference string based on specific rules, aiming to extract the primary identifier.
        /// It handles various formats, including those with slashes, parenthetical expressions,
        /// and specific hyphenated patterns like "CMXXA-YY-ZZZ".
        /// Examples:
        /// - "CM12A-27-01-MSSB/L00/0003A" becomes "CM12A-27-01"
        /// - "CM32A-17-ATL-SN350/UPIO/B01/3203A (INPUT BUS)" becomes "CM32A-17"
        /// - "PC-123/ABC" becomes "PC-123"
        /// - "SIMPLE-CABLE" remains "SIMPLE-CABLE"
        /// - "CABLE (NOTES)" becomes "CABLE"
        /// - "CM12A-27" remains "CM12A-27"
        /// </summary>
        /// <param name="cableReference">The original cable reference string.</param>
        /// <returns>The cleaned cable reference string.</returns>
        private string CleanCableReference(string cableReference)
        {
            if (string.IsNullOrEmpty(cableReference))
            {
                return cableReference;
            }

            string cleaned = cableReference.Trim();

            // 1. Remove any text within parentheses and the parentheses themselves.
            // Example: "CM32A-17-ATL-SN350 (INPUT BUS)" -> "CM32A-17-ATL-SN350"
            int openParenIndex = cleaned.IndexOf('(');
            if (openParenIndex != -1)
            {
                // Ensure there's a closing parenthesis to avoid cutting off valid parts prematurely
                int closeParenIndex = cleaned.IndexOf(')', openParenIndex);
                if (closeParenIndex != -1)
                {
                    cleaned = cleaned.Substring(0, openParenIndex).Trim();
                }
                else
                {
                    // If only an opening parenthesis exists, just trim from there
                    cleaned = cleaned.Substring(0, openParenIndex).Trim();
                }
            }

            // 2. Remove everything from the first '/' onwards.
            // Example: "CM32A-17-ATL-SN350/UPIO/B01/3203A" -> "CM32A-17-ATL-SN350"
            // Example: "CM12A-27-01-MSSB/L00/0003A" -> "CM12A-27-01-MSSB"
            int firstSlashIndex = cleaned.IndexOf('/');
            if (firstSlashIndex != -1)
            {
                cleaned = cleaned.Substring(0, firstSlashIndex).Trim();
            }

            // 3. Handle specific hyphenated patterns after initial cleaning.
            // This is crucial for distinguishing between "CM12A-27-01" and "CM32A-17-ATL".
            string[] parts = cleaned.Split('-');

            // Define a regex pattern for prefixes: two alphabetic characters followed by two numeric characters
            Regex prefixPattern = new Regex(@"^[A-Za-z]{2}\d{2}", RegexOptions.IgnoreCase);

            // Rule A: For patterns like "XXYYA-ZZ-AAA-BBB" where AAA is non-numeric, truncate to "XXYYA-ZZ".
            // This handles "CM32A-17-ATL-SN350" -> "CM32A-17" because "ATL" is parts[2] and non-numeric.
            if (parts.Length >= 3 && prefixPattern.IsMatch(parts[0]))
            {
                if (!int.TryParse(parts[2], out _)) // If the third part is NOT purely numeric
                {
                    return $"{parts[0]}-{parts[1]}";
                }
            }

            // Rule B: For patterns like "XXYYA-ZZ-01-AAA" where 01 is numeric (parts[2]) but AAA (parts[3]) is non-numeric,
            // truncate to "XXYYA-ZZ-01".
            // This specifically targets "CM42B-26-01-ATL-SN350" -> "CM42B-26-01".
            // This rule should come *after* Rule A, as Rule A is more aggressive.
            if (parts.Length >= 4 && prefixPattern.IsMatch(parts[0]))
            {
                // Check if the third part *is* numeric (like "01") AND the fourth part is *not* numeric (like "ATL").
                if (int.TryParse(parts[2], out _) && !int.TryParse(parts[3], out _))
                {
                    return $"{parts[0]}-{parts[1]}-{parts[2]}";
                }
            }


            // If none of the specific rules above applied (e.g., "SIMPLE-CABLE", or "CM12A-27"),
            // return the string as is after the initial trimming and slash/parentheses removal.
            return cleaned;
        }

        /// <summary>
        /// Collects data from Cable Tray elements in the Revit document.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="cableScheduleData">A list of processed cable data for lookups.</param>
        /// <returns>A list of TrayCableData objects.</returns>
        private List<TrayCableData> CollectTrayData(Document doc, List<PC_ExtensibleClass.CableData> cableScheduleData)
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
                            string processedValue = CleanCableReference(originalValue); // Use the improved cleaning method
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
            // Use To property as key for lookup since that is how we merge in PC_Extensible
            // This dictionary is for looking up CableData by its 'To' property, as used in PC_Extensible for merging.
            // However, the trays themselves hold 'Cable Reference', so we'll use FirstOrDefault for lookup later.
            var cableDataDict = cableScheduleData
           .Where(c => !string.IsNullOrEmpty(c.To)) // Ensure 'To' is not null/empty for dictionary key
                 .ToDictionary(c => c.To, c => c); // This dictionary is not directly used for lookup of CableReference from trays.

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
                    // For cable data lookup, we now use the CableReference as the key from the stored data,
                    // as the trays themselves hold the Cable Reference.
                    // If 'Cable Reference' is not unique, this logic might need adjustment.
                    var cableInfo = cableScheduleData.FirstOrDefault(c => c.CableReference == cableRef);

                    if (cableInfo != null)
                    {
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

        private List<CableLengthData> CollectCableLengthsData(Document doc, List<PC_ExtensibleClass.CableData> cableScheduleData)
        {
            // Dictionary to hold cable reference and its associated run data (lengths and tray types)
            var cableRunData = new Dictionary<string, CableRunInfo>();
            var cableGuids = GetCableGuids();

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting, // Include Cable Tray Fittings
                BuiltInCategory.OST_Conduit,           // Include Conduits
                BuiltInCategory.OST_ConduitFitting     // Include Conduit Fittings
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

                // Get the Type Name of the element (tray, conduit, or fitting)
                string elementTypeCategory = element.Category.Name; // e.g., "Cable Trays", "Conduits"
                string elementTypeName = doc.GetElement(element.GetTypeId())?.Name ?? "Unknown Type";


                foreach (var guid in cableGuids)
                {
                    Parameter cableParam = element.get_Parameter(guid);
                    if (cableParam != null && cableParam.HasValue)
                    {
                        string originalValue = cableParam.AsString();
                        string processedValue = CleanCableReference(originalValue); // Use the improved cleaning method
                        if (!string.IsNullOrEmpty(processedValue))
                        {
                            if (!cableRunData.ContainsKey(processedValue))
                            {
                                cableRunData[processedValue] = new CableRunInfo();
                            }
                            cableRunData[processedValue].Lengths.Add(lengthInMeters);

                            // Add the category name to identify the type of host element
                            cableRunData[processedValue].HostElementCategories.Add(elementTypeCategory);
                        }
                    }
                }
            }

            // FIX: Change the dictionary key from CableReference to To
            // This assumes 'To' is unique for each relevant CableData entry in the stored data.
            // If 'To' is not unique, you might need a composite key or a List<CableData> as value.
            // Create a dictionary from existing data for efficient lookups by 'CableReference' key
            // MODIFIED: To handle duplicate CableReference keys by taking the first one.
            var cableDataLookup = cableScheduleData
                               .Where(cd => !string.IsNullOrEmpty(cd.CableReference))
                               .GroupBy(cd => cd.CableReference) // Group by the key
                               .ToDictionary(g => g.Key, g => g.First()); // Select the first item from each group

            var exportData = new List<CableLengthData>();
            foreach (var kvp in cableRunData)
            {
                string cableRef = kvp.Key;
                double totalLength = kvp.Value.Lengths.Sum();
                double finalLength = Math.Round(totalLength, 1) + 5.0; // Adding 5m as buffer

                string fromValue = "Missing from PowerCAD";
                string toValue = "Missing from PowerCAD";
                string powerCadCableLength = "N/A"; // New: Initialize as N/A
                string variance = "N/A"; // New: Initialize as N/A
                string rtsComment = string.Join(":", kvp.Value.HostElementCategories.Distinct()); // Use distinct category names


                // Lookup using the 'CableReference' property
                var cableInfo = cableDataLookup.TryGetValue(cableRef, out PC_ExtensibleClass.CableData foundCableInfo) ? foundCableInfo : null;

                if (cableInfo != null)
                {
                    fromValue = cableInfo.From;
                    toValue = cableInfo.To;
                    powerCadCableLength = cableInfo.CableLength; // Get the Cable Length from PC_Extensible data

                    // Calculate Variance
                    if (double.TryParse(powerCadCableLength, out double pcLength) && finalLength > 0)
                    {
                        double diff = finalLength - pcLength;
                        double percentageVariance = (diff / finalLength) * 100;
                        variance = percentageVariance.ToString("F2") + "%"; // Format to 2 decimal places
                    }
                }
                else
                {
                    // If no match found in stored data, append to RTS_Comment
                    if (string.IsNullOrEmpty(rtsComment))
                    {
                        rtsComment = $"Cable Reference '{cableRef}' not found in stored data.";
                    }
                    else
                    {
                        rtsComment += $"; Cable Reference '{cableRef}' not found in stored data.";
                    }
                }

                exportData.Add(new CableLengthData
                {
                    PC_Cable_Reference = cableRef,
                    From = fromValue,
                    To = toValue,
                    PC_Cable_Length = finalLength.ToString("F1"),
                    CableLengthFromCsv = powerCadCableLength, // New: The "Cable Length (m)" from the PC_Extensible data
                    Variance = variance, // New: The calculated variance
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
            var cableGuids = GetCableGuids(); // Get the list of Cable_XX GUIDs

            using (Transaction tx = new Transaction(doc, "Update Cable Parameters")) // Changed transaction name
            {
                tx.Start();

                // Create a dictionary of the calculated data for quick lookup by RTS_ID (for Cable Trays only)
                var trayDataDict = trayDataList.ToDictionary(t => t.RtsId);

                // Define categories for elements that can contain Cable_XX parameters
                var categoriesToUpdateCableParams = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_ConduitFitting
                };
                var filterForCableParamUpdate = new ElementMulticategoryFilter(categoriesToUpdateCableParams);

                var elementsToUpdate = new FilteredElementCollector(doc)
                    .WherePasses(filterForCableParamUpdate)
                    .WhereElementIsNotElementType()
                    .ToList(); // Convert to list to avoid multiple enumerations

                foreach (Element element in elementsToUpdate) // Iterate through all relevant elements
                {
                    // Special handling for Cable Trays: Update their specific RT_ parameters
                    if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                    {
                        Parameter rtsIdParam = element.get_Parameter(rtsIdGuid);
                        if (rtsIdParam != null && rtsIdParam.HasValue)
                        {
                            string rtsId = rtsIdParam.AsString();
                            if (trayDataDict.ContainsKey(rtsId))
                            {
                                var data = trayDataDict[rtsId];

                                // Update RT_Tray Occupancy
                                Parameter occupancyParam = element.get_Parameter(rtTrayOccupancyGuid);
                                if (occupancyParam != null && !occupancyParam.IsReadOnly)
                                {
                                    occupancyParam.Set(data.TrayOccupancy);
                                }

                                // Update RT_Cables Weight
                                Parameter weightParam = element.get_Parameter(rtCablesWeightGuid);
                                if (weightParam != null && !weightParam.IsReadOnly)
                                {
                                    weightParam.Set(data.CablesWeight);
                                }

                                // Update RT_Tray Min Size
                                Parameter minSizeParam = element.get_Parameter(rtTrayMinSizeGuid);
                                if (minSizeParam != null && !minSizeParam.IsReadOnly)
                                {
                                    minSizeParam.Set(data.TrayMinSize);
                                }
                            }
                        }
                    }

                    // --- LOGIC FOR UPDATING CABLE_XX PARAMETERS (applies to all selected categories) ---
                    // First, collect all *unique* cable references currently on this element.
                    // This is to avoid writing duplicates and clear out old ones.
                    HashSet<string> currentUniqueCablesOnElement = new HashSet<string>();
                    foreach (var guid in cableGuids)
                    {
                        Parameter cableParam = element.get_Parameter(guid);
                        if (cableParam != null && cableParam.HasValue)
                        {
                            string originalValue = cableParam.AsString();
                            string cleanedValue = CleanCableReference(originalValue);
                            if (!string.IsNullOrEmpty(cleanedValue))
                            {
                                currentUniqueCablesOnElement.Add(cleanedValue);
                            }
                        }
                    }

                    // Now, re-write the Cable_XX parameters.
                    // We will take the unique cables found on the element, and fill the parameters sequentially.
                    // Any parameters beyond the count of unique cables will be cleared.
                    List<string> cablesToAssign = currentUniqueCablesOnElement.OrderBy(s => s).ToList(); // Sort for consistency

                    for (int i = 0; i < cableGuids.Count; i++)
                    {
                        Parameter cableParam = element.get_Parameter(cableGuids[i]);
                        if (cableParam != null && !cableParam.IsReadOnly)
                        {
                            if (i < cablesToAssign.Count)
                            {
                                // Assign the cable reference
                                cableParam.Set(cablesToAssign[i]);
                            }
                            else
                            {
                                // Clear any remaining Cable_XX parameters if there are fewer cables than parameters
                                if (!string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    cableParam.Set(string.Empty);
                                }
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

                var formattedLine = line.Select(val => val.Contains(",") ? $"\"{val}\"" : $"{(val ?? string.Empty).Trim()}");
                sb.AppendLine(string.Join(",", formattedLine));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private void ExportCableLengthsToCsv(List<CableLengthData> data, string filePath)
        {
            var sb = new StringBuilder();

            // Define headers, now including "Original Cable Length (m)" and "Variance"
            var headers = new List<string> { "PC_Cable Reference", "From", "To", "PC_Cable Length", "Cable Length (m)", "Variance", "RTS_Comment" };
            sb.AppendLine(string.Join(",", headers));

            // Append data rows
            foreach (var row in data)
            {
                var line = new List<string> {
                    row.PC_Cable_Reference,
                    row.From,
                    row.To,
                    row.PC_Cable_Length,
                    row.CableLengthFromCsv, // New: The "Cable Length (m)" from the PC_Extensible data
                    row.Variance, // New: The calculated variance
                    row.RTS_Comment
                };
                var formattedLine = line.Select(val => val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r") ? $"\"{val.Replace("\"", "\"\"")}\"" : $"{(val ?? string.Empty).Trim()}");
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
        // CableData class is no longer defined here as it's now accessed from PC_ExtensibleClass.

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
            public string CableLengthFromCsv { get; set; } // New: Stores the "Cable Length (m)" from PC_Extensible data
            public string Variance { get; set; } // New: The calculated percentage variance
            public string RTS_Comment { get; set; }
        }

        /// <summary>
        /// A helper class to store information about a cable's run through trays.
        /// </summary>
        private class CableRunInfo
        {
            public List<double> Lengths { get; set; }
            public HashSet<string> HostElementCategories { get; set; } // Renamed from TrayTypeNames to be more generic

            public CableRunInfo()
            {
                Lengths = new List<double>();
                HostElementCategories = new HashSet<string>(); // Use a HashSet to store unique category names
            }
        }
        #endregion
    }
}