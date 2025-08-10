//
// File: PC_Extensible.cs
//
// Namespace: PC_Extensible
//
// Class: PC_ExtensibleClass
//
// Function: This Revit external command provides a WPF interface for options to either import
//           cable data from a selected folder (processing multiple "Cables Summary" CSVs)
//           or import consultant cable data from a folder (processing multiple CSVs)
//           or clear existing cable data from either or both extensible storage locations.
//           A third extensible storage is defined for 'Model Generated Data' for use by other scripts.
//           When importing, if updated data is null or blank, existing stored data is retained.
//           The script now remembers the last used folder path within the same Revit session.
//
// Author: Kyle Vorster 
// Company: ReTick Solutions (RTS)
//
// Log:
// - August 16, 2025: Refactored to use the centralized ExtensibleStorageManager utility
//                    while maintaining backward compatibility for existing commands.
// - July 8, 2025: Corrected the ClearData method to properly delete the DataStorage element, which is the supported method in the Revit API.
// - July 8, 2025: Replaced TaskDialog-based UI with a modern WPF window for a better user experience.
// - July 8, 2025: Implemented session memory for the last browsed folder path.
// - July 8, 2025: Added "Load (A)" to the CSV import process and data model.
// - July 7, 2025: Corrected VendorId to match the .addin manifest file and resolve runtime error.
// - July 10, 2025: Fixed TaskDialog usage errors in `HandleClearData` by using static `TaskDialog.Show()` method.
// - July 10, 2025: Refactored `HandleClearData` to use `TaskDialog` instance for multiple command links, resolving CS1501.
// - August 20, 2025: Made CableData and ModelGeneratedData classes public for external reference.
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible;
using RTS.Models; // Added to reference the centralized data models
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Interop;
#endregion

namespace RTS.Commands.DataExchange.DataManagement
{
    /// <summary>
    /// The main class for the Revit external command.
    /// Implements the IExternalCommand interface.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_ExtensibleClass : IExternalCommand
    {
        // Expose the GUIDs and names for backward compatibility
        // These are now referenced from ExtensibleStorageManager for consistency
        public static readonly Guid PrimarySchemaGuid = ExtensibleStorageManager.PrimarySchemaGuid;
        public const string PrimarySchemaName = ExtensibleStorageManager.PrimarySchemaName;
        public const string PrimaryFieldName = ExtensibleStorageManager.PrimaryFieldName;
        public const string PrimaryDataStorageElementName = ExtensibleStorageManager.PrimaryDataStorageElementName;

        public static readonly Guid ConsultantSchemaGuid = ExtensibleStorageManager.ConsultantSchemaGuid;
        public const string ConsultantSchemaName = ExtensibleStorageManager.ConsultantSchemaName;
        public const string ConsultantFieldName = ExtensibleStorageManager.ConsultantFieldName;
        public const string ConsultantDataStorageElementName = ExtensibleStorageManager.ConsultantDataStorageElementName;

        public static readonly Guid ModelGeneratedSchemaGuid = ExtensibleStorageManager.ModelGeneratedSchemaGuid;
        public const string ModelGeneratedSchemaName = ExtensibleStorageManager.ModelGeneratedSchemaName;
        public const string ModelGeneratedFieldName = ExtensibleStorageManager.ModelGeneratedFieldName;
        public const string ModelGeneratedDataStorageElementName = ExtensibleStorageManager.ModelGeneratedDataStorageElementName;

        public const string VendorId = "ReTick_Solutions";

        private static string lastBrowsedPath = null;

        /// <summary>
        /// Data model for cable information used across multiple commands
        /// </summary>
        public class CableData
        {
            public string FileName { get; set; }
            public string ImportDate { get; set; }
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
            public string AsNsz3008CableDeratingFactor { get; set; }
            public string NumberOfActiveCables { get; set; }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; }
            public string Sheath { get; set; }
            public string CableMaxLengthM { get; set; }
            public string VoltageVac { get; set; }
            public string LoadA { get; set; }
        }

        /// <summary>
        /// Data model for model-generated cable information
        /// </summary>
        public class ModelGeneratedData
        {
            public string CableReference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string CableLengthM { get; set; }
            public string Variance { get; set; }
            public string Comment { get; set; }
        }

        /// <summary>
        /// The main entry point for the external command. Revit calls this method when the user clicks the button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --- 1. LAUNCH WPF WINDOW ---
            var window = new PC_Extensible_Window();
            new WindowInteropHelper(window).Owner = commandData.Application.MainWindowHandle;

            bool? result = window.ShowDialog();

            if (result != true) // User cancelled or closed the window
            {
                message = "Operation cancelled by user.";
                // Removed TaskDialog.Show here to avoid showing a dialog on cancel.
                return Result.Cancelled;
            }

            // --- 2. PROCESS USER'S CHOICE FROM WINDOW ---
            switch ((UserAction)window.SelectedAction)
            {
                case UserAction.ImportMyData:
                    return ImportAndMergeData(
                        doc,
                        PrimarySchemaGuid,
                        PrimarySchemaName,
                        PrimaryFieldName,
                        PrimaryDataStorageElementName,
                        "My PowerCAD Data",
                        ParseAndProcessCsvData
                    );

                case UserAction.ImportConsultantData:
                    return ImportAndMergeData(
                        doc,
                        ConsultantSchemaGuid,
                        ConsultantSchemaName,
                        ConsultantFieldName,
                        ConsultantDataStorageElementName,
                        "Consultant PowerCAD Data",
                        ParseAndProcessCsvData
                    );

                case UserAction.ClearData:
                    // The "Clear Data" button was clicked, now show the sub-dialog
                    return HandleClearData(doc, ref message);

                default:
                    message = "Operation cancelled by user.";
                    // No TaskDialog.Show here either.
                    return Result.Cancelled;
            }
        }

        /// <summary>
        /// Handles the sub-dialog for clearing data after the user selects "Clear Data" from the main WPF window.
        /// </summary>
        private Result HandleClearData(Document doc, ref string message)
        {
            // Create a TaskDialog instance to configure command links
            TaskDialog clearDialog = new TaskDialog("Clear Data Options");
            clearDialog.MainInstruction = "Which data set would you like to clear?";
            clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Clear My PowerCAD Data Only");
            clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Clear Consultant PowerCAD Data Only");
            clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Clear Model Generated Data Only");
            clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Clear All Data (My, Consultant, and Model Generated)");
            clearDialog.CommonButtons = TaskDialogCommonButtons.Cancel; // Add a Cancel button
            clearDialog.DefaultButton = TaskDialogResult.CommandLink1; // Set a default button

            TaskDialogResult clearResult = clearDialog.Show();

            if (clearResult == TaskDialogResult.CommandLink1)
            {
                return ClearData(doc, PrimarySchemaGuid, PrimaryDataStorageElementName, "My PowerCAD Data");
            }
            else if (clearResult == TaskDialogResult.CommandLink2)
            {
                return ClearData(doc, ConsultantSchemaGuid, ConsultantDataStorageElementName, "Consultant PowerCAD Data");
            }
            else if (clearResult == TaskDialogResult.CommandLink4)
            {
                return ClearData(doc, ModelGeneratedSchemaGuid, ModelGeneratedDataStorageElementName, "Model Generated Data");
            }
            else if (clearResult == TaskDialogResult.CommandLink3)
            {
                // Calling ClearData for each will show a confirmation dialog for each.
                Result primaryClearResult = ClearData(doc, PrimarySchemaGuid, PrimaryDataStorageElementName, "My PowerCAD Data");
                if (primaryClearResult == Result.Cancelled) { message = "Operation cancelled by user."; return Result.Cancelled; }

                Result consultantClearResult = ClearData(doc, ConsultantSchemaGuid, ConsultantDataStorageElementName, "Consultant PowerCAD Data");
                if (consultantClearResult == Result.Cancelled) { message = "Operation cancelled by user."; return Result.Cancelled; }

                Result modelGeneratedClearResult = ClearData(doc, ModelGeneratedSchemaGuid, ModelGeneratedDataStorageElementName, "Model Generated Data");
                if (modelGeneratedClearResult == Result.Cancelled) { message = "Operation cancelled by user."; return Result.Cancelled; }

                if (primaryClearResult == Result.Succeeded && consultantClearResult == Result.Succeeded && modelGeneratedClearResult == Result.Succeeded)
                {
                    TaskDialog.Show("Clear Data Complete", "All selected PowerCAD data stores were processed successfully.");
                    return Result.Succeeded;
                }
                else
                {
                    message = "One or more clear operations failed. Check previous messages for details.";
                    return Result.Failed;
                }
            }
            else
            {
                message = "Clear operation cancelled by user.";
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Handles the import, merge, and save logic for cable data from a selected folder.
        /// Made generic to work with different data types (CableData).
        /// </summary>
        private Result ImportAndMergeData<T>(
            Document doc,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName,
            string dataTypeName,
            Func<string, List<string>, List<T>> csvParser) where T : class, new() // Constraint for T
        {
            // --- 1. PROMPT USER FOR INPUT FOLDER ---
            string sourceFolderPath = GetSourceFolderPath($"Select Folder Containing {dataTypeName} CSV Files");
            if (string.IsNullOrEmpty(sourceFolderPath))
            {
                return Result.Cancelled; // User cancelled folder selection
            }

            try
            {
                // --- 2. DEFINE THE COLUMNS TO EXTRACT ---
                List<string> columnsToRead = new List<string>
                {
                    "Cable Reference", "From", "To", "Cable Type", "Cable Code", "Cable Configuration",
                    "Cores", "Cable Size (mm\u00B2)", "Conductor (Active)", "Insulation", "Neutral Size (mm\u00B2)",
                    "Earth Size (mm\u00B2)", "Conductor (Earth)", "Separate Earth for Multicore", "Cable Length (m)",
                    "Total Cable Run Weight (Incl. N & E) (kg)", "Nominal Overall Diameter (mm)", "AS/NSZ 3008 Cable Derating Factor",
                    "Sheath", "Cable Max. Length (m)", "Voltage (Vac)", "Load (A)"
                };

                // --- 3. AGGREGATE NEW DATA FROM MULTIPLE CSV FILES ---
                int totalCsvFilesFound = 0;
                int processedFilesCount = 0;
                List<string> skippedFiles = new List<string>();
                List<string> errorFiles = new List<string>();

                string[] csvFiles = Directory.GetFiles(sourceFolderPath, "*.csv", SearchOption.TopDirectoryOnly);
                totalCsvFilesFound = csvFiles.Length;

                if (!csvFiles.Any())
                {
                    TaskDialog.Show("No CSVs Found", $"No CSV files were found in the selected folder: {sourceFolderPath}. No data will be imported.");
                    return Result.Succeeded;
                }

                var aggregatedNewDataDict = new Dictionary<string, T>(StringComparer.OrdinalIgnoreCase);

                foreach (string csvFilePath in csvFiles)
                {
                    string fileName = Path.GetFileName(csvFilePath);
                    if (fileName.IndexOf("Cables Summary", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        skippedFiles.Add($"Skipped '{fileName}': Does not contain 'Cables Summary' in filename.");
                        continue;
                    }

                    try
                    {
                        List<T> fileSpecificNewData = csvParser(csvFilePath, columnsToRead);
                        if (fileSpecificNewData != null && fileSpecificNewData.Any())
                        {
                            foreach (var item in fileSpecificNewData)
                            {
                                string cableRef = GetPropertyValue(item, "CableReference");
                                if (!string.IsNullOrEmpty(cableRef))
                                {
                                    if (!aggregatedNewDataDict.ContainsKey(cableRef))
                                    {
                                        aggregatedNewDataDict.Add(cableRef, item);
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"CableRef '{cableRef}' from '{fileName}' was skipped; it already exists from a prior file (first-wins precedence).");
                                    }
                                }
                                else
                                {
                                    skippedFiles.Add($"Skipped record in '{fileName}': Missing 'Cable Reference'.");
                                }
                            }
                            processedFilesCount++;
                        }
                        else
                        {
                            skippedFiles.Add($"Skipped '{fileName}': No valid data found after parsing.");
                        }
                    }
                    catch (Exception ex)
                    {
                        errorFiles.Add($"Error processing '{fileName}': {ex.Message}");
                        Debug.WriteLine($"Error processing file {fileName}: {ex.Message}");
                    }
                }

                List<T> allNewCsvDataFromFolder = aggregatedNewDataDict.Values.ToList();

                if (!allNewCsvDataFromFolder.Any())
                {
                    TaskDialog.Show("No Relevant Data", $"No relevant 'Cables Summary' data found across all CSV files in '{sourceFolderPath}'. No changes will be made to project data.");
                    return Result.Succeeded;
                }

                // --- 4. RECALL EXISTING DATA FROM EXTENSIBLE STORAGE ---
                List<T> existingStoredData = RecallDataFromExtensibleStorage<T>(
                    doc, schemaGuid, schemaName, fieldName, dataStorageElementName);

                // --- 5. MERGE AGGREGATED NEW DATA WITH EXISTING DATA ---
                var mergedDataDict = existingStoredData
                    .Where(item => !string.IsNullOrEmpty(GetPropertyValue(item, "To")))
                    .GroupBy(item => GetPropertyValue(item, "To"))
                    .ToDictionary(g => g.Key, g => g.First());

                var finalMergedCableReferences = new HashSet<string>(
                    existingStoredData.Where(item => !string.IsNullOrEmpty(GetPropertyValue(item, "CableReference"))).Select(item => GetPropertyValue(item, "CableReference"))
                );

                int updatedEntries = 0;
                int addedEntries = 0;
                List<string> duplicateCableReferencesReport = new List<string>();

                foreach (var newEntry in allNewCsvDataFromFolder)
                {
                    string newEntryTo = GetPropertyValue(newEntry, "To");
                    string newEntryCableRef = GetPropertyValue(newEntry, "CableReference");

                    if (string.IsNullOrEmpty(newEntryTo))
                    {
                        Debug.WriteLine($"Skipping new aggregated entry due to empty 'To' value: Cable Ref: {newEntryCableRef ?? "N/A"}");
                        continue;
                    }

                    if (mergedDataDict.TryGetValue(newEntryTo, out T existingEntry))
                    {
                        PropertyInfo[] properties = typeof(T).GetProperties();
                        bool entryUpdated = false;

                        foreach (PropertyInfo prop in properties)
                        {
                            if (typeof(T) == typeof(CableData) && (prop.Name == nameof(CableData.FileName) || prop.Name == nameof(CableData.ImportDate)))
                            {
                                continue;
                            }

                            if (prop.PropertyType == typeof(string) && prop.CanWrite)
                            {
                                string newValue = (string)prop.GetValue(newEntry);
                                string currentValue = (string)prop.GetValue(existingEntry);

                                if (!string.IsNullOrWhiteSpace(newValue) && !string.Equals(newValue, currentValue, StringComparison.Ordinal))
                                {
                                    prop.SetValue(existingEntry, newValue);
                                    entryUpdated = true;
                                }
                            }
                        }
                        if (entryUpdated) updatedEntries++;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(newEntryCableRef) && finalMergedCableReferences.Contains(newEntryCableRef))
                        {
                            duplicateCableReferencesReport.Add($"Cable Reference: '{newEntryCableRef}', New To: '{newEntryTo}'");
                            Debug.WriteLine($"Skipping new aggregated entry due to duplicate 'Cable Reference' in final merge (no 'To' match): {newEntryCableRef}");
                            continue;
                        }
                        else
                        {
                            mergedDataDict.Add(newEntryTo, newEntry);
                            finalMergedCableReferences.Add(newEntryCableRef);
                            addedEntries++;
                        }
                    }
                }

                List<T> finalMergedData = mergedDataDict.Values.ToList();

                // --- 6. SAVE MERGED DATA TO EXTENSIBLE STORAGE ---
                using (Transaction tx = new Transaction(doc, $"Save Merged {dataTypeName} to Extensible Storage"))
                {
                    tx.Start();
                    try
                    {
                        // Use the new ExtensibleStorageManager utility class
                        ExtensibleStorageManager.SaveDataToExtensibleStorage(
                            doc,
                            finalMergedData,
                            schemaGuid,
                            schemaName,
                            fieldName,
                            dataStorageElementName
                        );
                        tx.Commit();
                        string summaryMessage = $"Data merge to extensible storage complete for {dataTypeName}.\n" +
                                                $"Updated entries: {updatedEntries}\nAdded entries: {addedEntries}";

                        if (duplicateCableReferencesReport.Any())
                        {
                            summaryMessage += "\n\nSkipped new entries (duplicate Cable Ref in final merged data, where 'To' didn't match existing data):\n" + string.Join("\n", duplicateCableReferencesReport.Take(10));
                            if (duplicateCableReferencesReport.Count > 10)
                            {
                                summaryMessage += $"\n... and {duplicateCableReferencesReport.Count - 10} more.";
                            }
                        }
                        TaskDialog.Show("Extensible Storage Save", summaryMessage);
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        string msg = $"Failed to save merged {dataTypeName} to extensible storage: {ex.Message}";
                        if (ex.Message.Contains("Writing of Entities of this Schema is not allowed to the current add-in"))
                        {
                            msg += "\n\nDIAGNOSIS: This often indicates a mismatch in the 'VendorId' specified in the SchemaBuilder " +
                                   "and the actual VendorId of your Revit add-in. " +
                                   "\n\nACTION REQUIRED: Ensure 'schemaBuilder.SetVendorId(\"" + VendorId + "\")' in this script " +
                                   "exactly matches the 'VendorId' attribute in your .addin manifest file, or the VendorId used to register your add-in. " +
                                   "This is crucial for write permissions to Extensible Storage.";
                        }
                        TaskDialog.Show("Extensible Storage Error", msg);
                        return Result.Failed;
                    }
                }

                // --- 7. FINAL SUMMARY OF FILE PROCESSING AND DATA MERGE ---
                StringBuilder finalSummary = new StringBuilder();
                finalSummary.AppendLine($"Bulk Import for {dataTypeName} Process Complete!");
                finalSummary.AppendLine($"Selected Folder: {sourceFolderPath}");
                finalSummary.AppendLine($"Total CSV files in folder: {totalCsvFilesFound}");
                finalSummary.AppendLine($"'Cables Summary' CSVs Processed: {processedFilesCount}");
                finalSummary.AppendLine($"CSVs Skipped (by filter or no data): {skippedFiles.Count}");
                if (skippedFiles.Any())
                {
                    finalSummary.AppendLine("Skipped Files/Records Details (first 5):");
                    foreach (var s in skippedFiles.Take(5)) finalSummary.AppendLine($"- {s}");
                    if (skippedFiles.Count > 5) finalSummary.AppendLine("...");
                }
                finalSummary.AppendLine($"CSVs with Errors: {errorFiles.Count}");
                if (errorFiles.Any())
                {
                    finalSummary.AppendLine("Error Files Details (first 5):");
                    foreach (var e in errorFiles.Take(5)) finalSummary.AppendLine($"- {e}");
                    if (errorFiles.Count > 5) finalSummary.AppendLine("...");
                }

                TaskDialog.Show("Bulk Import Summary", finalSummary.ToString());

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                string msg = $"An unexpected error occurred during {dataTypeName} bulk import process: {ex.Message}\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", msg);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Helper method to get a string property value using reflection.
        /// </summary>
        private string GetPropertyValue<T>(T obj, string propertyName) where T : class
        {
            PropertyInfo prop = typeof(T).GetProperty(propertyName);
            return prop?.GetValue(obj) as string;
        }

        /// <summary>
        /// Handles the clearing of all existing data from extensible storage for a specific schema
        /// by deleting the DataStorage element that contains it.
        /// </summary>
        private Result ClearData(Document doc, Guid schemaGuid, string dataStorageElementName, string dataTypeName)
        {
            // Using static TaskDialog.Show method for confirmation
            TaskDialogResult confirmResult = TaskDialog.Show(
                "Confirm Clear Data",
                $"Are you sure you want to clear ALL existing '{dataTypeName}' data from the project?\n\nThis action cannot be undone.",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                TaskDialogResult.No // Default button
            );

            if (confirmResult == TaskDialogResult.Yes)
            {
                using (Transaction tx = new Transaction(doc, $"Clear {dataTypeName} from Extensible Storage"))
                {
                    tx.Start();
                    try
                    {
                        bool clearResult = ExtensibleStorageManager.ClearData(doc, schemaGuid, dataStorageElementName);

                        if (!clearResult)
                        {
                            tx.RollBack(); // No element to delete, so rollback the transaction
                            TaskDialog.Show("Clear Data Info", $"No existing '{dataTypeName}' data was found in the project to clear.");
                            return Result.Succeeded; // Consider it succeeded as there was nothing to clear.
                        }
                        else
                        {
                            tx.Commit();
                            TaskDialog.Show("Clear Data Complete", $"The DataStorage element for '{dataTypeName}' was successfully deleted.");
                            return Result.Succeeded;
                        }
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        string msg = $"An error occurred while clearing '{dataTypeName}': {ex.Message}";
                        TaskDialog.Show("Clear Data Error", msg);
                        return Result.Failed;
                    }
                }
            }
            else
            {
                TaskDialog.Show("Clear Data Cancelled", $"Clear '{dataTypeName}' operation cancelled.");
                return Result.Cancelled;
            }
        }

        #region Public Methods (for backward compatibility with other code)

        /// <summary>
        /// Saves a list of generic data to extensible storage.
        /// This method is maintained for backward compatibility with existing code.
        /// </summary>
        public void SaveDataToExtensibleStorage<T>(
            Document doc,
            List<T> dataList,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName)
        {
            // Delegate to the new centralized utility
            ExtensibleStorageManager.SaveDataToExtensibleStorage(
                doc, dataList, schemaGuid, schemaName, fieldName, dataStorageElementName);
        }

        /// <summary>
        /// Recalls a list of generic data from extensible storage.
        /// This method is maintained for backward compatibility with existing code.
        /// </summary>
        public List<T> RecallDataFromExtensibleStorage<T>(
            Document doc,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName) where T : new()
        {
            // Delegate to the new centralized utility
            return ExtensibleStorageManager.RecallDataFromExtensibleStorage<T>(
                doc, schemaGuid, schemaName, fieldName, dataStorageElementName);
        }

        #endregion

        #region File Dialog Methods
        /// <summary>
        /// Prompts the user to select a source folder using a modern dialog that allows path pasting.
        /// Remembers the last selected folder within the current Revit session.
        /// </summary>
        private string GetSourceFolderPath(string dialogTitle)
        {
            string result = FileDialogs.PromptForFolder(dialogTitle, lastBrowsedPath);
            if (!string.IsNullOrEmpty(result))
            {
                lastBrowsedPath = result; // Remember this path for next time
            }
            return result;
        }
        #endregion

        #region CSV Parsing and Processing (Specific to Data Types)

        // Parser for CableData (Your original data and Consultant data)
        private List<CableData> ParseAndProcessCsvData(string filePath, List<string> requiredHeaders)
        {
            var cableDataList = new List<CableData>();
            var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            string[] headers = null;

            string fileName = Path.GetFileName(filePath);
            string importDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            using (var reader = new StreamReader(filePath, Encoding.Default))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (headers == null && line.Trim().StartsWith("Cable Reference", StringComparison.OrdinalIgnoreCase))
                    {
                        headers = ParseCsvLine(line).Select(h => h.Trim()).ToArray();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            string header = headers[i];
                            string cleanHeader = header.Replace("Â²", "\u00B2");
                            if (requiredHeaders.Contains(cleanHeader, StringComparer.OrdinalIgnoreCase) && !headerMap.ContainsKey(cleanHeader))
                            {
                                headerMap[cleanHeader] = i;
                            }
                        }
                        continue;
                    }

                    if (headers == null || string.IsNullOrWhiteSpace(line) || line.StartsWith("Transformer:", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    List<string> valuesList = ParseCsvLine(line);

                    if (valuesList.Count <= headerMap.Values.Max())
                    {
                        Debug.WriteLine($"Skipping malformed line in CSV: {line}");
                        continue;
                    }

                    string conductorActiveValue = GetValueOrDefault(valuesList, headerMap, "Conductor (Active)");
                    if (string.IsNullOrWhiteSpace(conductorActiveValue))
                    {
                        continue;
                    }

                    var dataRow = new CableData
                    {
                        FileName = fileName,
                        ImportDate = importDate,
                        CableReference = GetValueOrDefault(valuesList, headerMap, "Cable Reference"),
                        From = GetValueOrDefault(valuesList, headerMap, "From"),
                        To = GetValueOrDefault(valuesList, headerMap, "To"),
                        CableType = GetValueOrDefault(valuesList, headerMap, "Cable Type"),
                        CableCode = GetValueOrDefault(valuesList, headerMap, "Cable Code"),
                        CableConfiguration = GetValueOrDefault(valuesList, headerMap, "Cable Configuration"),
                        Cores = GetValueOrDefault(valuesList, headerMap, "Cores"),
                        ConductorActive = conductorActiveValue,
                        Insulation = GetValueOrDefault(valuesList, headerMap, "Insulation"),
                        ConductorEarth = GetValueOrDefault(valuesList, headerMap, "Conductor (Earth)"),
                        SeparateEarthForMulticore = GetValueOrDefault(valuesList, headerMap, "Separate Earth for Multicore"),
                        CableLength = GetValueOrDefault(valuesList, headerMap, "Cable Length (m)"),
                        TotalCableRunWeight = GetValueOrDefault(valuesList, headerMap, "Total Cable Run Weight (Incl. N & E) (kg)"),
                        NominalOverallDiameter = GetValueOrDefault(valuesList, headerMap, "Nominal Overall Diameter (mm)"),
                        AsNsz3008CableDeratingFactor = GetValueOrDefault(valuesList, headerMap, "AS/NSZ 3008 Cable Derating Factor"),
                        Sheath = GetValueOrDefault(valuesList, headerMap, "Sheath"),
                        CableMaxLengthM = GetValueOrDefault(valuesList, headerMap, "Cable Max. Length (m)"),
                        VoltageVac = GetValueOrDefault(valuesList, headerMap, "Voltage (Vac)"),
                        LoadA = GetValueOrDefault(valuesList, headerMap, "Load (A)")
                    };

                    double weight = 0.0;
                    double length = 0.0;
                    bool isWeightValid = double.TryParse(dataRow.TotalCableRunWeight, out weight);
                    bool isLengthValid = double.TryParse(dataRow.CableLength, out length);

                    if (isWeightValid && isLengthValid && length > 0)
                    {
                        double kgPerMeter = Math.Round(weight / length, 1);
                        dataRow.CablesKgPerM = kgPerMeter.ToString("F1");
                    }
                    else
                    {
                        dataRow.CablesKgPerM = "0.0";
                    }

                    string activeCableValue = GetValueOrDefault(valuesList, headerMap, "Cable Size (mm\u00B2)");
                    string numActive, sizeActive;
                    ProcessSplit(activeCableValue, out numActive, out sizeActive);
                    dataRow.NumberOfActiveCables = numActive;
                    dataRow.ActiveCableSize = sizeActive;

                    string neutralCableValue = GetValueOrDefault(valuesList, headerMap, "Neutral Size (mm\u00B2)");
                    string numNeutral, sizeNeutral;
                    ProcessSplit(neutralCableValue, out numNeutral, out sizeNeutral);
                    dataRow.NumberOfNeutralCables = numNeutral;
                    dataRow.NeutralCableSize = sizeNeutral;

                    string earthCableValue = GetValueOrDefault(valuesList, headerMap, "Earth Size (mm\u00B2)");
                    if (earthCableValue.Equals("No Earth", StringComparison.OrdinalIgnoreCase))
                    {
                        earthCableValue = "0";
                    }
                    string numEarth, sizeEarth;
                    ProcessSplit(earthCableValue, out numEarth, out sizeEarth);
                    dataRow.NumberOfEarthCables = numEarth;
                    dataRow.EarthCableSize = sizeEarth;

                    cableDataList.Add(dataRow);
                }
            }
            return cableDataList;
        }

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
                        fieldBuilder.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
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

            fields.Add(fieldBuilder.ToString());
            return fields;
        }

        private List<ModelGeneratedData> ParseAndProcessModelGeneratedCsvData(string filePath, List<string> requiredHeaders)
        {
            TaskDialog.Show("Information", "This command does not directly import 'Model Generated Data' via CSV. It is managed by other scripts.");
            return new List<ModelGeneratedData>();
        }

        private void ProcessSplit(string inputValue, out string countPart, out string sizePart)
        {
            string cleanedInput = inputValue.Trim();
            string[] parts = cleanedInput.Split(new[] { '×', 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 1)
            {
                countPart = parts[0].Trim();
                sizePart = parts[1].Trim();
            }
            else
            {
                countPart = "1";
                sizePart = cleanedInput;
            }
        }

        private string GetValueOrDefault(List<string> values, Dictionary<string, int> map, string headerName)
        {
            if (map.TryGetValue(headerName, out int index) && index < values.Count)
            {
                return values[index].Trim();
            }
            return string.Empty;
        }
        #endregion
    }

    /// <summary>
    /// Enum defining the available user actions in the PC_Extensible_Window
    /// </summary>
    public enum UserAction
    {
        None,
        ImportMyData,
        ImportConsultantData,
        ClearData
    }
}