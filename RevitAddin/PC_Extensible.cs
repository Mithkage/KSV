//
// File: PC_Extensible.cs
//
// Namespace: PC_Extensible
//
// Class: PC_ExtensibleClass
//
// Function: This Revit external command provides options to either import
//           cable data from a selected folder (processing multiple "Cables Summary" CSVs)
//           or import consultant cable data from a folder (processing multiple CSVs)
//           or clear existing cable data from either or both extensible storage locations.
//           A third extensible storage is defined for 'Model Generated Data' for use by other scripts.
//           When importing, if updated data is null or blank, existing stored data is retained.
//
// Author: Kyle Vorster (Modified by AI)
//
// Log:
// - July 2, 2025: Removed CSV Export functionality and added Consultant Import/Clear options.
// - July 2, 2025: Added FileName and ImportDate columns to CableData.
// - July 2, 2025: Added new storage for Model Generated Data, removed direct CSV import for it as per user feedback.
// - July 2, 2025: Included parsing for 'Sheath', 'Cable Max. Length (m)', and 'Voltage (Vac)' in CableData.
// - July 3, 2025: Modified import logic to allow selection of a folder and process multiple "Cables Summary" CSV files within it.
// - July 3, 2025: Adjusted aggregation logic for bulk imports: first processed file now takes precedence for duplicate Cable References.
// - July 3, 2025: Resolved multiple compilation errors (CS0019, CS1501, CS1061, CS1503) by adjusting type comparisons,
//                 using compatible string/dictionary methods, and ensuring argument type consistency.
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms; // For FolderBrowserDialog
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage;
using System.Text.Json; // For JSON serialization
using System.Text.Json.Serialization; // For JSON ignore attribute if needed
using System.Diagnostics; // Added for Debug.WriteLine
using System.Reflection; // For property iteration
#endregion

namespace PC_Extensible
{
    /// <summary>
    /// The main class for the Revit external command.
    /// Implements the IExternalCommand interface.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_ExtensibleClass : IExternalCommand
    {
        // Define unique GUIDs and names for your Schemas. These GUIDs must be truly unique.
        // PRIMARY DATA SCHEMA: Made public for access from other commands (e.g., RT_TrayOccupancy)
        public static readonly Guid PrimarySchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336"); // Existing GUID
        public const string PrimarySchemaName = "PC_ExtensibleDataSchema";
        public const string PrimaryFieldName = "PC_DataJson";
        public const string PrimaryDataStorageElementName = "PC_Extensible_PC_Data_Storage";

        // CONSULTANT DATA SCHEMA: Made public for access from other commands if needed
        public static readonly Guid ConsultantSchemaGuid = new Guid("B5E7F8C0-1234-5678-9ABC-DEF012345678"); // NEW unique GUID for consultant data
        public const string ConsultantSchemaName = "PC_ExtensibleConsultantDataSchema";
        public const string ConsultantFieldName = "PC_ConsultantDataJson";
        public const string ConsultantDataStorageElementName = "PC_Extensible_Consultant_Data_Storage";

        // MODEL GENERATED DATA SCHEMA: NEW unique GUID for model generated data (accessible by other scripts)
        public static readonly Guid ModelGeneratedSchemaGuid = new Guid("C7D8E9F0-1234-5678-9ABC-DEF012345678"); // Another NEW unique GUID
        public const string ModelGeneratedSchemaName = "PC_ExtensibleModelGeneratedDataSchema";
        public const string ModelGeneratedFieldName = "PC_ModelGeneratedDataJson";
        public const string ModelGeneratedDataStorageElementName = "PC_Extensible_Model_Generated_Data_Storage";

        // VENDOR ID: Made public as it's used in SchemaBuilder and should match your .addin manifest
        public const string VendorId = "ReTick_Solutions"; // Ensure this matches your .addin file VendorId

        /// <summary>
        /// The main entry point for the external command. Revit calls this method when the user clicks the button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Present user with import or clear options
            TaskDialog mainDialog = new TaskDialog("PC_Extensible Options");
            mainDialog.MainContent = "What operation would you like to perform on the cable data?";
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Import My PowerCAD Data (from Folder) and Update/Add to Project Storage"); // Updated text
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Import Consultant PowerCAD Data (from Folder) and Save to Separate Project Storage"); // Updated text
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Clear Existing Cable Data"); // Renamed to be more general
            mainDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            mainDialog.DefaultButton = TaskDialogResult.CommandLink1; // Default to Import My Data

            TaskDialogResult dialogResult = mainDialog.Show();

            if (dialogResult == TaskDialogResult.CommandLink1) // Import My Data (from Folder)
            {
                return ImportAndMergeData<CableData>(
                    doc,
                    PrimarySchemaGuid,
                    PrimarySchemaName,
                    PrimaryFieldName,
                    PrimaryDataStorageElementName,
                    "My PowerCAD Data",
                    ParseAndProcessCsvData // Pass the specific parser for CableData
                );
            }
            else if (dialogResult == TaskDialogResult.CommandLink3) // Import Consultant Data (from Folder)
            {
                return ImportAndMergeData<CableData>(
                    doc,
                    ConsultantSchemaGuid,
                    ConsultantSchemaName,
                    ConsultantFieldName,
                    ConsultantDataStorageElementName,
                    "Consultant PowerCAD Data",
                    ParseAndProcessCsvData // Use the same parser for CableData
                );
            }
            else if (dialogResult == TaskDialogResult.CommandLink2) // Clear Data options
            {
                // Offer choice to clear primary or consultant data, or both
                TaskDialog clearDialog = new TaskDialog("Clear Data Options");
                clearDialog.MainContent = "Which data set would you like to clear?";
                clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Clear My PowerCAD Data Only");
                clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Clear Consultant PowerCAD Data Only");
                clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Clear Model Generated Data Only"); // Still include a clear button for it
                clearDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Clear All Data (My, Consultant, and Model Generated)"); // Updated "Clear All"
                clearDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
                clearDialog.DefaultButton = TaskDialogResult.CommandLink1;

                TaskDialogResult clearResult = clearDialog.Show();

                // FIX: Use TaskDialogResult for comparison here
                if (clearResult == TaskDialogResult.CommandLink1)
                {
                    return ClearData(doc, PrimarySchemaGuid, PrimaryDataStorageElementName, "My PowerCAD Data");
                }
                else if (clearResult == TaskDialogResult.CommandLink2)
                {
                    return ClearData(doc, ConsultantSchemaGuid, ConsultantDataStorageElementName, "Consultant PowerCAD Data");
                }
                else if (clearResult == TaskDialogResult.CommandLink4) // Clear Model Generated Data
                {
                    return ClearData(doc, ModelGeneratedSchemaGuid, ModelGeneratedDataStorageElementName, "Model Generated Data");
                }
                else if (clearResult == TaskDialogResult.CommandLink3) // Clear All Data
                {
                    Result primaryClearResult = ClearData(doc, PrimarySchemaGuid, PrimaryDataStorageElementName, "My PowerCAD Data");
                    Result consultantClearResult = ClearData(doc, ConsultantSchemaGuid, ConsultantDataStorageElementName, "Consultant PowerCAD Data");
                    Result modelGeneratedClearResult = ClearData(doc, ModelGeneratedSchemaGuid, ModelGeneratedDataStorageElementName, "Model Generated Data");

                    if (primaryClearResult == Result.Succeeded && consultantClearResult == Result.Succeeded && modelGeneratedClearResult == Result.Succeeded)
                    {
                        TaskDialog.Show("Clear Data Complete", "All PowerCAD data (My, Consultant, and Model Generated) cleared successfully.");
                        return Result.Succeeded;
                    }
                    else if (primaryClearResult == Result.Cancelled || consultantClearResult == Result.Cancelled || modelGeneratedClearResult == Result.Cancelled)
                    {
                        message = "One or more clear operations were cancelled.";
                        return Result.Cancelled;
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
            else // Cancel main dialog
            {
                message = "Operation cancelled by user.";
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Handles the import, merge, and save logic for cable data from a selected folder.
        /// Made generic to work with different data types (CableData).
        /// </summary>
        /// <typeparam name="T">The type of data to import (e.g., CableData).</typeparam>
        /// <param name="doc">The Revit Document.</param>
        /// <param name="schemaGuid">The GUID for the schema to use.</param>
        /// <param name="schemaName">The name of the schema to use.</param>
        /// <param name="fieldName">The name of the field within the schema to use.</param>
        /// <param name="dataStorageElementName">The name of the DataStorage element to use.</param>
        /// <param name="dataTypeName">A descriptive name for the data type (e.g., "My PowerCAD Data", "Consultant PowerCAD Data").</param>
        /// <param name="csvParser">A delegate to the specific CSV parsing method for type T.</param>
        /// <returns>Result.Succeeded or Result.Failed/Cancelled.</returns>
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
                    // NEWLY ADDED COLUMNS:
                    "Sheath", "Cable Max. Length (m)", "Voltage (Vac)"
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

                // Temporary dictionary to handle duplicate Cable References from multiple files
                // The first file processed for a given Cable Reference will take precedence.
                var aggregatedNewDataDict = new Dictionary<string, T>(StringComparer.OrdinalIgnoreCase);

                foreach (string csvFilePath in csvFiles)
                {
                    string fileName = Path.GetFileName(csvFilePath);
                    // FIX: Use IndexOf for string search compatible with older .NET Framework versions
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
                                    // FIX: Manual TryAdd logic for compatibility, ensuring first-wins precedence
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

                // --- 5. MERGE AGGREGATED NEW DATA WITH EXISTING DATA (Conditional Update with additional duplicate check) ---
                // The primary merge key is 'To'. Handle duplicates in existing data if 'To' is not unique.
                var mergedDataDict = existingStoredData
                    .Where(item => !string.IsNullOrEmpty(GetPropertyValue(item, "To")))
                    .GroupBy(item => GetPropertyValue(item, "To"))
                    .ToDictionary(g => g.Key, g => g.First());


                // Keep track of Cable References already existing or newly added in this merge session
                var finalMergedCableReferences = new HashSet<string>(
                    existingStoredData.Where(item => !string.IsNullOrEmpty(GetPropertyValue(item, "CableReference"))).Select(item => GetPropertyValue(item, "CableReference"))
                );

                int updatedEntries = 0;
                int addedEntries = 0;
                List<string> duplicateCableReferencesReport = new List<string>(); // For new entries with duplicate Cable Ref where 'To' doesn't match a stored one

                foreach (var newEntry in allNewCsvDataFromFolder) // Iterate aggregated data (first-wins for intra-folder CableRef)
                {
                    string newEntryTo = GetPropertyValue(newEntry, "To");
                    string newEntryCableRef = GetPropertyValue(newEntry, "CableReference");

                    if (string.IsNullOrEmpty(newEntryTo))
                    {
                        Debug.WriteLine($"Skipping new aggregated entry due to empty 'To' value: Cable Ref: {newEntryCableRef ?? "N/A"}");
                        // If 'To' is empty, we cannot merge by 'To'. For now, skipping this entry from the merge.
                        continue;
                    }

                    if (mergedDataDict.TryGetValue(newEntryTo, out T existingEntry))
                    {
                        // Match found by 'To': Update existing entry's values
                        PropertyInfo[] properties = typeof(T).GetProperties();
                        bool entryUpdated = false;

                        foreach (PropertyInfo prop in properties)
                        {
                            // Skip FileName and ImportDate during merge for CableData, these should only be set on initial import
                            if (typeof(T) == typeof(CableData) && (prop.Name == nameof(CableData.FileName) || prop.Name == nameof(CableData.ImportDate)))
                            {
                                continue;
                            }

                            if (prop.PropertyType == typeof(string) && prop.CanWrite)
                            {
                                string newValue = (string)prop.GetValue(newEntry);
                                string currentValue = (string)prop.GetValue(existingEntry);

                                // Only update if the new value is not null or blank AND it's different from the current value
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
                        // No match found by 'To' value. Check for duplicate 'Cable Reference' across the *final merged dataset* before adding as new.
                        if (!string.IsNullOrEmpty(newEntryCableRef) && finalMergedCableReferences.Contains(newEntryCableRef))
                        {
                            // This means a CableRef exists, but its 'To' value is different from the current 'newEntryTo'.
                            // Or it's a duplicate CableRef from existing data that couldn't be merged by 'To'.
                            // We report it as a skipped duplicate for adding.
                            duplicateCableReferencesReport.Add($"Cable Reference: '{newEntryCableRef}', New To: '{newEntryTo}'");
                            Debug.WriteLine($"Skipping new aggregated entry due to duplicate 'Cable Reference' in final merge (no 'To' match): {newEntryCableRef}");
                            continue;
                        }
                        else
                        {
                            // Add as new entry to the merged dictionary
                            mergedDataDict.Add(newEntryTo, newEntry);
                            finalMergedCableReferences.Add(newEntryCableRef); // Track that this CableRef is now in the final set
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
                        SaveDataToExtensibleStorage(
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
                                   "\n\nACTION REQUIRED: Ensure 'schemaBuilder.SetVendorId(\"ReTick_Solutions\")' in this script " +
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
        /// Handles the clearing of all existing data from extensible storage for a specific schema.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <param name="schemaGuid">The GUID for the schema to clear.</param>
        /// <param name="dataStorageElementName">The name of the DataStorage element to clear.</param>
        /// <param name="dataTypeName">A descriptive name for the data type (e.g., "My PowerCAD Data", "Consultant PowerCAD Data").</param>
        /// <returns>Result.Succeeded or Result.Failed.</returns>
        private Result ClearData(Document doc, Guid schemaGuid, string dataStorageElementName, string dataTypeName)
        {
            TaskDialog confirmDialog = new TaskDialog("Confirm Clear Data");
            confirmDialog.MainContent = $"Are you sure you want to clear ALL existing {dataTypeName} from the project's extensible storage?\nThis action cannot be undone.";
            confirmDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            confirmDialog.DefaultButton = TaskDialogResult.No;

            if (confirmDialog.Show() == TaskDialogResult.Yes)
            {
                using (Transaction tx = new Transaction(doc, $"Clear {dataTypeName} from Extensible Storage"))
                {
                    tx.Start();
                    try
                    {
                        // Find the DataStorage element for our schema
                        FilteredElementCollector collector = new FilteredElementCollector(doc)
                            .OfClass(typeof(DataStorage));

                        DataStorage dataStorageToDelete = null;
                        foreach (DataStorage ds in collector)
                        {
                            // Check if this DataStorage element holds our schema's entity
                            Schema foundSchema = Schema.Lookup(schemaGuid);
                            if (foundSchema != null && ds.GetEntity(foundSchema) != null)
                            {
                                // Also check the specific name, especially important for multiple storage elements
                                if (ds.Name == dataStorageElementName)
                                {
                                    dataStorageToDelete = ds;
                                    break;
                                }
                            }
                        }

                        if (dataStorageToDelete != null)
                        {
                            doc.Delete(dataStorageToDelete.Id);
                            tx.Commit();
                            TaskDialog.Show("Clear Data Complete", $"All existing {dataTypeName} successfully cleared from project's extensible storage.");
                            return Result.Succeeded;
                        }
                        else
                        {
                            tx.RollBack(); // No actual change, but ensure transaction is handled
                            TaskDialog.Show("Clear Data Info", $"No existing {dataTypeName} found in project's extensible storage to clear.");
                            return Result.Succeeded; // Still a successful outcome from user perspective
                        }
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        string msg = $"An error occurred while clearing {dataTypeName}: {ex.Message}";
                        if (ex.Message.Contains("Writing of Entities of this Schema is not allowed to the current add-in"))
                        {
                            msg += "\n\nDIAGNOSIS: This often indicates a mismatch in the 'VendorId' specified in the SchemaBuilder " +
                                   "and the actual VendorId of your Revit add-in. " +
                                   "\n\nACTION REQUIRED: Ensure 'schemaBuilder.SetVendorId(\"ReTick_Solutions\")' in this script " +
                                   "exactly matches the 'VendorId' attribute in your .addin manifest file, or the VendorId used to register your add-in. " +
                                   "This is crucial for write permissions to Extensible Storage.";
                        }
                        TaskDialog.Show("Clear Data Error", msg);
                        return Result.Failed;
                    }
                }
            }
            else
            {
                TaskDialog.Show("Clear Data Cancelled", $"Clear {dataTypeName} operation cancelled.");
                return Result.Cancelled;
            }
        }

        #region Extensible Storage Methods (Generic)

        /// <summary>
        /// Gets or creates the Schema for storing generic data.
        /// </summary>
        private Schema GetOrCreateCableDataSchema(Guid schemaGuid, string schemaName, string fieldName)
        {
            Schema schema = Schema.Lookup(schemaGuid);

            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid);
                schemaBuilder.SetSchemaName(schemaName);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor); // Only this add-in can write to this schema

                schemaBuilder.SetVendorId(VendorId); // Use the common VendorId defined at the class level

                // Create a field to store the JSON string of all data
                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(fieldName, typeof(string));

                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        /// <summary>
        /// Gets the existing DataStorage element, or creates a new one if it doesn't exist.
        /// This method assumes a transaction is open if called with an intent to create/modify.
        /// </summary>
        private DataStorage GetOrCreateDataStorage(Document doc, Guid schemaGuid, string dataStorageElementName)
        {
            // Find existing DataStorage elements for our schema
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                // Check if this DataStorage element holds our schema's entity AND has the correct name
                Schema foundSchema = Schema.Lookup(schemaGuid);
                if (foundSchema != null && ds.GetEntity(foundSchema) != null)
                {
                    // Also check the specific name, especially important for multiple storage elements
                    if (ds.Name == dataStorageElementName)
                    {
                        dataStorage = ds;
                        break;
                    }
                }
            }

            if (dataStorage == null)
            {
                // No existing DataStorage found, create a new one
                dataStorage = DataStorage.Create(doc);
                dataStorage.Name = dataStorageElementName; // Assign a name for easier identification in Revit
            }

            return dataStorage;
        }

        /// <summary>
        /// Saves a list of generic data to extensible storage.
        /// This method assumes a transaction is already open.
        /// </summary>
        /// <typeparam name="T">The type of data to save.</typeparam>
        public void SaveDataToExtensibleStorage<T>( // Made public for use by other scripts
            Document doc,
            List<T> dataList,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName)
        {
            Schema schema = GetOrCreateCableDataSchema(schemaGuid, schemaName, fieldName);
            DataStorage dataStorage = GetOrCreateDataStorage(doc, schemaGuid, dataStorageElementName); // Ensure DataStorage exists within transaction

            // Serialize the list of data to a JSON string
            var options = new JsonSerializerOptions { WriteIndented = true }; // For readable JSON
            string jsonString = JsonSerializer.Serialize(dataList, options);

            // Create an Entity to store the data
            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(fieldName), jsonString);

            // Set the entity to the DataStorage element
            dataStorage.SetEntity(entity);
        }

        /// <summary>
        /// Recalls a list of generic data from extensible storage.
        /// </summary>
        /// <typeparam name="T">The type of data to recall.</typeparam>
        /// <returns>A List of T, or an empty list if no data is found or an error occurs during recall.</returns>
        public List<T> RecallDataFromExtensibleStorage<T>( // Made public for use by other scripts
            Document doc,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName)
        {
            Schema schema = Schema.Lookup(schemaGuid); // Look up the schema by its GUID

            if (schema == null)
            {
                // Schema does not exist in this project yet, so no data is stored under this schema.
                return new List<T>();
            }

            // Find existing DataStorage elements associated with our schema and name
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                // Check if the DataStorage element contains an entity for our schema AND has the correct name
                if (ds.Name == dataStorageElementName && ds.GetEntity(schema) != null)
                {
                    dataStorage = ds;
                    break;
                }
            }

            if (dataStorage == null)
            {
                // No DataStorage element found containing our schema's data
                return new List<T>();
            }

            // Get the Entity from the DataStorage
            Entity entity = dataStorage.GetEntity(schema);

            if (!entity.IsValid())
            {
                return new List<T>();
            }

            // Get the JSON string from the field
            string jsonString = entity.Get<string>(schema.GetField(fieldName));

            if (string.IsNullOrEmpty(jsonString))
            {
                return new List<T>();
            }

            // Deserialize the JSON string back into a List<T>
            try
            {
                return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>();
            }
            catch (JsonException ex)
            {
                // Use TaskDialog.Show for user-friendly error messages in Revit context
                TaskDialog.Show("Data Recall Error", $"Failed to deserialize stored {schemaName} data: {ex.Message}. The stored data might be corrupt or incompatible.");
                return new List<T>();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Data Recall Error", $"An unexpected error occurred during {schemaName} data recall: {ex.Message}");
                return new List<T>();
            }
        }

        #endregion

        #region File Dialog Methods
        /// <summary>
        /// Prompts the user to select a source folder.
        /// </summary>
        private string GetSourceFolderPath(string dialogTitle)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = dialogTitle;
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

        #region CSV Parsing and Processing (Specific to Data Types)

        // Parser for CableData (Your original data and Consultant data)
        private List<CableData> ParseAndProcessCsvData(string filePath, List<string> requiredHeaders)
        {
            var cableDataList = new List<CableData>();
            var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); // Case-insensitive header matching
            string[] headers = null;

            string fileName = Path.GetFileName(filePath); // Get the file name from the path
            string importDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // Get current date/time

            using (var reader = new StreamReader(filePath, Encoding.Default))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    // Look for the header line starting with "Cable Reference"
                    if (headers == null && line.Trim().StartsWith("Cable Reference", StringComparison.OrdinalIgnoreCase))
                    {
                        headers = ParseCsvLine(line).Select(h => h.Trim()).ToArray();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            string header = headers[i];
                            string cleanHeader = header.Replace("Â²", "\u00B2"); // Handle potential encoding issues for squared symbol
                            if (requiredHeaders.Contains(cleanHeader) && !headerMap.ContainsKey(cleanHeader))
                            {
                                headerMap[cleanHeader] = i;
                            }
                        }
                        continue; // Skip the header line itself
                    }

                    // Skip empty lines or lines that are not data (e.g., "Transformer:")
                    if (headers == null || string.IsNullOrWhiteSpace(line) || line.StartsWith("Transformer:", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    List<string> valuesList = ParseCsvLine(line);
                    // FIX: Use valuesList directly instead of converting to array, to match GetValueOrDefault expected type
                    // string[] values = valuesList.ToArray(); 

                    // Ensure there are enough values for the headers we care about
                    if (valuesList.Count < headerMap.Values.Max() + 1) // Using valuesList directly
                    {
                        Debug.WriteLine($"Skipping malformed line in CSV: {line}");
                        continue;
                    }

                    string conductorActiveValue = GetValueOrDefault(valuesList, headerMap, "Conductor (Active)"); // Using valuesList
                    if (string.IsNullOrWhiteSpace(conductorActiveValue))
                    {
                        // Skip rows that don't have a Conductor (Active) value as per original logic
                        continue;
                    }

                    var dataRow = new CableData
                    {
                        FileName = fileName, // Set the new FileName property
                        ImportDate = importDate, // Set the new ImportDate property
                        CableReference = GetValueOrDefault(valuesList, headerMap, "Cable Reference"), // Using valuesList
                        From = GetValueOrDefault(valuesList, headerMap, "From"), // Using valuesList
                        To = GetValueOrDefault(valuesList, headerMap, "To"), // Using valuesList
                        CableType = GetValueOrDefault(valuesList, headerMap, "Cable Type"), // Using valuesList
                        CableCode = GetValueOrDefault(valuesList, headerMap, "Cable Code"), // Using valuesList
                        CableConfiguration = GetValueOrDefault(valuesList, headerMap, "Cable Configuration"), // Using valuesList
                        Cores = GetValueOrDefault(valuesList, headerMap, "Cores"), // Using valuesList
                        ConductorActive = conductorActiveValue,
                        Insulation = GetValueOrDefault(valuesList, headerMap, "Insulation"), // Using valuesList
                        ConductorEarth = GetValueOrDefault(valuesList, headerMap, "Conductor (Earth)"), // Using valuesList
                        SeparateEarthForMulticore = GetValueOrDefault(valuesList, headerMap, "Separate Earth for Multicore"), // Using valuesList
                        CableLength = GetValueOrDefault(valuesList, headerMap, "Cable Length (m)"), // Using valuesList
                        TotalCableRunWeight = GetValueOrDefault(valuesList, headerMap, "Total Cable Run Weight (Incl. N & E) (kg)"), // Using valuesList
                        NominalOverallDiameter = GetValueOrDefault(valuesList, headerMap, "Nominal Overall Diameter (mm)"), // Using valuesList
                        AsNsz3008CableDeratingFactor = GetValueOrDefault(valuesList, headerMap, "AS/NSZ 3008 Cable Derating Factor"), // Using valuesList
                        // NEWLY ADDED COLUMNS:
                        Sheath = GetValueOrDefault(valuesList, headerMap, "Sheath"), // Using valuesList
                        CableMaxLengthM = GetValueOrDefault(valuesList, headerMap, "Cable Max. Length (m)"), // Using valuesList
                        VoltageVac = GetValueOrDefault(valuesList, headerMap, "Voltage (Vac)") // Using valuesList
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

                    string activeCableValue = GetValueOrDefault(valuesList, headerMap, "Cable Size (mm\u00B2)"); // Using valuesList
                    string numActive, sizeActive;
                    ProcessSplit(activeCableValue, out numActive, out sizeActive);
                    dataRow.NumberOfActiveCables = numActive;
                    dataRow.ActiveCableSize = sizeActive;

                    string neutralCableValue = GetValueOrDefault(valuesList, headerMap, "Neutral Size (mm\u00B2)"); // Using valuesList
                    string numNeutral, sizeNeutral;
                    ProcessSplit(neutralCableValue, out numNeutral, out sizeNeutral);
                    dataRow.NumberOfNeutralCables = numNeutral;
                    dataRow.NeutralCableSize = sizeNeutral;

                    string earthCableValue = GetValueOrDefault(valuesList, headerMap, "Earth Size (mm\u00B2)"); // Using valuesList
                    if (earthCableValue.Equals("No Earth", StringComparison.OrdinalIgnoreCase))
                    {
                        earthCableValue = "0"; // Normalize "No Earth" to "0" for splitting
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

        // Generic CSV line parsing method
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

        // Dummy parser for ModelGeneratedData. This method exists to satisfy the `Func` delegate
        // in `ImportAndMergeData`, even though this specific UI command will not be importing it via CSV.
        // Other scripts will call SaveDataToExtensibleStorage<ModelGeneratedData> directly.
        private List<ModelGeneratedData> ParseAndProcessModelGeneratedCsvData(string filePath, List<string> requiredHeaders)
        {
            // This method is intentionally left minimal or can throw an exception if called,
            // as Model Generated Data is not intended to be imported via CSV through this UI.
            TaskDialog.Show("Information", "This command does not directly import 'Model Generated Data' via CSV. It is managed by other scripts.");
            return new List<ModelGeneratedData>(); // Return an empty list or throw an exception if you prefer.
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

        // FIX: Change parameter type from string[] to List<string> for values, to match ParseCsvLine return type
        private string GetValueOrDefault(List<string> values, Dictionary<string, int> map, string headerName)
        {
            if (map.TryGetValue(headerName, out int index) && index < values.Count)
            {
                return values[index].Trim();
            }
            return string.Empty;
        }
        #endregion

        #region Data Classes
        /// <summary>
        /// A data class to hold the values for a single row of processed cable data from the CSV.
        /// Public for JSON serialization and accessibility from other assemblies.
        /// Used for My PowerCAD Data and Consultant PowerCAD Data.
        /// </summary>
        public class CableData
        {
            // New fields for provenance
            public string FileName { get; set; } // Name of the CSV file this record came from
            public string ImportDate { get; set; } // Date and time of import

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
            public string AsNsz3008CableDeratingFactor { get; set; } // Added as it's read and used in calculations

            // Properties derived/calculated in ParseAndProcessCsvData, now included for storage:
            public string NumberOfActiveCables { get; set; }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; } // Calculated in ParseAndProcessCsvData and now stored

            // NEWLY ADDED CABLE DATA PROPERTIES:
            public string Sheath { get; set; } // "Sheath"
            public string CableMaxLengthM { get; set; } // "Cable Max. Length (m)"
            public string VoltageVac { get; set; } // "Voltage (Vac)"
        }

        /// <summary>
        /// A data class to hold values for model-generated data.
        /// Public for JSON serialization and accessibility from other assemblies.
        /// </summary>
        public class ModelGeneratedData
        {
            public string CableReference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string CableLengthM { get; set; } // Using 'M' suffix to differentiate from CableData.CableLength
            public string Variance { get; set; }
            public string Comment { get; set; }
        }
        #endregion
    }
}