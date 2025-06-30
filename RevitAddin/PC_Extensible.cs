//
// File: PC_Extensible.cs
//
// Namespace: PC_Extensible
//
// Class: PC_ExtensibleClass
//
// Function: This Revit external command provides options to either import
//           new cable data from a CSV (merging/updating existing data in extensible storage),
//           clear all existing cable data from extensible storage, or export
//           the current data from extensible storage to a CSV.
//           It also exports the current cleaned/merged data to "Cleaned_Cable_Schedule.csv" during import.
//           When importing, if updated data is null or blank, existing stored data is retained.
//
// Author: Kyle Vorster (Modified by AI)
//
// Date: June 30, 2025 (Updated with Import/Clear/Export Logic and Conditional Merge)
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
        // Define a unique GUID for your Schema. This GUID must be truly unique for your application.
        // You can generate a new one using online GUID generators or Visual Studio's "Create GUID" tool.
        private static readonly Guid SchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336"); // Keep this GUID unique
        private const string SchemaName = "PC_ExtensibleDataSchema"; // General schema name
        private const string FieldName = "PC_DataJson"; // Field to store the JSON string (now explicitly "PC_Data")
        private const string DataStorageElementName = "PC_Extensible_PC_Data_Storage"; // Name of the DataStorage element
        private const string DefaultExportFileName = "Cleaned_Cable_Schedule_Exported.csv"; // Default name for export

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
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Import Data (from CSV) and Update/Add to Project Storage");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Clear All Existing Cable Data from Project Storage");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Export Current Data from Project Storage to CSV"); // New Option
            mainDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            mainDialog.DefaultButton = TaskDialogResult.CommandLink1; // Default to Import

            TaskDialogResult dialogResult = mainDialog.Show();

            if (dialogResult == TaskDialogResult.CommandLink1) // Import Data
            {
                return ImportAndMergeData(doc);
            }
            else if (dialogResult == TaskDialogResult.CommandLink2) // Clear Data
            {
                return ClearAllData(doc);
            }
            else if (dialogResult == TaskDialogResult.CommandLink3) // Export Current Data
            {
                return ExportCurrentData(doc);
            }
            else // Cancel
            {
                message = "Operation cancelled by user.";
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Handles the import, merge, and save logic for cable data.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>Result.Succeeded or Result.Failed/Cancelled.</returns>
        private Result ImportAndMergeData(Document doc)
        {
            // --- 1. PROMPT USER FOR INPUT CSV FILE ---
            string sourceCsvPath = GetSourceCsvFilePath();
            if (string.IsNullOrEmpty(sourceCsvPath))
            {
                return Result.Cancelled; // User cancelled file selection
            }

            try
            {
                // --- 2. DEFINE THE COLUMNS TO EXTRACT ---
                var columnsToRead = new List<string>
                {
                    "Cable Reference", "From", "To", "Cable Type", "Cable Code", "Cable Configuration",
                    "Cores", "Cable Size (mm\u00B2)", "Conductor (Active)", "Insulation", "Neutral Size (mm\u00B2)",
                    "Earth Size (mm\u00B2)", "Conductor (Earth)", "Separate Earth for Multicore", "Cable Length (m)",
                    "Total Cable Run Weight (Incl. N & E) (kg)", "Nominal Overall Diameter (mm)", "AS/NSZ 3008 Cable Derating Factor"
                };

                // --- 3. READ, PARSE, AND CLEAN THE NEW CSV DATA ---
                List<CableData> newCsvData = ParseAndProcessCsvData(sourceCsvPath, columnsToRead);

                if (newCsvData.Count == 0)
                {
                    TaskDialog.Show("No New Data", "No valid cable data was found in the selected CSV file. No changes will be made to project data.");
                    return Result.Succeeded;
                }

                // --- 4. RECALL EXISTING DATA FROM EXTENSIBLE STORAGE ---
                List<CableData> existingStoredData = RecallCableDataFromExtensibleStorage(doc);

                // --- 5. MERGE NEW DATA WITH EXISTING DATA (Conditional Update) ---
                // Create a dictionary from existing data for efficient lookups by 'To' key
                var mergedDataDict = existingStoredData
                                           .Where(cd => !string.IsNullOrEmpty(cd.To))
                                           .ToDictionary(cd => cd.To, cd => cd);

                int updatedEntries = 0;
                int addedEntries = 0;

                foreach (var newCableEntry in newCsvData)
                {
                    if (string.IsNullOrEmpty(newCableEntry.To))
                    {
                        // Skip new entries that don't have a 'To' value, as we can't key them
                        Debug.WriteLine($"Skipping new CSV entry due to empty 'To' value: Cable Ref: {newCableEntry.CableReference ?? "N/A"}");
                        continue;
                    }

                    if (mergedDataDict.TryGetValue(newCableEntry.To, out CableData existingCableEntry))
                    {
                        // Match found: Iterate through properties and update if new value is not null/blank
                        PropertyInfo[] properties = typeof(CableData).GetProperties();
                        bool entryUpdated = false;

                        foreach (PropertyInfo prop in properties)
                        {
                            // We only care about string properties for this logic
                            if (prop.PropertyType == typeof(string) && prop.CanWrite)
                            {
                                string newValue = (string)prop.GetValue(newCableEntry);

                                // If the new value is not null or blank, update it
                                if (!string.IsNullOrWhiteSpace(newValue))
                                {
                                    string currentValue = (string)prop.GetValue(existingCableEntry);
                                    if (currentValue != newValue) // Only update if value actually changed
                                    {
                                        prop.SetValue(existingCableEntry, newValue);
                                        entryUpdated = true;
                                    }
                                }
                                // Else, if the new value IS null or blank, we retain the existing value
                                // So, no action is needed here as existingCableEntry already has the old value.
                            }
                        }
                        if (entryUpdated) updatedEntries++;
                    }
                    else
                    {
                        // No match found: Add the new entry as is
                        mergedDataDict.Add(newCableEntry.To, newCableEntry);
                        addedEntries++;
                    }
                }

                List<CableData> finalMergedData = mergedDataDict.Values.ToList();

                // --- 6. PROMPT USER FOR OUTPUT CSV FILE PATH (for the *merged* data) ---
                string outputCsvFilePath = GetOutputCsvFilePath();
                if (string.IsNullOrEmpty(outputCsvFilePath))
                {
                    return Result.Cancelled; // User cancelled file selection for output
                }

                // --- 7. EXPORT THE MERGED DATA TO A NEW CSV ---
                ExportDataToCsv(finalMergedData, outputCsvFilePath);

                // --- 8. SAVE MERGED DATA TO EXTENSIBLE STORAGE ---
                using (Transaction tx = new Transaction(doc, "Save Merged PC_Data to Extensible Storage"))
                {
                    tx.Start();
                    try
                    {
                        SaveCableDataToExtensibleStorage(doc, finalMergedData);
                        tx.Commit();
                        TaskDialog.Show("Extensible Storage Save", $"Merged cable data (PC_Data) saved to project's extensible storage successfully.\n" +
                                                                   $"Updated entries: {updatedEntries}\nAdded entries: {addedEntries}");
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        string msg = $"Failed to save merged data to extensible storage: {ex.Message}";
                        // Add more diagnostic info for the common VendorId mismatch issue
                        if (ex.Message.Contains("Writing of Entities of this Schema is not allowed to the current add-in"))
                        {
                            msg += "\n\nDIAGNOSIS: This often indicates a mismatch in the 'VendorId' specified in the SchemaBuilder " +
                                   "and the actual VendorId of your Revit add-in. " +
                                   "\n\nACTION REQUIRED: Ensure 'schemaBuilder.SetVendorId(\"ReTick_Solutions\")' in this script " +
                                   "exactly matches the 'VendorId' attribute in your .addin manifest file, or the VendorId used to register your add-in. " +
                                   "This is crucial for write permissions to Extensible Storage.";
                        }
                        TaskDialog.Show("Extensible Storage Error", msg);
                        return Result.Failed; // Indicate failure
                    }
                }

                // --- 9. NOTIFY USER OF OVERALL SUCCESS ---
                TaskDialog.Show("Process Complete", $"Import and merge process complete.\n\n" +
                                                   $"The 'Cleaned_Cable_Schedule.csv' file (containing merged data) has been saved to:\n'{outputCsvFilePath}'\n\n" +
                                                   $"The cleaned cable data (PC_Data) has also been successfully updated in extensible storage in the current Revit project for later recall.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                string msg = $"An unexpected error occurred during import/merge process:\n{ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", msg);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Handles the clearing of all existing cable data from extensible storage.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>Result.Succeeded or Result.Failed.</returns>
        private Result ClearAllData(Document doc)
        {
            TaskDialog confirmDialog = new TaskDialog("Confirm Clear Data");
            confirmDialog.MainContent = "Are you sure you want to clear ALL existing cable data from the project's extensible storage?\nThis action cannot be undone.";
            confirmDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            confirmDialog.DefaultButton = TaskDialogResult.No;

            if (confirmDialog.Show() == TaskDialogResult.Yes)
            {
                using (Transaction tx = new Transaction(doc, "Clear PC_Data from Extensible Storage"))
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
                            // Note: Schema.Lookup(SchemaGuid) might return null if the schema hasn't been created yet.
                            // It's safer to get the schema first.
                            Schema foundSchema = Schema.Lookup(SchemaGuid);
                            if (foundSchema != null && ds.GetEntity(foundSchema) != null)
                            {
                                dataStorageToDelete = ds;
                                break;
                            }
                        }

                        if (dataStorageToDelete != null)
                        {
                            doc.Delete(dataStorageToDelete.Id);
                            tx.Commit();
                            TaskDialog.Show("Clear Data Complete", "All existing cable data (PC_Data) successfully cleared from project's extensible storage.");
                            return Result.Succeeded;
                        }
                        else
                        {
                            tx.RollBack(); // No actual change, but ensure transaction is handled
                            TaskDialog.Show("Clear Data Info", "No existing cable data (PC_Data) found in project's extensible storage to clear.");
                            return Result.Succeeded; // Still a successful outcome from user perspective
                        }
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        string msg = $"An error occurred while clearing data: {ex.Message}";
                        // Add more diagnostic info for the common VendorId mismatch issue
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
                TaskDialog.Show("Clear Data Cancelled", "Clear data operation cancelled.");
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Handles the export of all existing cable data from extensible storage to a CSV file.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>Result.Succeeded or Result.Failed/Cancelled.</returns>
        private Result ExportCurrentData(Document doc)
        {
            try
            {
                // 1. Recall existing data from extensible storage
                List<CableData> storedData = RecallCableDataFromExtensibleStorage(doc);

                if (storedData == null || !storedData.Any())
                {
                    TaskDialog.Show("No Data to Export", "No cable data (PC_Data) found in the project's extensible storage to export.");
                    return Result.Succeeded;
                }

                // 2. Prompt user for output folder
                string outputFolderPath = GetOutputFolderPathForExport();
                if (string.IsNullOrEmpty(outputFolderPath))
                {
                    return Result.Cancelled; // User cancelled folder selection
                }

                // 3. Define the full export file path
                string exportFilePath = Path.Combine(outputFolderPath, DefaultExportFileName);

                // 4. Export the data to CSV
                ExportDataToCsv(storedData, exportFilePath);

                TaskDialog.Show("Export Complete", $"Current cable data (PC_Data) successfully exported to:\n'{exportFilePath}'");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                string msg = $"An unexpected error occurred during export process:\n{ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Export Error", msg);
                return Result.Failed;
            }
        }


        #region Extensible Storage Methods

        /// <summary>
        /// Gets or creates the Schema for storing cable data.
        /// </summary>
        /// <returns>The Schema object.</returns>
        private Schema GetOrCreateCableDataSchema()
        {
            Schema schema = Schema.Lookup(SchemaGuid);

            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(SchemaGuid);
                schemaBuilder.SetSchemaName(SchemaName);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor); // Only this add-in can write to this schema

                // IMPORTANT: The VendorId MUST exactly match the VendorId declared for your Revit Add-in.
                // This is typically specified in your add-in's .addin manifest file, e.g.:
                // <AddIn Type="Command">
                //    <VendorId>ReTick_Solutions</VendorId>
                //    ...
                // </AddIn>
                // Make sure "ReTick_Solutions" (or whatever you use here) is consistent.
                schemaBuilder.SetVendorId("ReTick_Solutions"); // Updated to match .addin file VendorId

                // Create a field to store the JSON string of all CableData
                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(FieldName, typeof(string));

                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        /// <summary>
        /// Gets the existing DataStorage element for cable data, or creates a new one if it doesn't exist.
        /// This method assumes a transaction is open if called with an intent to create/modify.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>The DataStorage element.</returns>
        private DataStorage GetOrCreateDataStorage(Document doc)
        {
            // Find existing DataStorage elements for our schema
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                // Check if this DataStorage element holds our schema's entity
                // Note: Schema.Lookup(SchemaGuid) might return null if the schema hasn't been created yet.
                // It's safer to get the schema first.
                Schema foundSchema = Schema.Lookup(SchemaGuid);
                if (foundSchema != null && ds.GetEntity(foundSchema) != null)
                {
                    dataStorage = ds;
                    break;
                }
            }

            if (dataStorage == null)
            {
                // No existing DataStorage found, create a new one
                dataStorage = DataStorage.Create(doc);
                dataStorage.Name = DataStorageElementName; // Assign a name for easier identification in Revit
            }

            return dataStorage;
        }

        /// <summary>
        /// Saves the cleaned cable data to extensible storage.
        /// This method assumes a transaction is already open.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <param name="cableDataList">The list of CableData to save.</param>
        private void SaveCableDataToExtensibleStorage(Document doc, List<CableData> cableDataList)
        {
            Schema schema = GetOrCreateCableDataSchema();
            DataStorage dataStorage = GetOrCreateDataStorage(doc); // Ensure DataStorage exists within transaction

            // Serialize the list of CableData to a JSON string
            var options = new JsonSerializerOptions { WriteIndented = true }; // For readable JSON
            string jsonString = JsonSerializer.Serialize(cableDataList, options);

            // Create an Entity to store the data
            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(FieldName), jsonString);

            // Set the entity to the DataStorage element
            dataStorage.SetEntity(entity);
        }

        /// <summary>
        /// Recalls the cleaned cable data (PC_Data) from extensible storage.
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
        private string GetSourceCsvFilePath()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select Cable Data CSV File to Import";
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
        /// Prompts the user to select a file path to save the output CSV during import.
        /// </summary>
        /// <returns>The selected file path, or null if cancelled.</returns>
        private string GetOutputCsvFilePath()
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Save Cleaned Cable Schedule CSV File (During Import)";
                saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "csv";
                saveFileDialog.FileName = "Cleaned_Cable_Schedule.csv"; // Default filename
                saveFileDialog.RestoreDirectory = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// Prompts the user to select a folder to save the exported CSV file.
        /// </summary>
        /// <returns>The selected folder path, or null if cancelled.</returns>
        private string GetOutputFolderPathForExport()
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a Folder to Save the Exported Cable Data CSV";
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

        #region CSV Parsing and Processing
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

        private List<CableData> ParseAndProcessCsvData(string filePath, List<string> requiredHeaders)
        {
            var cableDataList = new List<CableData>();
            var headerMap = new Dictionary<string, int>();
            string[] headers = null;

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
                    string[] values = valuesList.ToArray();

                    // Ensure there are enough values for the headers we care about
                    if (values.Length < headerMap.Values.Max() + 1) // +1 because Max() gives 0-based index
                    {
                        // Log or inform if a line is malformed, but don't stop the process
                        Debug.WriteLine($"Skipping malformed line in CSV: {line}");
                        continue;
                    }


                    string conductorActiveValue = GetValueOrDefault(values, headerMap, "Conductor (Active)");
                    if (string.IsNullOrWhiteSpace(conductorActiveValue))
                    {
                        // Skip rows that don't have a Conductor (Active) value as per original logic
                        continue;
                    }

                    var dataRow = new CableData
                    {
                        CableReference = GetValueOrDefault(values, headerMap, "Cable Reference"),
                        From = GetValueOrDefault(values, headerMap, "From"),
                        To = GetValueOrDefault(values, headerMap, "To"),
                        CableType = GetValueOrDefault(values, headerMap, "Cable Type"),
                        CableCode = GetValueOrDefault(values, headerMap, "Cable Code"),
                        CableConfiguration = GetValueOrDefault(values, headerMap, "Cable Configuration"),
                        Cores = GetValueOrDefault(values, headerMap, "Cores"),
                        ConductorActive = conductorActiveValue,
                        Insulation = GetValueOrDefault(values, headerMap, "Insulation"),
                        ConductorEarth = GetValueOrDefault(values, headerMap, "Conductor (Earth)"),
                        SeparateEarthForMulticore = GetValueOrDefault(values, headerMap, "Separate Earth for Multicore"),
                        CableLength = GetValueOrDefault(values, headerMap, "Cable Length (m)"),
                        TotalCableRunWeight = GetValueOrDefault(values, headerMap, "Total Cable Run Weight (Incl. N & E) (kg)"),
                        NominalOverallDiameter = GetValueOrDefault(values, headerMap, "Nominal Overall Diameter (mm)"),
                        AsNsz3008CableDeratingFactor = GetValueOrDefault(values, headerMap, "AS/NSZ 3008 Cable Derating Factor")
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

                    string activeCableValue = GetValueOrDefault(values, headerMap, "Cable Size (mm\u00B2)");
                    string numActive, sizeActive;
                    ProcessSplit(activeCableValue, out numActive, out sizeActive);
                    dataRow.NumberOfActiveCables = numActive;
                    dataRow.ActiveCableSize = sizeActive;

                    string neutralCableValue = GetValueOrDefault(values, headerMap, "Neutral Size (mm\u00B2)");
                    string numNeutral, sizeNeutral;
                    ProcessSplit(neutralCableValue, out numNeutral, out sizeNeutral);
                    dataRow.NumberOfNeutralCables = numNeutral;
                    dataRow.NeutralCableSize = sizeNeutral;

                    string earthCableValue = GetValueOrDefault(values, headerMap, "Earth Size (mm\u00B2)");
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

        #region Data Classes
        /// <summary>
        /// A data class to hold the values for a single row of processed cable data from the CSV.
        /// Public for JSON serialization.
        /// </summary>
        public class CableData
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
            public string NumberOfActiveCables
            {
                get;
                set;
            }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; }
            public string AsNsz3008CableDeratingFactor { get; set; }
        }
        #endregion
    }
}