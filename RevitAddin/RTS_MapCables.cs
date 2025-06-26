// RTS_MapCables.cs
//
// This file defines the main class and entry point for the RTS_MapCables Revit add-in.
// It provides functionality to map RTS (Rack, Tray, and Support) parameters with
// client-specific parameters, typically read from a CSV file. This allows for
// automated population of RTS data within Revit models based on client-defined data.
//
namespace RTS_MapCables
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.DB.Electrical;
    using Autodesk.Revit.UI;
    using System; // Required for Guid, DateTime
    using System.Collections.Generic; // Required for List, HashSet
    using System.IO;
    using System.Linq; // Required for LINQ extensions like Any(), OrderBy(), OfType()
    using System.Reflection; // Required for getting assembly location
    using System.Text;
    using System.Windows.Forms; // For FolderBrowserDialog and OpenFileDialog

    // Defines a structure to hold data parsed from a CSV mapping entry
    public class CsvMappingEntry
    {
        public string RTSGuid { get; set; } // RTS Parameter GUID (target)
        public string RTSName { get; set; } // RTS Parameter Name (target)
        public string DataType { get; set; } // Expected Data Type (e.g., TEXT, NUMBER)
        public string ClientGuid { get; set; } // Client Parameter GUID (source)
        public string ClientParameterName { get; set; } // Client Parameter Name (source)
        public string CsvNotes { get; set; } // Value from the 'Notes' column in CSV
    }

    // Defines a structure to hold data for the System Notes export report
    public class SystemNotesReportEntry
    {
        public string ElementId { get; set; }
        public string ElementName { get; set; }
        public string ElementCategory { get; set; }
        public string SystemNotesValue { get; set; }
        public string Status { get; set; } // e.g., "Updated", "Skipped", "Error"
    }

    // RTS_MapCablesClass
    //
    // This class contains the core logic for the RTS_MapCables Revit add-in.
    // It will handle the parsing of the CSV mapping file, iterating through
    // Revit elements, and updating parameters based on the defined mappings.
    //
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class RTS_MapCablesClass : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Show the WPF user selection window
            UserSelectionWindow selectionWindow = new UserSelectionWindow();
            bool? dialogResult = selectionWindow.ShowDialog();

            if (dialogResult == true && selectionWindow.GenerateTemplate)
            {
                // User wants to generate the template file
                // Pass commandData.Application.Application to the template generation method
                return GenerateCsvTemplate(commandData.Application.Application);
            }
            else if (dialogResult == false)
            {
                // User chose not to generate the template, proceed with mapping/import
                // Call the method to process the CSV and then perform mapping
                return ProcessAndMapParameters(commandData.Application.ActiveUIDocument);
            }
            else
            {
                // Dialog was cancelled or closed without a clear choice
                message = "RTS_MapCables operation cancelled by user.";
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Generates a CSV template file with predefined headers and sample data,
        /// now including 'Client GUID', 'Notes', and 'System Notes' columns.
        /// Updated "GUID" to "RTS GUID" and "NAME" to "RTS NAME".
        /// </summary>
        /// <param name="application">The Revit ApplicationServices.Application instance.</param>
        /// <returns>Result.Succeeded if the file is generated, Result.Failed otherwise.</returns>
        private Result GenerateCsvTemplate(Autodesk.Revit.ApplicationServices.Application application)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select a folder to save the RTS_Mapping_Template.csv";
                    // Default to the location of the currently executing assembly (your add-in's DLL)
                    folderDialog.SelectedPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        string folderPath = folderDialog.SelectedPath;
                        string filePath = Path.Combine(folderPath, "RTS_Mapping_Template.csv");

                        StringBuilder csvContent = new StringBuilder();
                        // Updated header to include 'Client GUID', 'Notes', and 'System Notes'
                        // Changed "GUID" to "RTS GUID" and "NAME" to "RTS NAME"
                        csvContent.AppendLine("RTS GUID,RTS NAME,DATATYPE,CLIENT GUID,CLIENT PARAMETER NAME,NOTES,SYSTEM NOTES"); // All uppercase
                        // Updated names from RT_ to RTS_ as requested (already done in previous iteration)
                        csvContent.AppendLine("a6f087c7-cecc-4335-831b-249cb9398abc,RTS_Tray Occupancy,TEXT,,,,");
                        csvContent.AppendLine("51d670fa-0338-42e7-ac9e-f2c44a56ffcc,RTS_Cables Weight,TEXT,,,,");
                        csvContent.AppendLine("5ed6b64c-af5c-4425-ab69-85a7fa5fdffe,RTS_Tray Min Size,TEXT,,,,");
                        csvContent.AppendLine("3175a27e-d386-4567-bf10-2da1a9cbb73b,RTS_ID,TEXT,,,,");
                        csvContent.AppendLine("cf0d478e-1e98-4e83-ab80-6ee867f61798,RTS_Cable_01,TEXT,,,,");
                        csvContent.AppendLine("2551d308-44ed-405c-8aad-fb78624d086e,RTS_Cable_02,TEXT,,,,");
                        csvContent.AppendLine("c1dfc402-2101-4e53-8f52-f6af64584a9f,RTS_Cable_03,TEXT,,,,");
                        csvContent.AppendLine("f297daa6-a9e0-4dd5-bda3-c628db7c28bd,RTS_Cable_04,TEXT,,,,");
                        csvContent.AppendLine("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab,RTS_Cable_05,TEXT,,,,");
                        csvContent.AppendLine("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9,RTS_Cable_06,TEXT,,,,");
                        csvContent.AppendLine("9bc78bce-0d39-4538-b507-7b98e8a13404,RTS_Cable_07,TEXT,,,,");
                        csvContent.AppendLine("e9d50153-a0e9-4685-bc92-d89f244f7e8e,RTS_Cable_08,TEXT,,,,");
                        csvContent.AppendLine("5713d65a-91df-4d2e-97bf-1c3a10ea5225,RTS_Cable_09,TEXT,,,,");
                        csvContent.AppendLine("64af3105-b2fd-44bc-9ad3-17264049ff62,RTS_Cable_10,TEXT,,,,");
                        csvContent.AppendLine("f3626002-0e62-4b75-93cc-35d0b11dfd67,RTS_Cable_11,TEXT,,,,");
                        csvContent.AppendLine("63dc0a2e-0770-4002-a859-a9d40a2ce023,RTS_Cable_12,TEXT,,,,");
                        csvContent.AppendLine("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2,RTS_Cable_13,TEXT,,,,");
                        csvContent.AppendLine("0e0572e5-c568-42b7-8730-a97433bd9b54,RTS_Cable_14,TEXT,,,,");
                        csvContent.AppendLine("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f,RTS_Cable_15,TEXT,,,,");
                        csvContent.AppendLine("f6d2af67-027e-4b9c-9def-336ebaa87336,RTS_Cable_16,TEXT,,,,");
                        csvContent.AppendLine("f6a4459d-46a1-44c0-8545-ee44e4778854,RTS_Cable_17,TEXT,,,,");
                        csvContent.AppendLine("0d66d2fa-f261-4daa-8041-9eadeefac49a,RTS_Cable_18,TEXT,,,,");
                        csvContent.AppendLine("af483914-c8d2-4ce6-be6e-ab81661e5bf1,RTS_Cable_19,TEXT,,,,");
                        csvContent.AppendLine("c8d2d2fc-c248-483f-8d52-e630eb730cd7,RTS_Cable_20,TEXT,,,,");
                        csvContent.AppendLine("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506,RTS_Cable_21,TEXT,,,,");
                        csvContent.AppendLine("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2,RTS_Cable_22,TEXT,,,,");
                        csvContent.AppendLine("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2,RTS_Cable_23,TEXT,,,,");
                        csvContent.AppendLine("7f745b2b-a537-42d9-8838-7a5521cc7d0c,RTS_Cable_24,TEXT,,,,");
                        csvContent.AppendLine("9a76c2dc-1022-4a54-ab66-5ca625b50365,RTS_Cable_25,TEXT,,,,");
                        csvContent.AppendLine("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e,RTS_Cable_26,TEXT,,,,");
                        csvContent.AppendLine("8ad24640-036b-44d2-af9c-b891f6e64271,RTS_Cable_27,TEXT,,,,");
                        csvContent.AppendLine("c046c4d7-e1fd-4cf7-a99f-14ae96b722be,RTS_Cable_28,TEXT,,,,");
                        csvContent.AppendLine("cdf00587-7e11-4af4-8e54-48586481cf22,RTS_Cable_29,TEXT,,,,");
                        csvContent.AppendLine("a92bb0f9-2781-4971-a3b1-9c47d62b947b,RTS_Cable_30,TEXT,,,,");

                        File.WriteAllText(filePath, csvContent.ToString());
                        TaskDialog.Show("RTS_MapCables", $"CSV template file generated successfully at:\n{filePath}");
                        return Result.Succeeded;
                    }
                    else
                    {
                        TaskDialog.Show("RTS_MapCables", "CSV template generation cancelled by user.");
                        return Result.Cancelled;
                    }
                }
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("RTS_MapCables Error", $"Failed to generate CSV template: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Prompts the user to select a CSV mapping file, parses it, and then
        /// proceeds to map parameters on relevant Revit elements.
        /// </summary>
        /// <param name="uidoc">The active Revit UI document.</param>
        /// <returns>Result.Succeeded if parameters are processed and mapped, Result.Failed or Result.Cancelled otherwise.</returns>
        private Result ProcessAndMapParameters(UIDocument uidoc)
        {
            if (uidoc == null)
            {
                TaskDialog.Show("RTS_MapCables Error", "No active Revit document found.");
                return Result.Failed;
            }

            Document doc = uidoc.Document;

            // 1. Prompt the user to import the CSV file
            string csvFilePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.Title = "Select the CSV mapping file";
                openFileDialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    csvFilePath = openFileDialog.FileName;
                }
                else
                {
                    TaskDialog.Show("RTS_MapCables", "CSV file selection cancelled.");
                    return Result.Cancelled;
                }
            }

            if (string.IsNullOrEmpty(csvFilePath))
            {
                return Result.Cancelled; // Should not happen if dialog was OK, but good for safety
            }

            List<CsvMappingEntry> csvMappingEntries = new List<CsvMappingEntry>();
            try
            {
                // 2. Parse the CSV file
                string[] lines = File.ReadAllLines(csvFilePath);

                // Skip header (first line)
                if (lines.Length > 0)
                {
                    for (int i = 1; i < lines.Length; i++)
                    {
                        string line = lines[i];
                        string[] parts = line.Split(',');

                        // Expecting at least 6 columns (RTS GUID, RTS NAME, DATATYPE, Client GUID, Client Parameter Name, Notes)
                        // System Notes column (index 6) exists in template but its value is generated, not read from CSV for mapping.
                        if (parts.Length >= 6)
                        {
                            string rtsGuid = parts[0].Trim();
                            string rtsName = parts[1].Trim();
                            string dataType = parts[2].Trim();
                            string clientGuid = parts[3].Trim();
                            string clientParameterName = parts[4].Trim();
                            string csvNotes = parts[5].Trim(); // Read the 'Notes' column content

                            // Only add if at least one of Client GUID or Client Parameter Name is provided for the primary mapping,
                            // or if a note is explicitly provided (indicating a desire to process this row).
                            if (!string.IsNullOrEmpty(clientGuid) || !string.IsNullOrEmpty(clientParameterName) || !string.IsNullOrEmpty(csvNotes))
                            {
                                csvMappingEntries.Add(new CsvMappingEntry
                                {
                                    RTSGuid = rtsGuid,
                                    RTSName = rtsName,
                                    DataType = dataType,
                                    ClientGuid = clientGuid,
                                    ClientParameterName = clientParameterName,
                                    CsvNotes = csvNotes // Store the notes content
                                });
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("RTS_MapCables Error", $"Failed to read or parse CSV file: {ex.Message}");
                return Result.Failed;
            }

            if (!csvMappingEntries.Any())
            {
                TaskDialog.Show("RTS_MapCables", "No valid mapping entries found in the selected CSV file (ensure 'Client GUID', 'Client Parameter Name', or 'Notes' column is populated).");
                return Result.Succeeded; // No parameters to map, but process was successful
            }

            // Now, perform the actual parameter mapping on elements
            return MapParametersOnElements(doc, csvMappingEntries);
        }

        /// <summary>
        /// Maps parameters on specified Revit elements (Cable Trays, Cable Tray Fittings, Conduits, Conduit Fittings)
        /// based on the provided CSV mapping entries. Values are transferred from Client parameters to RTS parameters,
        /// and 'Notes'/'System Notes' are also updated.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="mappingEntries">A list of mapping rules parsed from the CSV.</param>
        /// <returns>Result.Succeeded if mapping is completed, Result.Failed otherwise.</returns>
        private Result MapParametersOnElements(Document doc, List<CsvMappingEntry> mappingEntries)
        {
            // Collect categories for filtering elements
            List<BuiltInCategory> targetCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            // Use a Multi-category filter to collect relevant elements
            ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(targetCategories);

            // Filter for instances (not types)
            ElementClassFilter instanceFilter = new ElementClassFilter(typeof(FamilyInstance)); // Covers fittings
            ElementClassFilter conduitFilter = new ElementClassFilter(typeof(Conduit)); // Covers conduits
            ElementClassFilter cableTrayFilter = new ElementClassFilter(typeof(CableTray)); // Covers cable trays

            // Combine filters using logical OR filter to get all desired elements
            LogicalOrFilter elementFilter = new LogicalOrFilter(instanceFilter, new LogicalOrFilter(conduitFilter, cableTrayFilter));
            LogicalAndFilter finalFilter = new LogicalAndFilter(categoryFilter, elementFilter);


            // Collect all relevant elements in the document
            List<Element> elementsToProcess = new FilteredElementCollector(doc)
                .WherePasses(finalFilter)
                .WhereElementIsNotElementType() // Only instances
                .ToList();

            if (!elementsToProcess.Any())
            {
                TaskDialog.Show("RTS_MapCables", "No Cable Trays, Cable Tray Fittings, Conduits, or Conduit Fittings found in the project to map parameters.");
                return Result.Succeeded;
            }

            List<SystemNotesReportEntry> systemNotesReport = new List<SystemNotesReportEntry>();
            int parametersUpdatedCount = 0;
            string generatedReportFilePath = string.Empty; // To store the path of the generated CSV report

            // Start a transaction to modify the document
            using (Transaction t = new Transaction(doc, "Map RTS Parameters"))
            {
                try
                {
                    t.Start();

                    foreach (Element element in elementsToProcess)
                    {
                        string currentSystemNotesValue = "";
                        string systemNotesStatus = "Skipped (no action)"; // Default status

                        // Attempt to read current System Notes value for reporting
                        Parameter existingSystemNotesParam = GetParameterByNameOrGuid(element, "", "System Notes");
                        if (existingSystemNotesParam != null && existingSystemNotesParam.StorageType == StorageType.String)
                        {
                            currentSystemNotesValue = existingSystemNotesParam.AsString();
                        }


                        foreach (CsvMappingEntry entry in mappingEntries)
                        {
                            // --- Primary Client to RTS Parameter Mapping ---
                            // Source is now Client parameter, Target is now RTS parameter
                            Parameter sourceClientParam = GetParameterByNameOrGuid(element, entry.ClientGuid, entry.ClientParameterName);

                            if (sourceClientParam == null)
                            {
                                // Removed report.AppendLine for each individual skipped parameter,
                                // as per the request to exclude detailed reporting to user.
                                continue; // Skip to next mapping entry for this element
                            }
                            else
                            {
                                Parameter targetRTSParam = GetParameterByNameOrGuid(element, entry.RTSGuid, entry.RTSName);

                                if (targetRTSParam == null)
                                {
                                    // Removed report.AppendLine for each individual skipped parameter.
                                    continue; // Skip to next mapping entry for this element
                                }
                                else if (targetRTSParam.IsReadOnly) // Ensure target is not read-only
                                {
                                    // Removed report.AppendLine for each individual skipped parameter.
                                    continue;
                                }
                                else
                                {
                                    // 3. Transfer the value from source (Client) to target (RTS) for primary mapping
                                    try
                                    {
                                        bool valueSet = false;
                                        switch (sourceClientParam.StorageType)
                                        {
                                            case StorageType.String:
                                                targetRTSParam.Set(sourceClientParam.AsString());
                                                valueSet = true;
                                                break;
                                            case StorageType.Double:
                                                targetRTSParam.Set(sourceClientParam.AsDouble());
                                                valueSet = true;
                                                break;
                                            case StorageType.Integer:
                                                targetRTSParam.Set(sourceClientParam.AsInteger());
                                                valueSet = true;
                                                break;
                                            case StorageType.ElementId:
                                                targetRTSParam.Set(sourceClientParam.AsElementId());
                                                valueSet = true;
                                                break;
                                            default:
                                                targetRTSParam.Set(sourceClientParam.AsValueString());
                                                valueSet = true;
                                                break;
                                        }

                                        if (valueSet)
                                        {
                                            parametersUpdatedCount++;
                                            // Removed report.AppendLine for successful mapping.
                                        }
                                        else
                                        {
                                            // Removed report.AppendLine for failed setting due to StorageType.
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        // Removed report.AppendLine for error setting parameter.
                                    }
                                }
                            }

                            // --- Handle 'Notes' parameter ---
                            if (!string.IsNullOrEmpty(entry.CsvNotes))
                            {
                                Parameter notesParam = GetParameterByNameOrGuid(element, "", "Notes"); // Assuming "Notes" is the target parameter name by default
                                if (notesParam != null && notesParam.StorageType == StorageType.String && !notesParam.IsReadOnly)
                                {
                                    try
                                    {
                                        notesParam.Set(entry.CsvNotes);
                                        parametersUpdatedCount++;
                                    }
                                    catch (Exception)
                                    {
                                        // Error during setting Notes parameter
                                    }
                                }
                            }

                            // --- Handle 'System Notes' parameter ---
                            Parameter systemNotesParam = GetParameterByNameOrGuid(element, "", "System Notes");
                            if (systemNotesParam != null && systemNotesParam.StorageType == StorageType.String && !systemNotesParam.IsReadOnly)
                            {
                                try
                                {
                                    string statusMessage = $"Mapped by RTS_MapCables on {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}";
                                    systemNotesParam.Set(statusMessage);
                                    currentSystemNotesValue = statusMessage; // Update value for report
                                    systemNotesStatus = "Updated";
                                    parametersUpdatedCount++;
                                }
                                catch (Exception)
                                {
                                    currentSystemNotesValue = "Error setting System Notes";
                                    systemNotesStatus = "Error updating";
                                }
                            }
                            else
                            {
                                currentSystemNotesValue = "System Notes parameter not found or read-only";
                                systemNotesStatus = "Skipped (param missing/read-only)";
                            }
                        }

                        // Add entry to report list after processing all mappings for the element
                        systemNotesReport.Add(new SystemNotesReportEntry
                        {
                            ElementId = element.Id.ToString(),
                            ElementName = element.Name,
                            ElementCategory = element.Category?.Name ?? "N/A",
                            SystemNotesValue = currentSystemNotesValue,
                            Status = systemNotesStatus
                        });
                    }

                    t.Commit(); // Commit the transaction if all operations are successful

                    // Export the System Notes report to CSV and get the file path
                    generatedReportFilePath = ExportSystemNotesReportToCsv(systemNotesReport);

                    // Show a simple TaskDialog confirming completion and report location
                    TaskDialog.Show("RTS_MapCables",
                                    $"Mapping complete.\n\n" +
                                    $"Processed {elementsToProcess.Count} elements.\n" +
                                    $"Updated {parametersUpdatedCount} parameters.\n\n" +
                                    $"System Notes Report saved to:\n{generatedReportFilePath}");

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    // If any error occurs, rollback the transaction
                    t.RollBack();
                    TaskDialog.Show("RTS_MapCables Error", $"Mapping Failed: {ex.Message}\nTransaction rolled back. No changes were applied to the model.");
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Exports the collected System Notes report data to a CSV file.
        /// </summary>
        /// <param name="reportEntries">List of SystemNotesReportEntry objects.</param>
        /// <returns>The full path to the generated CSV file, or an empty string if cancelled/failed.</returns>
        private string ExportSystemNotesReportToCsv(List<SystemNotesReportEntry> reportEntries)
        {
            string filePath = string.Empty;
            try
            {
                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                    saveFileDialog.Title = "Save System Notes Report CSV";
                    saveFileDialog.FileName = $"RTS_SystemNotes_Report_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.csv";
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // Default to My Documents

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = saveFileDialog.FileName;
                        StringBuilder csvContent = new StringBuilder();

                        // Add CSV header
                        csvContent.AppendLine("ELEMENT ID,ELEMENT NAME,CATEGORY,SYSTEM NOTES VALUE,STATUS"); // All uppercase

                        // Add data rows
                        foreach (var entry in reportEntries)
                        {
                            // Enclose values in quotes to handle commas within strings
                            csvContent.AppendLine($"\"{entry.ElementId}\",\"{entry.ElementName.Replace("\"", "\"\"")}\",\"{entry.ElementCategory.Replace("\"", "\"\"")}\",\"{entry.SystemNotesValue.Replace("\"", "\"\"")}\",\"{entry.Status}\"");
                        }

                        File.WriteAllText(filePath, csvContent.ToString());
                        // No TaskDialog here, it will be handled by the calling method.
                    }
                    else
                    {
                        // TaskDialog.Show("RTS_MapCables", "System Notes Report export cancelled by user."); // Removed as per request
                        return string.Empty; // Return empty string if cancelled
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RTS_MapCables Error", $"Failed to export System Notes Report: {ex.Message}");
                return string.Empty; // Return empty string on error
            }
            return filePath;
        }

        /// <summary>
        /// Helper method to retrieve a Parameter from an Element, prioritizing GUID lookup
        /// for Shared Parameters, then falling back to name lookup for Project or Built-in Parameters.
        /// </summary>
        /// <param name="element">The Revit element to search within.</param>
        /// <param name="guidString">The string representation of the parameter's GUID.</param>
        /// <param name="nameString">The name of the parameter.</param>
        /// <returns>The found Parameter object, or null if not found.</returns>
        private Parameter GetParameterByNameOrGuid(Element element, string guidString, string nameString)
        {
            Parameter param = null;

            // 1. Try to find by GUID first (primarily for Shared Parameters)
            if (!string.IsNullOrEmpty(guidString) && Guid.TryParse(guidString, out Guid parameterGuid) && parameterGuid != Guid.Empty)
            {
                param = element.get_Parameter(parameterGuid);
                if (param != null) return param;
            }

            // 2. If not found by GUID, try to find by Name (for Project Parameters or Built-in Parameters)
            if (!string.IsNullOrEmpty(nameString))
            {
                // Using LookupParameter for built-in or project parameters by name
                param = element.LookupParameter(nameString);
                if (param != null) return param;
            }

            return null; // Parameter not found by either method
        }
    }
}
