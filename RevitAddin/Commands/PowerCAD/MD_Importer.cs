// Filename: MD_Importer.cs
// Description: Revit external command to update Detail Item parameters from an Excel file.
// Reads a table named "TB_Submains" from an .xlsx file.
// Uses the "PC_SWB to" column as a key to find matching Detail Items.
// Updates the "PC_SWB Load" parameter on those Detail Items.
// Revit 2024 / Visual Studio 2022 / ClosedXML

using Autodesk.Revit.ApplicationServices; // Revit's Application class is here
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms; // This namespace also has an 'Application' class
using System.IO;
using RTS.Commands.Support;
using ClosedXML.Excel;    // For Excel manipulation

namespace RTS.Commands.PowerCAD
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MD_ImporterClass : IExternalCommand
    {
        // Define Shared Parameter GUIDs
        // Parameter on Detail Item to use as a key (matches Excel's "PC_SWB to" column)
        private readonly Guid _pcSwbToGuid = new Guid("e142b0ed-d084-447a-991b-d9a3a3f67a8d"); // PC_SWB To (TEXT)
        // Parameter on Detail Item to update (value comes from Excel's "PC_SWB Load" column)
        private readonly Guid _pcSwbLoadGuid = new Guid("60f670e1-7d54-4ffc-b0f5-ed62c08d3b90"); // PC_SWB Load (TEXT)

        // Excel Table and Column Names
        private const string ExcelTableName = "TB_Submains";
        private const string ExcelKeyColumnName = "PC_SWB to";
        private const string ExcelValueColumnName = "PC_SWB Load";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            // Fully qualify the Application type to resolve ambiguity
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application; // <<<< CORRECTED HERE
            Document doc = uidoc.Document;

            // 1. Prompt user to select the Excel file
            string filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                openFileDialog.Title = $"Select Excel File with Table '{ExcelTableName}'";
                openFileDialog.RestoreDirectory = true; // Remembers the last opened directory

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    message = "File selection cancelled.";
                    return Result.Cancelled;
                }
                filePath = openFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                message = "Invalid file path or file does not exist.";
                CustomTaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // 2. Read data from the specified table in the Excel file
            // Using a case-insensitive dictionary for Excel keys for robustness
            Dictionary<string, string> excelDataMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (XLWorkbook workbook = new XLWorkbook(filePath))
                {
                    IXLTable table = null;
                    try
                    {
                        // Attempt to get the table by name
                        table = workbook.Table(ExcelTableName);
                    }
                    catch (ArgumentException) // ClosedXML throws ArgumentException if table not found
                    {
                        // More robust check if table exists by iterating through all tables in all worksheets
                        foreach (var worksheet in workbook.Worksheets)
                        {
                            table = worksheet.Tables.FirstOrDefault(t => t.Name.Equals(ExcelTableName, StringComparison.OrdinalIgnoreCase));
                            if (table != null) break;
                        }
                    }
                    catch (Exception ex) // Catch other potential errors like corrupted file
                    {
                        message = $"Error trying to access table '{ExcelTableName}'. The file might be corrupted or the table name contains invalid characters. Details: {ex.Message}";
                        CustomTaskDialog.Show("Excel Error", message);
                        return Result.Failed;
                    }


                    if (table == null)
                    {
                        message = $"The Excel Table named '{ExcelTableName}' was not found in the selected file '{Path.GetFileName(filePath)}'.\n" +
                                  $"Please ensure an Excel Table with this exact name exists within one of the worksheets.";
                        CustomTaskDialog.Show("Table Not Found", message);
                        return Result.Failed;
                    }

                    // Verify that the required columns exist in the table
                    if (!table.Fields.Any(f => f.Name.Equals(ExcelKeyColumnName, StringComparison.OrdinalIgnoreCase)))
                    {
                        message = $"The key column '{ExcelKeyColumnName}' was not found in the table '{ExcelTableName}'.\n" +
                                  "Please check the column header spelling and case.";
                        CustomTaskDialog.Show("Column Not Found", message);
                        return Result.Failed;
                    }

                    if (!table.Fields.Any(f => f.Name.Equals(ExcelValueColumnName, StringComparison.OrdinalIgnoreCase)))
                    {
                        message = $"The value column '{ExcelValueColumnName}' was not found in the table '{ExcelTableName}'.\n" +
                                  "Please check the column header spelling and case.";
                        CustomTaskDialog.Show("Column Not Found", message);
                        return Result.Failed;
                    }

                    // Populate the dictionary with data from the table
                    foreach (IXLTableRow row in table.DataRange.Rows())
                    {
                        // GetString() can return null if cell is empty. Trim() handles leading/trailing spaces.
                        string key = row.Field(ExcelKeyColumnName).GetString()?.Trim();
                        string value = row.Field(ExcelValueColumnName).GetString()?.Trim();

                        // Only add if the key is not null or whitespace
                        if (!string.IsNullOrWhiteSpace(key))
                        {
                            if (!excelDataMap.ContainsKey(key))
                            {
                                excelDataMap.Add(key, value ?? string.Empty); // Store empty string if value is null
                            }
                            else
                            {
                                // Log a warning if duplicate keys are found in Excel.
                                // The first one encountered will be used.
                                System.Diagnostics.Debug.WriteLine($"Warning: Duplicate key '{key}' found in Excel table '{ExcelTableName}'. Using the first occurrence.");
                            }
                        }
                    }
                }

                if (excelDataMap.Count == 0)
                {
                    message = $"No data was extracted from the table '{ExcelTableName}' (or its key column '{ExcelKeyColumnName}' was empty or contained only whitespace).\n" +
                              "Please check the Excel file content.";
                    CustomTaskDialog.Show("No Data Extracted", message);
                    return Result.Succeeded; // Operation completed, but no data to process.
                }
            }
            catch (Exception ex)
            {
                message = $"An error occurred while reading the Excel file: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
                CustomTaskDialog.Show("Excel Read Error", message);
                return Result.Failed;
            }

            // 3. Get all Detail Items in the current Revit document
            // Using a more specific filter for Detail Components
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> detailItems = collector.OfCategory(BuiltInCategory.OST_DetailComponents)
                                                        .WhereElementIsNotElementType() // Exclude family types
                                                        .ToElements();

            if (!detailItems.Any())
            {
                CustomTaskDialog.Show("Info", "No Detail Items found in the current Revit document.");
                return Result.Succeeded; // No elements to process.
            }

            // 4. Update parameters in a Revit transaction
            int updatedCount = 0;
            int skippedNoKeyParam = 0; // Detail item missing the 'PC_SWB To' parameter or it's blank
            int skippedNoMatch = 0;    // 'PC_SWB To' value from Revit not found in Excel data
            int skippedNoTargetParam = 0; // Detail item missing the 'PC_SWB Load' parameter
            int skippedNotText = 0;       // 'PC_SWB Load' parameter is not of StorageType.String
            int skippedReadOnly = 0;      // 'PC_SWB Load' parameter is read-only


            using (Transaction trans = new Transaction(doc, "Update Detail Item PC_SWB Load from Excel"))
            {
                if (trans.Start() == TransactionStatus.Started)
                {
                    foreach (Element el in detailItems)
                    {
                        // Attempt to get the 'PC_SWB To' parameter using its GUID
                        Parameter pcSwbToParam = el.get_Parameter(_pcSwbToGuid);

                        // Check if the parameter exists and has a value
                        if (pcSwbToParam != null && pcSwbToParam.HasValue)
                        {
                            string revitKeyValue = pcSwbToParam.AsString()?.Trim();

                            // Skip if the key parameter on the Revit element is blank or whitespace
                            if (string.IsNullOrWhiteSpace(revitKeyValue))
                            {
                                // This case is subtly different from skippedNoKeyParam;
                                // the param exists but its value is unusable.
                                // Consider if this needs separate tracking or can be merged with skippedNoKeyParam.
                                continue;
                            }

                            // Try to find a match in the Excel data (case-insensitive key comparison)
                            if (excelDataMap.TryGetValue(revitKeyValue, out string excelLoadValue))
                            {
                                // Attempt to get the 'PC_SWB Load' parameter using its GUID
                                Parameter pcSwbLoadParam = el.get_Parameter(_pcSwbLoadGuid);
                                if (pcSwbLoadParam != null)
                                {
                                    // Check if the parameter is read-only
                                    if (pcSwbLoadParam.IsReadOnly)
                                    {
                                        skippedReadOnly++;
                                        System.Diagnostics.Debug.WriteLine($"Warning: Parameter 'PC_SWB Load' (GUID: {_pcSwbLoadGuid}) on element ID {el.Id} is read-only. Skipping update.");
                                        continue;
                                    }
                                    // Ensure the parameter is of TEXT type (StorageType.String)
                                    if (pcSwbLoadParam.StorageType == StorageType.String)
                                    {
                                        try
                                        {
                                            // Set the parameter value
                                            pcSwbLoadParam.Set(excelLoadValue);
                                            updatedCount++;
                                        }
                                        catch (Exception ex)
                                        {
                                            // Log any errors during the Set operation
                                            System.Diagnostics.Debug.WriteLine($"Error setting parameter 'PC_SWB Load' for Element ID {el.Id}: {ex.Message}");
                                            // Optionally, add to a list of elements that failed to update for more detailed reporting
                                        }
                                    }
                                    else
                                    {
                                        skippedNotText++;
                                        System.Diagnostics.Debug.WriteLine($"Warning: Parameter 'PC_SWB Load' (GUID: {_pcSwbLoadGuid}) on element ID {el.Id} is not a Text type (Actual StorageType: {pcSwbLoadParam.StorageType}). Skipping update.");
                                    }
                                }
                                else
                                {
                                    skippedNoTargetParam++; // 'PC_SWB Load' parameter not found on the element
                                }
                            }
                            else
                            {
                                skippedNoMatch++; // Key from Revit element not found in Excel data
                            }
                        }
                        else
                        {
                            skippedNoKeyParam++; // 'PC_SWB To' parameter not found on the element or has no value
                        }
                    }

                    // Attempt to commit the transaction
                    if (trans.Commit() == TransactionStatus.Committed)
                    {
                        string summary = $"Update process completed.\n\n" +
                                         $"Successfully updated parameters for {updatedCount} Detail Item(s).\n" +
                                         $"Skipped (no matching 'PC_SWB to' value in Excel): {skippedNoMatch}\n" +
                                         $"Skipped (Detail Item missing 'PC_SWB To' param or value is blank): {skippedNoKeyParam}\n" + // Clarified this counter
                                         $"Skipped (Detail Item missing 'PC_SWB Load' param): {skippedNoTargetParam}\n" +
                                         $"Skipped ('PC_SWB Load' param is read-only): {skippedReadOnly}\n" +
                                         $"Skipped ('PC_SWB Load' param is not Text type): {skippedNotText}";
                        CustomTaskDialog.Show("Update Complete", summary);
                        return Result.Succeeded;
                    }
                    else
                    {
                        // If commit fails, roll back changes
                        trans.RollBack();
                        message = "Transaction could not be committed. Changes have been rolled back.";
                        CustomTaskDialog.Show("Transaction Error", message);
                        return Result.Failed;
                    }
                }
                else
                {
                    message = "Failed to start Revit transaction.";
                    CustomTaskDialog.Show("Transaction Error", message);
                    return Result.Failed;
                }
            }
        }
    }
}
