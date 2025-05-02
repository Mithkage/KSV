using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Globalization; // Added for robust number parsing

namespace PC_Exporter
{
    [Transaction(TransactionMode.Manual)]
    public class PowerCAD_ExportSLD : IExternalCommand
    {
        // Helper class to store data for each relevant detail item
        private class DetailItemData
        {
            public string OriginalCableReference { get; set; } // Raw value from Revit parameter
            public string FinalCableReference { get; set; }    // The definitive reference used for the group
            public string SWBFrom { get; set; }
            public string SWBTo { get; set; }
            public string SWBType { get; set; } // 'S' for Bus, 'T' for TOB, potentially others
            public string SWBLoad { get; set; }
            public string SWBLoadScope { get; set; }
            public string SWBPF { get; set; }
            public string CableLength { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string CableSizeActive { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string CableSizeNeutral { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string CableSizeEarthing { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string ActiveConductorMaterial { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string NumPhases { get; set; } // ** Populated based on Type Name containing "1-Phase", Potentially nulled if SWBType = 'S' **
            public string CableType { get; set; } // ** FINAL VALUE SET IN STEP 5b based on SWBType, Length/Rating **
            public string CableInsulation { get; set; } // ** UPDATED based on PC_From containing "SAFETY", Potentially nulled if SWBType = 'S' **
            public string InstallationMethod { get; set; } // ** UPDATED to always be "PT", Potentially nulled if SWBType = 'S' **
            public string CableAdditionalDerating { get; set; } // ** Potentially nulled if SWBType = 'S' **
            public string SwitchgearTripUnitType { get; set; } // ** FINAL VALUE SET IN STEP 5b based on NumPhases and SWBType **
            public string SwitchgearManufacturer { get; set; } // ** UPDATED to always be "NAW Controls - LS Susol" **
            public string BusType { get; set; } // Set to "Bus Bar"
            public string BusChassisRating { get; set; }
            public string UpstreamDiversity { get; set; } // ** UPDATED to always be "STD" **
            public string IsolatorType { get; set; }
            public string IsolatorRating { get; set; }
            public string ProtectiveDeviceRating { get; set; } // Used for Cable Type logic
            public string ProtectiveDeviceManufacturer { get; set; }
            public string ProtectiveDeviceType { get; set; }
            public string ProtectiveDeviceModel { get; set; }
            public string ProtectiveDeviceOCRTripUnit { get; set; }
            public string ProtectiveDeviceTripSetting { get; set; }
        }

        // Helper function to escape characters for CSV format
        private string EscapeCsvField(string field)
        {
            if (field == null) return "";
            // Check if the field contains characters that require quoting in CSV
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                // Escape double quotes by doubling them
                string escapedField = field.Replace("\"", "\"\"");
                // Enclose the entire field in double quotes
                return $"\"{escapedField}\"";
            }
            else
            {
                // No special characters, return the field as is
                return field;
            }
        }

        // --- Member variables for Pre-order Traversal ---
        private List<DetailItemData> _preOrderSortedData;
        private HashSet<string> _visitedNodesDuringTraversal;
        private Dictionary<string, List<DetailItemData>> _itemsOriginatingFrom;


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            // Filter for Detail Items in the active document
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            // Store the detail items for potential later lookup (needed for Cable Type logic if not 'T')
            List<Element> detailItems = collector.OfCategory(BuiltInCategory.OST_DetailComponents)
                                                 .WhereElementIsNotElementType() // Get instances, not types
                                                 .ToList();

            // Check if any detail items were found
            if (detailItems.Count == 0)
            {
                TaskDialog.Show("Info", "No Detail Items found in the current view/document.");
                return Result.Succeeded; // Nothing to process
            }

            // --- Step 1: Data Collection & Pre-processing ---
            var groupedData = new Dictionary<string, List<DetailItemData>>(); // Group items by their 'SWB To' value
            var allNodes = new HashSet<string>(); // Keep track of all unique 'SWB To' and 'SWB From' values

            foreach (Element detailItem in detailItems)
            {
                // Get relevant parameters from the detail item
                Parameter pcPowerCADParam = detailItem.LookupParameter("PC_PowerCAD");
                Parameter pcSWBToParam = detailItem.LookupParameter("PC_SWB To");
                Parameter pcSWBFromParam = detailItem.LookupParameter("PC_From");
                Parameter pcCableLengthParam = detailItem.LookupParameter("PC_Cable Length"); // Get Cable Length parameter
                Parameter pcProtectiveDeviceRatingParam = detailItem.LookupParameter("PC_Protective Device Trip Setting (A)"); // Get Rating parameter


                // Process only if PC_PowerCAD is 1 (true) and PC_SWB To has a value
                if (pcPowerCADParam != null && pcPowerCADParam.AsInteger() == 1 &&
                    pcSWBToParam != null && !string.IsNullOrWhiteSpace(pcSWBToParam.AsString()))
                {
                    string swbToValue = pcSWBToParam.AsString();
                    string swbFromValue = pcSWBFromParam?.AsString() ?? ""; // Use null-conditional and coalesce for safety
                    // Default 'SWB From' to "SOURCE" if it's empty
                    if (string.IsNullOrWhiteSpace(swbFromValue))
                    {
                        swbFromValue = "SOURCE";
                    }

                    Parameter pcCableReferenceParam = detailItem.LookupParameter("PC_Cable Reference");
                    string cableReferenceValue = pcCableReferenceParam?.AsString() ?? "";

                    // Add the 'From' and 'To' nodes to the set for graph building
                    allNodes.Add(swbToValue);
                    allNodes.Add(swbFromValue);

                    // Initialize variables to store type-based information
                    string swbTypeValue = "";
                    string cableTypeValue = ""; // Will be potentially overwritten later
                    string cableInsulationValue = "";
                    string numPhasesValue = ""; // Initialize NumPhases
                    string switchgearTripUnitTypeValue = ""; // Initialize - will be set definitively in Step 5b
                    string cableLengthValue = pcCableLengthParam?.AsString() ?? ""; // Store cable length
                    string protectiveDeviceRatingValue = pcProtectiveDeviceRatingParam?.AsString() ?? ""; // Store rating

                    // Get the Element Type of the detail item
                    ElementId elementTypeId = detailItem.GetTypeId();
                    if (elementTypeId != null && elementTypeId != ElementId.InvalidElementId)
                    {
                        ElementType elementType = doc.GetElement(elementTypeId) as ElementType;
                        if (elementType != null)
                        {
                            string typeName = elementType.Name;
                            if (!string.IsNullOrEmpty(typeName))
                            {
                                // Determine SWB Type based on Type Name
                                if (typeName.IndexOf("BUS", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    swbTypeValue = "S"; // Switchboard/Bus
                                }
                                else if (typeName.IndexOf("TOB", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    swbTypeValue = "T"; // Tap-Off Box
                                }

                                // Check if Type Name contains "1-Phase" ONLY to set NumPhases
                                if (typeName.IndexOf("1-Phase", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    numPhasesValue = "R"; // Set # of Phases
                                }
                            }
                        }
                    }

                    // Check PC_From for "SAFETY" to set Cable Insulation
                    if (swbFromValue.IndexOf("SAFETY", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cableInsulationValue = "X-HF-110";
                    }

                    // Create a data object for the current detail item
                    var data = new DetailItemData
                    {
                        OriginalCableReference = cableReferenceValue,
                        SWBFrom = swbFromValue,
                        SWBTo = swbToValue,
                        SWBType = swbTypeValue,
                        SWBLoad = protectiveDeviceRatingValue,
                        SWBLoadScope = "Local",
                        SWBPF = "1",
                        CableLength = cableLengthValue,
                        CableSizeActive = "",
                        CableSizeNeutral = "",
                        CableSizeEarthing = "",
                        ActiveConductorMaterial = "",
                        NumPhases = numPhasesValue, // Set based on 1-Phase check in type name
                        CableType = cableTypeValue, // Initial value "", will be set in Step 5b
                        CableInsulation = cableInsulationValue,
                        InstallationMethod = "PT",
                        CableAdditionalDerating = "",
                        SwitchgearTripUnitType = switchgearTripUnitTypeValue, // Initial value "", will be set in Step 5b
                        SwitchgearManufacturer = "NAW Controls - LS Susol",
                        BusType = "Bus Bar",
                        BusChassisRating = detailItem.LookupParameter("PC_Bus/Chassis Rating (A)")?.AsString() ?? "",
                        UpstreamDiversity = "STD",
                        IsolatorType = "",
                        IsolatorRating = "",
                        ProtectiveDeviceRating = protectiveDeviceRatingValue,
                        ProtectiveDeviceManufacturer = "",
                        ProtectiveDeviceType = "",
                        ProtectiveDeviceModel = "",
                        ProtectiveDeviceOCRTripUnit = "",
                        ProtectiveDeviceTripSetting = protectiveDeviceRatingValue
                    };

                    // Add the data object to the dictionary, grouped by SWB To value
                    if (!groupedData.ContainsKey(swbToValue))
                    {
                        groupedData[swbToValue] = new List<DetailItemData>();
                    }
                    groupedData[swbToValue].Add(data);
                }
            }

            // Check if any valid items were found for export
            if (groupedData.Count == 0)
            {
                TaskDialog.Show("Info", "No Detail Items marked for export (PC_PowerCAD = 1) or with valid 'PC_SWB To' values were found.");
                return Result.Succeeded;
            }

            // --- Step 2: Determine Definitive Cable Reference & Create Flat List ---
            var definitiveCableReferencePerGroup = new Dictionary<string, string>();
            var allProcessedData = new List<DetailItemData>();
            foreach (var kvp in groupedData)
            {
                string swbTo = kvp.Key;
                List<DetailItemData> itemsInGroup = kvp.Value;
                string foundReference = itemsInGroup.FirstOrDefault(item => !string.IsNullOrWhiteSpace(item.OriginalCableReference))?.OriginalCableReference ?? "";
                definitiveCableReferencePerGroup[swbTo] = foundReference;
                foreach (var itemData_ref in itemsInGroup)
                {
                    itemData_ref.FinalCableReference = foundReference;
                    allProcessedData.Add(itemData_ref);
                }
            }

            // --- Step 2b: Initial Sort by Cable Reference ---
            allProcessedData = allProcessedData
                .OrderBy(d => d.FinalCableReference ?? string.Empty)
                .ToList();

            // --- Step 3: Build Dependency Graph ---
            var inDegree = new Dictionary<string, int>();
            foreach(string node in allNodes)
            {
                if (!inDegree.ContainsKey(node)) inDegree[node] = 0;
            }
            _itemsOriginatingFrom = allProcessedData
                .GroupBy(d => d.SWBFrom)
                .ToDictionary(g => g.Key, g => g.ToList());
            foreach (var item in allProcessedData)
            {
                if (inDegree.ContainsKey(item.SWBTo) && (item.SWBFrom != "SOURCE" || allNodes.Contains("SOURCE")))
                {
                    inDegree[item.SWBTo]++;
                }
            }

            // --- Step 4: Perform Pre-order Traversal Sort ---
            _preOrderSortedData = new List<DetailItemData>();
            _visitedNodesDuringTraversal = new HashSet<string>();
            var startingNodes = inDegree.Where(kvp => kvp.Value == 0).Select(kvp => kvp.Key).ToList();
            if (allNodes.Contains("SOURCE") && !startingNodes.Contains("SOURCE") && !allProcessedData.Any(d => d.SWBTo == "SOURCE"))
            {
                startingNodes.Add("SOURCE");
            }
            startingNodes.Sort();
            foreach (string startNode in startingNodes)
            {
                PreOrderVisit(startNode);
            }

            // --- Step 5: Merge Data for Unique SWB To ---
            var mergedUniqueDataList = new List<DetailItemData>();
            var lookupMap = new Dictionary<string, DetailItemData>();
            foreach(var newItemData in _preOrderSortedData)
            {
                if (lookupMap.TryGetValue(newItemData.SWBTo, out DetailItemData existingItemData))
                {
                    MergeDetailItemData(existingItemData, newItemData);
                }
                else
                {
                    mergedUniqueDataList.Add(newItemData);
                    lookupMap.Add(newItemData.SWBTo, newItemData);
                }
            }

            // --- Step 5b: Apply Final Logic Rules ---
            // Iterate through the merged list and apply final rules
            foreach (var itemData in mergedUniqueDataList)
            {
                // --- Nullify Cable Params for 'S' type ---
                if (itemData.SWBType == "S")
                {
                    // Nullify cable-related parameters for Bus type
                    itemData.CableLength = "";
                    itemData.CableSizeActive = "";
                    itemData.CableSizeNeutral = "";
                    itemData.CableSizeEarthing = "";
                    itemData.ActiveConductorMaterial = "";
                    itemData.NumPhases = ""; // Also nullify NumPhases for Bus type
                    itemData.CableType = ""; // Set Cable Type to blank for S type
                    itemData.CableInsulation = "";
                    itemData.InstallationMethod = "";
                    itemData.CableAdditionalDerating = "";
                }
                else // SWB Type is not "S"
                {
                    // --- MODIFICATION: Set Cable Type based on SWB Type or Length/Rating ---
                    if (itemData.SWBType == "T") // If it's a Tap-Off Box type
                    {
                        itemData.CableType = "SDI";
                    }
                    else // If not 'S' and not 'T', apply length/rating logic
                    {
                        bool lengthParsed = double.TryParse(itemData.CableLength, NumberStyles.Any, CultureInfo.InvariantCulture, out double cableLength);
                        bool ratingParsed = double.TryParse(itemData.ProtectiveDeviceRating, NumberStyles.Any, CultureInfo.InvariantCulture, out double deviceRating);
                        if (lengthParsed && ratingParsed && cableLength >= 100.0 && deviceRating >= 160.0)
                        {
                            itemData.CableType = "SDI";
                        }
                        else
                        {
                            itemData.CableType = "Multi"; // Default
                        }
                    }
                    // NumPhases retains its value ("R" or "") set during Step 1/Merging for non-'S' types
                }

                // --- Set Switchgear Trip Unit Type (Final Override Logic) ---
                // 1. Default: Set to Electronic initially.
                itemData.SwitchgearTripUnitType = "Electronic";

                // 2. Override if NumPhases is "R"
                if (itemData.NumPhases == "R")
                {
                    itemData.SwitchgearTripUnitType = "Thermal Magnetic";
                }

                // 3. Final Override: If SWBType is "S", it MUST be blank.
                if (itemData.SWBType == "S")
                {
                    itemData.SwitchgearTripUnitType = "";
                }
                // --- End Switchgear Trip Unit Type Logic ---

            } // End foreach loop for final logic


            // --- Step 6: Write Output Files ---
            string baseFileName = "PowerCAD_Export_SLD_Data"; // Base name for output files
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // Get user's desktop path
            string csvFilePath = Path.Combine(desktopPath, baseFileName + ".csv");
            string xlsxFilePath = Path.Combine(desktopPath, baseFileName + ".xlsx");

            bool csvSuccess = false;
            bool xlsxSuccess = false;
            string csvError = null;
            string xlsxError = null;

            // Define the headers for the output files
            string[] headers = {
                "Cable Reference", "SWB From", "SWB To", "SWB Type", "SWB Load",
                "SWB Load Scope", "SWB PF", "Cable Length", "Cable Size - Active conductors",
                "Cable Size - Neutral conductors", "Cable Size - Earthing conductor",
                "Active Conductor material", "# of Phases", "Cable Type", "Cable Insulation",
                "Installation Method", "Cable Additional De-rating", "Switchgear Trip Unit Type",
                "Switchgear Manufacturer", "Bus Type", "Bus/Chassis Rating (A)", "Upstream Diversity",
                "Isolator Type", "Isolator Rating (A)", "Protective Device Rating (A)",
                "Protective Device Manufacturer", "Protective Device Type", "Protective Device Model",
                "Protective Device OCR/Trip Unit", "Protective Device Trip Setting (A)"
            };

            // --- Write CSV File ---
            try
            {
                StringBuilder csvContent = new StringBuilder();
                // Add header row, escaping fields just in case headers contain commas etc.
                csvContent.AppendLine(string.Join(",", headers.Select(h => EscapeCsvField(h))));

                // Add data rows
                foreach (var itemData in mergedUniqueDataList) // Use the final processed list
                {
                    // Display empty string instead of "SOURCE" for SWB From in output
                    string displaySwbFrom = (itemData.SWBFrom == "SOURCE") ? "" : itemData.SWBFrom;

                    // Special handling for Cable Reference to force text format in Excel CSV import
                    string rawCableRef = itemData.FinalCableReference ?? "";
                    // Format as ="<value>" with internal quotes escaped
                    string excelFormattedCableRef = $"=\"{rawCableRef.Replace("\"", "\"\"")}\"";

                    // Create a list of fields for the current row, escaping each one
                    var fields = new List<string> {
                        excelFormattedCableRef, // Use the specially formatted value for Cable Reference
                        EscapeCsvField(displaySwbFrom),
                        EscapeCsvField(itemData.SWBTo),
                        EscapeCsvField(itemData.SWBType),
                        EscapeCsvField(itemData.SWBLoad),
                        EscapeCsvField(itemData.SWBLoadScope),
                        EscapeCsvField(itemData.SWBPF),
                        EscapeCsvField(itemData.CableLength),
                        EscapeCsvField(itemData.CableSizeActive),
                        EscapeCsvField(itemData.CableSizeNeutral),
                        EscapeCsvField(itemData.CableSizeEarthing),
                        EscapeCsvField(itemData.ActiveConductorMaterial),
                        EscapeCsvField(itemData.NumPhases), // Final value after Step 5b
                        EscapeCsvField(itemData.CableType), // Final value after Step 5b
                        EscapeCsvField(itemData.CableInsulation),
                        EscapeCsvField(itemData.InstallationMethod),
                        EscapeCsvField(itemData.CableAdditionalDerating),
                        EscapeCsvField(itemData.SwitchgearTripUnitType), // Final value after Step 5b
                        EscapeCsvField(itemData.SwitchgearManufacturer),
                        EscapeCsvField(itemData.BusType),
                        EscapeCsvField(itemData.BusChassisRating),
                        EscapeCsvField(itemData.UpstreamDiversity),
                        EscapeCsvField(itemData.IsolatorType),
                        EscapeCsvField(itemData.IsolatorRating),
                        EscapeCsvField(itemData.ProtectiveDeviceRating),
                        EscapeCsvField(itemData.ProtectiveDeviceManufacturer),
                        EscapeCsvField(itemData.ProtectiveDeviceType),
                        EscapeCsvField(itemData.ProtectiveDeviceModel),
                        EscapeCsvField(itemData.ProtectiveDeviceOCRTripUnit),
                        EscapeCsvField(itemData.ProtectiveDeviceTripSetting)
                    };
                    // Join the fields with commas and add the line to the CSV content
                    csvContent.AppendLine(string.Join(",", fields));
                }
                // Write the complete CSV content to the file (use UTF8 without BOM)
                File.WriteAllText(csvFilePath, csvContent.ToString(), new UTF8Encoding(false));
                csvSuccess = true;
            }
            catch (Exception ex)
            {
                csvError = $"Failed to write CSV file: {ex.GetType().Name} - {ex.Message}";
            }

            // --- Write XLSX File ---
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                excelApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1] = headers[i];
                }

                Excel.Range cableRefColumn = (Excel.Range)worksheet.Columns[1];
                cableRefColumn.NumberFormat = "@"; // Set Excel column format explicitly to Text

                int row = 2;
                foreach (var itemData in mergedUniqueDataList)
                {
                    string displaySwbFrom = (itemData.SWBFrom == "SOURCE") ? "" : itemData.SWBFrom;

                    worksheet.Cells[row, 1] = itemData.FinalCableReference; // Excel handles text format via NumberFormat
                    worksheet.Cells[row, 2] = displaySwbFrom;
                    worksheet.Cells[row, 3] = itemData.SWBTo;
                    worksheet.Cells[row, 4] = itemData.SWBType;
                    worksheet.Cells[row, 5] = itemData.SWBLoad;
                    worksheet.Cells[row, 6] = itemData.SWBLoadScope;
                    worksheet.Cells[row, 7] = itemData.SWBPF;
                    worksheet.Cells[row, 8] = itemData.CableLength;
                    worksheet.Cells[row, 9] = itemData.CableSizeActive;
                    worksheet.Cells[row, 10] = itemData.CableSizeNeutral;
                    worksheet.Cells[row, 11] = itemData.CableSizeEarthing;
                    worksheet.Cells[row, 12] = itemData.ActiveConductorMaterial;
                    worksheet.Cells[row, 13] = itemData.NumPhases; // Final value
                    worksheet.Cells[row, 14] = itemData.CableType; // Final value
                    worksheet.Cells[row, 15] = itemData.CableInsulation;
                    worksheet.Cells[row, 16] = itemData.InstallationMethod;
                    worksheet.Cells[row, 17] = itemData.CableAdditionalDerating;
                    worksheet.Cells[row, 18] = itemData.SwitchgearTripUnitType; // Final value
                    worksheet.Cells[row, 19] = itemData.SwitchgearManufacturer;
                    worksheet.Cells[row, 20] = itemData.BusType;
                    worksheet.Cells[row, 21] = itemData.BusChassisRating;
                    worksheet.Cells[row, 22] = itemData.UpstreamDiversity;
                    worksheet.Cells[row, 23] = itemData.IsolatorType;
                    worksheet.Cells[row, 24] = itemData.IsolatorRating;
                    worksheet.Cells[row, 25] = itemData.ProtectiveDeviceRating;
                    worksheet.Cells[row, 26] = itemData.ProtectiveDeviceManufacturer;
                    worksheet.Cells[row, 27] = itemData.ProtectiveDeviceType;
                    worksheet.Cells[row, 28] = itemData.ProtectiveDeviceModel;
                    worksheet.Cells[row, 29] = itemData.ProtectiveDeviceOCRTripUnit;
                    worksheet.Cells[row, 30] = itemData.ProtectiveDeviceTripSetting;

                    row++;
                }

                worksheet.Columns.AutoFit();
                workbook.SaveAs(xlsxFilePath);
                xlsxSuccess = true;
            }
            catch (Exception ex)
            {
                 xlsxError = $"Failed to write XLSX file: {ex.GetType().Name} - {ex.Message}";
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // --- Final Reporting ---
            StringBuilder finalMessage = new StringBuilder();
            finalMessage.AppendLine("Export Process Completed.");
            finalMessage.AppendLine("---");
            if (csvSuccess) { finalMessage.AppendLine($"CSV export successful:\n{csvFilePath}"); }
            else { finalMessage.AppendLine($"CSV export FAILED: {csvError ?? "Unknown error"}"); }
            finalMessage.AppendLine("---");
            if (xlsxSuccess) { finalMessage.AppendLine($"XLSX export successful:\n{xlsxFilePath}"); }
            else { finalMessage.AppendLine($"XLSX export FAILED: {xlsxError ?? "Unknown error"}"); }

            TaskDialog.Show("Export Results", finalMessage.ToString());

            if (csvSuccess && xlsxSuccess) { return Result.Succeeded; }
            else if (csvSuccess || xlsxSuccess) { return Result.Succeeded; }
            else { message = "Both CSV and XLSX exports failed."; return Result.Failed; }
        }

        // --- Recursive Pre-order Traversal Method ---
        private void PreOrderVisit(string nodeName)
        {
            if (!_visitedNodesDuringTraversal.Add(nodeName)) { return; } // Cycle detected or already processed path
            if (_itemsOriginatingFrom.TryGetValue(nodeName, out var outgoingItems))
            {
                foreach (var childItem in outgoingItems)
                {
                    _preOrderSortedData.Add(childItem);
                    PreOrderVisit(childItem.SWBTo);
                }
            }
            _visitedNodesDuringTraversal.Remove(nodeName); // Allow node to be visited via other paths
        }


        // --- Helper method to merge data ---
        private void MergeDetailItemData(DetailItemData existing, DetailItemData newItem)
        {
            // Prioritize existing non-blank values when merging duplicates from traversal
            if (string.IsNullOrWhiteSpace(existing.FinalCableReference) && !string.IsNullOrWhiteSpace(newItem.FinalCableReference)) existing.FinalCableReference = newItem.FinalCableReference;
            if (string.IsNullOrWhiteSpace(existing.SWBFrom) && !string.IsNullOrWhiteSpace(newItem.SWBFrom)) existing.SWBFrom = newItem.SWBFrom;
            if (string.IsNullOrWhiteSpace(existing.SWBType) && !string.IsNullOrWhiteSpace(newItem.SWBType)) existing.SWBType = newItem.SWBType;
            if (string.IsNullOrWhiteSpace(existing.SWBLoad) && !string.IsNullOrWhiteSpace(newItem.SWBLoad)) existing.SWBLoad = newItem.SWBLoad;
            if (string.IsNullOrWhiteSpace(existing.SWBLoadScope)) existing.SWBLoadScope = newItem.SWBLoadScope; // Should be "Local"
            if (string.IsNullOrWhiteSpace(existing.SWBPF)) existing.SWBPF = newItem.SWBPF; // Should be "1"
            if (string.IsNullOrWhiteSpace(existing.CableLength) && !string.IsNullOrWhiteSpace(newItem.CableLength)) existing.CableLength = newItem.CableLength;
            if (string.IsNullOrWhiteSpace(existing.CableSizeActive) && !string.IsNullOrWhiteSpace(newItem.CableSizeActive)) existing.CableSizeActive = newItem.CableSizeActive;
            if (string.IsNullOrWhiteSpace(existing.CableSizeNeutral) && !string.IsNullOrWhiteSpace(newItem.CableSizeNeutral)) existing.CableSizeNeutral = newItem.CableSizeNeutral;
            if (string.IsNullOrWhiteSpace(existing.CableSizeEarthing) && !string.IsNullOrWhiteSpace(newItem.CableSizeEarthing)) existing.CableSizeEarthing = newItem.CableSizeEarthing;
            if (string.IsNullOrWhiteSpace(existing.ActiveConductorMaterial) && !string.IsNullOrWhiteSpace(newItem.ActiveConductorMaterial)) existing.ActiveConductorMaterial = newItem.ActiveConductorMaterial;
            if (string.IsNullOrWhiteSpace(existing.NumPhases) && !string.IsNullOrWhiteSpace(newItem.NumPhases)) existing.NumPhases = newItem.NumPhases; // Merge "R" if found
            // CableType is set definitively in Step 5b
            if (string.IsNullOrWhiteSpace(existing.CableInsulation) && !string.IsNullOrWhiteSpace(newItem.CableInsulation)) existing.CableInsulation = newItem.CableInsulation; // Merge "X-HF-110" if found
            if (string.IsNullOrWhiteSpace(existing.InstallationMethod)) existing.InstallationMethod = "PT"; // Ensure "PT"
            if (string.IsNullOrWhiteSpace(existing.CableAdditionalDerating) && !string.IsNullOrWhiteSpace(newItem.CableAdditionalDerating)) existing.CableAdditionalDerating = newItem.CableAdditionalDerating;
            // SwitchgearTripUnitType is set definitively in Step 5b
            if (string.IsNullOrWhiteSpace(existing.SwitchgearManufacturer)) existing.SwitchgearManufacturer = "NAW Controls - LS Susol"; // Ensure Manufacturer
            if (string.IsNullOrWhiteSpace(existing.BusType)) existing.BusType = "Bus Bar"; // Ensure Bus Type
            if (string.IsNullOrWhiteSpace(existing.BusChassisRating) && !string.IsNullOrWhiteSpace(newItem.BusChassisRating)) existing.BusChassisRating = newItem.BusChassisRating;
            if (string.IsNullOrWhiteSpace(existing.UpstreamDiversity)) existing.UpstreamDiversity = "STD"; // Ensure Diversity
            if (string.IsNullOrWhiteSpace(existing.IsolatorType) && !string.IsNullOrWhiteSpace(newItem.IsolatorType)) existing.IsolatorType = newItem.IsolatorType;
            if (string.IsNullOrWhiteSpace(existing.IsolatorRating) && !string.IsNullOrWhiteSpace(newItem.IsolatorRating)) existing.IsolatorRating = newItem.IsolatorRating;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceRating) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceRating)) existing.ProtectiveDeviceRating = newItem.ProtectiveDeviceRating;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceManufacturer) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceManufacturer)) existing.ProtectiveDeviceManufacturer = newItem.ProtectiveDeviceManufacturer;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceType) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceType)) existing.ProtectiveDeviceType = newItem.ProtectiveDeviceType;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceModel) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceModel)) existing.ProtectiveDeviceModel = newItem.ProtectiveDeviceModel;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceOCRTripUnit) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceOCRTripUnit)) existing.ProtectiveDeviceOCRTripUnit = newItem.ProtectiveDeviceOCRTripUnit;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceTripSetting) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceTripSetting)) existing.ProtectiveDeviceTripSetting = newItem.ProtectiveDeviceTripSetting;
        }

    } // End class PowerCAD_ExportSLD
} // End namespace PC_Exporter
