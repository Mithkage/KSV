// File: PC_SWB_Exporter.cs
// C# Revit 2024 Add-in for exporting PowerCAD SWB data to CSV
// Includes Detail Items and Generic Annotations
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// using Excel = Microsoft.Office.Interop.Excel; // Removed
using System;
using System.Collections.Generic;
using System.Linq;
// using System.Runtime.InteropServices; // Removed (Only needed for Excel COM Interop)
using System.Text;
using System.IO;
using System.Globalization; // Added for robust number parsing

namespace PC_SWB_Exporter
{
    [Transaction(TransactionMode.Manual)]
    public class PC_SWB_ExporterClass : IExternalCommand
    {
        // Helper class to store data for each relevant element (Detail Item or Generic Annotation)
        private class DetailItemData // Name kept for simplicity, but now represents data from multiple categories
        {
            public string OriginalCableReference { get; set; }
            public string FinalCableReference { get; set; }
            public string SWBFrom { get; set; }
            public string SWBTo { get; set; }
            public string SWBType { get; set; }
            public string SWBLoad { get; set; } // This will be populated by "PC_SWB Load" or fallback
            public string SWBLoadScope { get; set; }
            public string SWBPF { get; set; }
            public string CableLength { get; set; }
            public string CableSizeActive { get; set; }
            public string CableSizeNeutral { get; set; }
            public string CableSizeEarthing { get; set; }
            public string ActiveConductorMaterial { get; set; }
            public string NumPhases { get; set; }
            public string CableType { get; set; }
            public string CableInsulation { get; set; }
            public string InstallationMethod { get; set; }
            public string CableAdditionalDerating { get; set; }
            public string SwitchgearTripUnitType { get; set; }
            public string SwitchgearManufacturer { get; set; }
            public string BusType { get; set; }
            public string BusChassisRating { get; set; }
            public string UpstreamDiversity { get; set; }
            public string IsolatorType { get; set; }
            public string IsolatorRating { get; set; }
            public string ProtectiveDeviceRating { get; set; }
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

            // --- MODIFICATION: Collect both Detail Items and Generic Annotations ---
            // Filter for Detail Items in the active document
            FilteredElementCollector detailItemCollector = new FilteredElementCollector(doc);
            List<Element> detailItems = detailItemCollector.OfCategory(BuiltInCategory.OST_DetailComponents)
                                                   .WhereElementIsNotElementType() // Get instances, not types
                                                   .ToList();

            // Filter for Generic Annotations in the active document
            FilteredElementCollector genericAnnotationCollector = new FilteredElementCollector(doc);
            List<Element> genericAnnotations = genericAnnotationCollector.OfCategory(BuiltInCategory.OST_GenericAnnotation)
                                                                .WhereElementIsNotElementType() // Get instances, not types
                                                                .ToList();

            // Combine both lists
            List<Element> allRelevantElements = new List<Element>();
            allRelevantElements.AddRange(detailItems);
            allRelevantElements.AddRange(genericAnnotations);
            // --- END MODIFICATION ---


            // Check if any relevant elements were found
            if (allRelevantElements.Count == 0)
            {
                TaskDialog.Show("Info", "No Detail Items or Generic Annotations found in the current view/document.");
                return Result.Succeeded; // Nothing to process
            }

            // --- Step 1: Data Collection & Pre-processing ---
            var groupedData = new Dictionary<string, List<DetailItemData>>(); // Group items by their 'SWB To' value
            var allNodes = new HashSet<string>(); // Keep track of all unique 'SWB To' and 'SWB From' values

            // --- MODIFICATION: Iterate over allRelevantElements ---
            foreach (Element element in allRelevantElements)
            // --- END MODIFICATION ---
            {
                // Get relevant parameters from the element
                Parameter pcPowerCADParam = element.LookupParameter("PC_PowerCAD");
                Parameter pcSWBToParam = element.LookupParameter("PC_SWB To");
                Parameter pcSWBFromParam = element.LookupParameter("PC_SWB From"); //was PC_From but updated to PC_SWB From for clarity
                Parameter pcCableLengthParam = element.LookupParameter("PC_Cable Length");
                // Use "PC_Protective Device Trip Setting (A)" for protective device fields
                Parameter pcProtectiveDeviceTripSettingParam = element.LookupParameter("PC_Protective Device Trip Setting (A)");
                // Get the "PC_SWB Load" parameter
                Parameter pcSWBLoadParam = element.LookupParameter("PC_SWB Load");


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

                    Parameter pcCableReferenceParam = element.LookupParameter("PC_Cable Reference");
                    string cableReferenceValue = pcCableReferenceParam?.AsString() ?? "";

                    // Add the 'From' and 'To' nodes to the set for graph building
                    allNodes.Add(swbToValue);
                    allNodes.Add(swbFromValue);

                    // Initialize variables to store type-based information
                    string swbTypeValue = "";
                    string cableTypeValue = "";
                    string cableInsulationValue = "";
                    string numPhasesValue = "";
                    string switchgearTripUnitTypeValue = "";
                    string cableLengthValue = pcCableLengthParam?.AsString() ?? "";

                    // Get the protective device trip setting value (used as fallback and for other PD fields)
                    string protectiveDeviceTripSettingValue = pcProtectiveDeviceTripSettingParam?.AsString() ?? "";
                    // Get the specific SWB Load value
                    string swbLoadActualValue = pcSWBLoadParam?.AsString();

                    // Initialize Isolator specific variables
                    string isolatorTypeValue = "";
                    string isolatorRatingValue = "";

                    // Get the Element Type of the element
                    ElementId elementTypeId = element.GetTypeId();
                    if (elementTypeId != null && elementTypeId != ElementId.InvalidElementId)
                    {
                        ElementType elementType = doc.GetElement(elementTypeId) as ElementType;
                        if (elementType != null)
                        {
                            string typeName = elementType.Name;
                            string familyName = elementType.FamilyName; // Get Family Name

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

                                // Check if Family Name contains "Isolator Type"
                                if (!string.IsNullOrEmpty(familyName) && familyName.IndexOf("Isolator Type", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Get Isolator Rating parameter
                                    Parameter pcIsolatorRatingParam = element.LookupParameter("PC_Isolator Rating (A)");
                                    isolatorRatingValue = pcIsolatorRatingParam?.AsString() ?? "";

                                    // Check Element Type Name for specific keywords
                                    if (typeName.IndexOf("Off Load", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                        typeName.IndexOf("CB (Non-Auto)", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        isolatorTypeValue = "Switch (Isolating)";
                                    }
                                    else if (typeName.IndexOf("On Load", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        isolatorTypeValue = "Switch (Load Break)";
                                    }
                                    else if (typeName.IndexOf("CB (Auto)", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        isolatorTypeValue = "Circuit Breaker";
                                    }
                                }
                            }
                        }
                    }

                    // Check PC_From for "SAFETY" to set Cable Insulation
                    if (swbFromValue.IndexOf("SAFETY", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cableInsulationValue = "X-HF-110";
                    }

                    // Create a data object for the current element
                    var data = new DetailItemData
                    {
                        OriginalCableReference = cableReferenceValue,
                        SWBFrom = swbFromValue,
                        SWBTo = swbToValue,
                        SWBType = swbTypeValue,
                        SWBLoad = !string.IsNullOrWhiteSpace(swbLoadActualValue) ? swbLoadActualValue : protectiveDeviceTripSettingValue,
                        SWBLoadScope = "Local",
                        SWBPF = "1",
                        CableLength = cableLengthValue,
                        CableSizeActive = "",
                        CableSizeNeutral = "",
                        CableSizeEarthing = "",
                        ActiveConductorMaterial = "",
                        NumPhases = numPhasesValue,
                        CableType = cableTypeValue,
                        CableInsulation = cableInsulationValue,
                        InstallationMethod = "PT",
                        CableAdditionalDerating = "",
                        SwitchgearTripUnitType = switchgearTripUnitTypeValue,
                        SwitchgearManufacturer = "NAW Controls - LS Susol",
                        BusType = "Bus Bar",
                        BusChassisRating = element.LookupParameter("PC_Bus/Chassis Rating (A)")?.AsString() ?? "",
                        UpstreamDiversity = "STD",
                        IsolatorType = isolatorTypeValue,
                        IsolatorRating = isolatorRatingValue,
                        ProtectiveDeviceRating = protectiveDeviceTripSettingValue,
                        ProtectiveDeviceManufacturer = "",
                        ProtectiveDeviceType = "",
                        ProtectiveDeviceModel = "",
                        ProtectiveDeviceOCRTripUnit = "",
                        ProtectiveDeviceTripSetting = protectiveDeviceTripSettingValue
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
                // --- MODIFICATION: Update message ---
                TaskDialog.Show("Info", "No Detail Items or Generic Annotations marked for export (PC_PowerCAD = 1) or with valid 'PC_SWB To' values were found.");
                // --- END MODIFICATION ---
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
            foreach (string node in allNodes)
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
            foreach (var newItemData in _preOrderSortedData)
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
                    // --- Set Cable Type based on SWB Type or Length/Rating ---
                    if (itemData.SWBType == "T") // If it's a Tap-Off Box type
                    {
                        itemData.CableType = "SDI";
                    }
                    else // If not 'S' and not 'T', apply length/rating logic
                    {
                        // Use ProtectiveDeviceTripSetting for rating check here as well
                        bool lengthParsed = double.TryParse(itemData.CableLength, NumberStyles.Any, CultureInfo.InvariantCulture, out double cableLength);
                        bool ratingParsed = double.TryParse(itemData.ProtectiveDeviceTripSetting, NumberStyles.Any, CultureInfo.InvariantCulture, out double deviceRating);

                        // Set to SDI if (Length >= 100 AND Rating >= 160) OR (Rating >= 229)
                        if ((lengthParsed && ratingParsed && cableLength >= 100.0 && deviceRating >= 160.0) ||
                            (ratingParsed && deviceRating >= 229.0))
                        {
                            itemData.CableType = "SDI";
                        }
                        else // Neither condition met
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


                // --- Final Isolator Check/Default Logic ---
                if (string.IsNullOrWhiteSpace(itemData.IsolatorType))
                {
                    if (itemData.SWBType == "S")
                    {
                        itemData.IsolatorType = "None";
                        itemData.IsolatorRating = ""; // Ensure rating is blank for "None" type
                    }
                    else // SWB Type is not "S"
                    {
                        // Try parsing the Protective Device Trip Setting
                        bool ratingParsed = double.TryParse(itemData.ProtectiveDeviceTripSetting, NumberStyles.Any, CultureInfo.InvariantCulture, out double tripSetting);

                        if (ratingParsed)
                        {
                            if (tripSetting <= 250.0)
                            {
                                itemData.IsolatorType = "Switch (Load Break)";
                                itemData.IsolatorRating = "250";
                            }
                            else if (tripSetting > 250.0 && tripSetting <= 630.0)
                            {
                                itemData.IsolatorType = "Switch (Isolating)";
                                itemData.IsolatorRating = "630";
                            }
                            else // tripSetting > 630.0
                            {
                                itemData.IsolatorType = "Circuit Breaker";
                                itemData.IsolatorRating = "3200";
                            }
                        }
                        else
                        {
                            // If parsing fails, perhaps leave blank or set a default?
                            itemData.IsolatorType = ""; // Explicitly keep blank if rating is unparseable
                            itemData.IsolatorRating = "";
                        }
                    }
                }
                // --- END Final Isolator Check ---

            } // End foreach loop for final logic


            // --- Step 6: Write Output File (CSV Only) ---
            string baseFileName = "PowerCAD_Export_SLD_Data";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string csvFilePath = Path.Combine(desktopPath, baseFileName + ".csv");

            bool csvSuccess = false;
            string csvError = null;

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
                        excelFormattedCableRef,
                        EscapeCsvField(displaySwbFrom),
                        EscapeCsvField(itemData.SWBTo),
                        EscapeCsvField(itemData.SWBType),
                        EscapeCsvField(itemData.SWBLoad), // This now reflects "PC_SWB Load" or its fallback
                        EscapeCsvField(itemData.SWBLoadScope),
                        EscapeCsvField(itemData.SWBPF),
                        EscapeCsvField(itemData.CableLength),
                        EscapeCsvField(itemData.CableSizeActive),
                        EscapeCsvField(itemData.CableSizeNeutral),
                        EscapeCsvField(itemData.CableSizeEarthing),
                        EscapeCsvField(itemData.ActiveConductorMaterial),
                        EscapeCsvField(itemData.NumPhases),
                        EscapeCsvField(itemData.CableType),
                        EscapeCsvField(itemData.CableInsulation),
                        EscapeCsvField(itemData.InstallationMethod),
                        EscapeCsvField(itemData.CableAdditionalDerating),
                        EscapeCsvField(itemData.SwitchgearTripUnitType),
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

            // --- XLSX File Writing Section Removed ---


            // --- Final Reporting ---
            StringBuilder finalMessage = new StringBuilder();
            finalMessage.AppendLine("Export Process Completed.");
            finalMessage.AppendLine("---");
            if (csvSuccess) { finalMessage.AppendLine($"CSV export successful:\n{csvFilePath}"); }
            else { finalMessage.AppendLine($"CSV export FAILED: {csvError ?? "Unknown error"}"); }
            // Removed XLSX reporting lines

            TaskDialog.Show("Export Results", finalMessage.ToString());

            // Determine overall result based on CSV success only
            if (csvSuccess) { return Result.Succeeded; }
            else { message = "CSV export failed."; return Result.Failed; }
        }

        // --- Recursive Pre-order Traversal Method ---
        private void PreOrderVisit(string nodeName)
        {
            // Check if the node has already been fully processed in the current path to prevent infinite loops in cycles
            if (!_visitedNodesDuringTraversal.Add(nodeName))
            {
                // Cycle detected or node already visited in this path
                return;
            }

            // Find all items that originate from the current node
            if (_itemsOriginatingFrom.TryGetValue(nodeName, out var outgoingItems))
            {
                // Sort outgoing items based on their 'SWB To' to ensure consistent order if needed
                outgoingItems = outgoingItems.OrderBy(item => item.SWBTo).ToList();

                foreach (var childItem in outgoingItems)
                {
                    // Add the child item to the sorted list *before* visiting its children (Pre-order)
                    _preOrderSortedData.Add(childItem);
                    // Recursively visit the destination node of the child item
                    PreOrderVisit(childItem.SWBTo);
                }
            }

            // Backtrack: Remove the node from the visited set for the *current path*
            _visitedNodesDuringTraversal.Remove(nodeName);
        }


        // --- Helper method to merge data ---
        private void MergeDetailItemData(DetailItemData existing, DetailItemData newItem)
        {
            // Prioritize existing non-blank values, especially for Isolator fields handled earlier
            if (string.IsNullOrWhiteSpace(existing.FinalCableReference) && !string.IsNullOrWhiteSpace(newItem.FinalCableReference)) existing.FinalCableReference = newItem.FinalCableReference;
            if (string.IsNullOrWhiteSpace(existing.SWBFrom) && !string.IsNullOrWhiteSpace(newItem.SWBFrom)) existing.SWBFrom = newItem.SWBFrom;
            if (string.IsNullOrWhiteSpace(existing.SWBType) && !string.IsNullOrWhiteSpace(newItem.SWBType)) existing.SWBType = newItem.SWBType;

            // SWBLoad merge logic: If existing is blank and new is not, take new.
            // The initial population of SWBLoad already handles the "PC_SWB Load" vs "ProtectiveDeviceTripSetting" priority.
            if (string.IsNullOrWhiteSpace(existing.SWBLoad) && !string.IsNullOrWhiteSpace(newItem.SWBLoad)) existing.SWBLoad = newItem.SWBLoad;

            if (string.IsNullOrWhiteSpace(existing.SWBLoadScope)) existing.SWBLoadScope = newItem.SWBLoadScope; // Typically "Local"
            if (string.IsNullOrWhiteSpace(existing.SWBPF)) existing.SWBPF = newItem.SWBPF; // Typically "1"
            if (string.IsNullOrWhiteSpace(existing.CableLength) && !string.IsNullOrWhiteSpace(newItem.CableLength)) existing.CableLength = newItem.CableLength;
            if (string.IsNullOrWhiteSpace(existing.CableSizeActive) && !string.IsNullOrWhiteSpace(newItem.CableSizeActive)) existing.CableSizeActive = newItem.CableSizeActive;
            if (string.IsNullOrWhiteSpace(existing.CableSizeNeutral) && !string.IsNullOrWhiteSpace(newItem.CableSizeNeutral)) existing.CableSizeNeutral = newItem.CableSizeNeutral;
            if (string.IsNullOrWhiteSpace(existing.CableSizeEarthing) && !string.IsNullOrWhiteSpace(newItem.CableSizeEarthing)) existing.CableSizeEarthing = newItem.CableSizeEarthing;
            if (string.IsNullOrWhiteSpace(existing.ActiveConductorMaterial) && !string.IsNullOrWhiteSpace(newItem.ActiveConductorMaterial)) existing.ActiveConductorMaterial = newItem.ActiveConductorMaterial;
            if (string.IsNullOrWhiteSpace(existing.NumPhases) && !string.IsNullOrWhiteSpace(newItem.NumPhases)) existing.NumPhases = newItem.NumPhases;
            // CableType is determined in Step 5b, so no direct merge here.
            if (string.IsNullOrWhiteSpace(existing.CableInsulation) && !string.IsNullOrWhiteSpace(newItem.CableInsulation)) existing.CableInsulation = newItem.CableInsulation;
            if (string.IsNullOrWhiteSpace(existing.InstallationMethod)) existing.InstallationMethod = "PT"; // Default
            if (string.IsNullOrWhiteSpace(existing.CableAdditionalDerating) && !string.IsNullOrWhiteSpace(newItem.CableAdditionalDerating)) existing.CableAdditionalDerating = newItem.CableAdditionalDerating;
            // SwitchgearTripUnitType is determined in Step 5b.
            if (string.IsNullOrWhiteSpace(existing.SwitchgearManufacturer)) existing.SwitchgearManufacturer = "NAW Controls - LS Susol"; // Default
            if (string.IsNullOrWhiteSpace(existing.BusType)) existing.BusType = "Bus Bar"; // Default
            if (string.IsNullOrWhiteSpace(existing.BusChassisRating) && !string.IsNullOrWhiteSpace(newItem.BusChassisRating)) existing.BusChassisRating = newItem.BusChassisRating;
            if (string.IsNullOrWhiteSpace(existing.UpstreamDiversity)) existing.UpstreamDiversity = "STD"; // Default

            // Merge Isolator Type/Rating only if the existing one is blank (prioritize values set in Step 1 or defaults from Step 5b)
            if (string.IsNullOrWhiteSpace(existing.IsolatorType) && !string.IsNullOrWhiteSpace(newItem.IsolatorType)) existing.IsolatorType = newItem.IsolatorType;
            if (string.IsNullOrWhiteSpace(existing.IsolatorRating) && !string.IsNullOrWhiteSpace(newItem.IsolatorRating)) existing.IsolatorRating = newItem.IsolatorRating;

            // Merge Protective Device fields - prioritize Trip Setting if existing is blank
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceTripSetting) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceTripSetting)) existing.ProtectiveDeviceTripSetting = newItem.ProtectiveDeviceTripSetting;

            // If ProtectiveDeviceRating is blank, try to fill it from the (potentially newly merged) ProtectiveDeviceTripSetting
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceRating)) existing.ProtectiveDeviceRating = existing.ProtectiveDeviceTripSetting;

            // If SWBLoad is still blank after its primary merge attempt, try to fill it from ProtectiveDeviceTripSetting as a final fallback.
            // This is secondary to its initial population from "PC_SWB Load" or the fallback during item creation.
            if (string.IsNullOrWhiteSpace(existing.SWBLoad)) existing.SWBLoad = existing.ProtectiveDeviceTripSetting;


            // Merge other PD fields if blank
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceManufacturer) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceManufacturer)) existing.ProtectiveDeviceManufacturer = newItem.ProtectiveDeviceManufacturer;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceType) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceType)) existing.ProtectiveDeviceType = newItem.ProtectiveDeviceType;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceModel) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceModel)) existing.ProtectiveDeviceModel = newItem.ProtectiveDeviceModel;
            if (string.IsNullOrWhiteSpace(existing.ProtectiveDeviceOCRTripUnit) && !string.IsNullOrWhiteSpace(newItem.ProtectiveDeviceOCRTripUnit)) existing.ProtectiveDeviceOCRTripUnit = newItem.ProtectiveDeviceOCRTripUnit;
        }

    } // End class PC_SWB_ExporterClass
} // End namespace PC_SWB_Exporter
