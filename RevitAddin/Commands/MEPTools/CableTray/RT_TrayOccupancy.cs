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
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Refactored to use the centralized SharedParameters utility, removing all local hardcoded GUIDs.
// - [APPLIED FIX]: Removed the CleanCableReference method. The script now uses raw, unmodified
//                  parameter values for all calculations and updates.
//
// Author: Kyle Vorster
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Reflection;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Utilities; // Added to access the new SharedParameters
#endregion

namespace RTS.Commands.MEPTools.CableTray
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
                PC_ExtensibleClass pcExtensible = new PC_ExtensibleClass();
                List<PC_ExtensibleClass.CableData> cleanedData = pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    doc,
                    PC_ExtensibleClass.PrimarySchemaGuid,
                    PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName,
                    PC_ExtensibleClass.PrimaryDataStorageElementName
                );

                if (cleanedData == null || cleanedData.Count == 0)
                {
                    TaskDialog.Show("No Data Found", "No valid cable data was found in the project's primary extensible storage. Please run the 'PC_Extensible' command to import data first. The process will now exit.");
                    return Result.Succeeded;
                }

                // --- 2. PROMPT USER FOR OUTPUT FOLDER ---
                string outputFolderPath = GetOutputFolderPath();
                if (string.IsNullOrEmpty(outputFolderPath))
                {
                    return Result.Cancelled;
                }

                // --- 3. EXPORT THE RECALLED DATA TO A NEW CSV (Optional, for verification/record keeping) ---
                string cleanedScheduleFilePath = Path.Combine(outputFolderPath, "Cleaned_Cable_Schedule_From_Storage.csv");
                ExportDataToCsv(cleanedData, cleanedScheduleFilePath);


                // --- 4. PROCESS AND EXPORT CABLE TRAY DATA FROM REVIT MODEL ---
                List<TrayCableData> trayData = CollectTrayData(doc, cleanedData);
                if (trayData.Any())
                {
                    string trayDataFilePath = Path.Combine(outputFolderPath, "CableTray_Data.csv");
                    ExportTrayDataToCsv(trayData, trayDataFilePath);

                    // --- 5. UPDATE REVIT MODEL PARAMETERS ---
                    UpdateRevitParameters(doc, trayData);
                }

                // --- 6. PROCESS AND EXPORT CABLE LENGTH DATA ---
                List<CableLengthData> cableLengths = CollectCableLengthsData(doc, cleanedData);
                if (cableLengths.Any())
                {
                    string cableLengthsFilePath = Path.Combine(outputFolderPath, "Cable_Lengths.csv");
                    ExportCableLengthsToCsv(cableLengths, cableLengthsFilePath);

                    // --- 7. SAVE/MERGE CALCULATED CABLE LENGTHS DATA TO MODEL GENERATED EXTENSIBLE STORAGE ---
                    List<PC_ExtensibleClass.ModelGeneratedData> existingModelGeneratedData = pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                        doc,
                        PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                        PC_ExtensibleClass.ModelGeneratedSchemaName,
                        PC_ExtensibleClass.ModelGeneratedFieldName,
                        PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                    );

                    var mergedModelDataDict = existingModelGeneratedData
                        .Where(mgd => !string.IsNullOrEmpty(mgd.CableReference))
                        .GroupBy(mgd => mgd.CableReference)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    int updatedEntriesCount = 0;
                    int addedEntriesCount = 0;

                    foreach (var newCalculatedCable in cableLengths)
                    {
                        if (string.IsNullOrEmpty(newCalculatedCable.PC_Cable_Reference))
                        {
                            continue;
                        }

                        var newModelDataEntry = new PC_ExtensibleClass.ModelGeneratedData
                        {
                            CableReference = newCalculatedCable.PC_Cable_Reference,
                            From = newCalculatedCable.From,
                            To = newCalculatedCable.To,
                            CableLengthM = newCalculatedCable.PC_Cable_Length,
                            Variance = string.Empty,
                            Comment = newCalculatedCable.RTS_Comment
                        };

                        if (mergedModelDataDict.TryGetValue(newModelDataEntry.CableReference, out PC_ExtensibleClass.ModelGeneratedData existingEntry))
                        {
                            PropertyInfo[] properties = typeof(PC_ExtensibleClass.ModelGeneratedData).GetProperties();
                            bool entryChanged = false;

                            foreach (PropertyInfo prop in properties)
                            {
                                if (prop.PropertyType == typeof(string) && prop.CanWrite)
                                {
                                    string newValue = (string)prop.GetValue(newModelDataEntry);
                                    string currentValue = (string)prop.GetValue(existingEntry);

                                    if (!string.IsNullOrWhiteSpace(newValue) && !string.Equals(newValue, currentValue, StringComparison.Ordinal))
                                    {
                                        prop.SetValue(existingEntry, newValue);
                                        entryChanged = true;
                                    }
                                }
                            }
                            if (entryChanged) updatedEntriesCount++;
                        }
                        else
                        {
                            mergedModelDataDict.Add(newModelDataEntry.CableReference, newModelDataEntry);
                            addedEntriesCount++;
                        }
                    }

                    List<PC_ExtensibleClass.ModelGeneratedData> finalModelGeneratedData = mergedModelDataDict.Values.ToList();

                    using (Transaction tx = new Transaction(doc, "Save Model Generated Data"))
                    {
                        tx.Start();
                        try
                        {
                            pcExtensible.SaveDataToExtensibleStorage(
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
        private string GetOutputFolderPath()
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a Folder to Save the Exported Files";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return null;
        }
        #endregion

        #region CSV Export Methods
        private void ExportDataToCsv(List<PC_ExtensibleClass.CableData> data, string filePath)
        {
            var sb = new StringBuilder();
            var headers = new List<string>
            {
                "File Name", "Import Date",
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
                    row.FileName, row.ImportDate,
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
                    val = val.Trim();
                    if (val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r"))
                    {
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

        private List<TrayCableData> CollectTrayData(Document doc, List<PC_ExtensibleClass.CableData> cableScheduleData)
        {
            // *** FIX: Use centralized SharedParameters instead of hardcoded GUIDs ***
            Guid rtsIdGuid = SharedParameters.General.RTS_ID;
            List<Guid> cableGuids = SharedParameters.Cable.AllCableGuids;

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
                            if (!string.IsNullOrEmpty(originalValue))
                            {
                                groupedCables[rtsId].Add(originalValue);
                            }
                        }
                    }
                }
            }

            var trayDataList = new List<TrayCableData>();
            var cableDataDict = cableScheduleData
                .Where(c => !string.IsNullOrEmpty(c.To))
                .ToDictionary(c => c.To, c => c);

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

                foreach (string cableRef in sortedCables)
                {
                    var cableInfo = cableScheduleData.FirstOrDefault(c => c.CableReference == cableRef);

                    if (cableInfo != null)
                    {
                        if (double.TryParse(cableInfo.CablesKgPerM, out double kgPerM))
                        {
                            totalWeight += kgPerM;
                        }

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

                for (int i = 0; i < 30; i++)
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
            var cableRunData = new Dictionary<string, CableRunInfo>();
            // *** FIX: Use centralized SharedParameters ***
            List<Guid> cableGuids = SharedParameters.Cable.AllCableGuids;

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };
            var categoryFilter = new ElementMulticategoryFilter(categories);

            var collector = new FilteredElementCollector(doc)
               .WherePasses(categoryFilter)
               .WhereElementIsNotElementType();

            foreach (Element element in collector)
            {
                double lengthInFeet = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0;
                double lengthInMeters = lengthInFeet * 0.3048;

                if (lengthInMeters == 0) continue;

                string elementTypeCategory = element.Category.Name;
                string elementTypeName = doc.GetElement(element.GetTypeId())?.Name ?? "Unknown Type";


                foreach (var guid in cableGuids)
                {
                    Parameter cableParam = element.get_Parameter(guid);
                    if (cableParam != null && cableParam.HasValue)
                    {
                        string originalValue = cableParam.AsString();
                        if (!string.IsNullOrEmpty(originalValue))
                        {
                            if (!cableRunData.ContainsKey(originalValue))
                            {
                                cableRunData[originalValue] = new CableRunInfo();
                            }
                            cableRunData[originalValue].Lengths.Add(lengthInMeters);
                            cableRunData[originalValue].HostElementCategories.Add(elementTypeCategory);
                        }
                    }
                }
            }

            var cableDataLookup = cableScheduleData
                               .Where(cd => !string.IsNullOrEmpty(cd.CableReference))
                               .GroupBy(cd => cd.CableReference)
                               .ToDictionary(g => g.Key, g => g.First());

            var exportData = new List<CableLengthData>();
            foreach (var kvp in cableRunData)
            {
                string cableRef = kvp.Key;
                double totalLength = kvp.Value.Lengths.Sum();
                double finalLength = Math.Round(totalLength, 1) + 5.0;

                string fromValue = "Missing from PowerCAD";
                string toValue = "Missing from PowerCAD";
                string powerCadCableLength = "N/A";
                string variance = "N/A";
                string rtsComment = string.Join(":", kvp.Value.HostElementCategories.Distinct());

                var cableInfo = cableDataLookup.TryGetValue(cableRef, out PC_ExtensibleClass.CableData foundCableInfo) ? foundCableInfo : null;

                if (cableInfo != null)
                {
                    fromValue = cableInfo.From;
                    toValue = cableInfo.To;
                    powerCadCableLength = cableInfo.CableLength;

                    if (double.TryParse(powerCadCableLength, out double pcLength) && finalLength > 0)
                    {
                        double diff = finalLength - pcLength;
                        double percentageVariance = diff / finalLength * 100;
                        variance = percentageVariance.ToString("F2") + "%";
                    }
                }
                else
                {
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
                    CableLengthFromCsv = powerCadCableLength,
                    Variance = variance,
                    RTS_Comment = rtsComment
                });
            }

            return exportData.OrderBy(c => c.PC_Cable_Reference).ToList();
        }

        private void UpdateRevitParameters(Document doc, List<TrayCableData> trayDataList)
        {
            // *** FIX: Use centralized SharedParameters instead of hardcoded GUIDs ***
            Guid rtsIdGuid = SharedParameters.General.RTS_ID;
            Guid rtTrayOccupancyGuid = SharedParameters.Cable.RT_Tray_Occupancy;
            Guid rtCablesWeightGuid = SharedParameters.Cable.RT_Cables_Weight;
            Guid rtTrayMinSizeGuid = SharedParameters.Cable.RT_Tray_Min_Size;
            List<Guid> cableGuids = SharedParameters.Cable.AllCableGuids;

            using (Transaction tx = new Transaction(doc, "Update Cable Parameters"))
            {
                tx.Start();

                var trayDataDict = trayDataList.ToDictionary(t => t.RtsId);

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
                    .ToList();

                foreach (Element element in elementsToUpdate)
                {
#if REVIT2024_OR_GREATER
                    if (element.Category.Id.Value == (long)BuiltInCategory.OST_CableTray)
#else
                    if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
#endif
                    {
                        Parameter rtsIdParam = element.get_Parameter(rtsIdGuid);
                        if (rtsIdParam != null && rtsIdParam.HasValue)
                        {
                            string rtsId = rtsIdParam.AsString();
                            if (trayDataDict.ContainsKey(rtsId))
                            {
                                var data = trayDataDict[rtsId];

                                Parameter occupancyParam = element.get_Parameter(rtTrayOccupancyGuid);
                                if (occupancyParam != null && !occupancyParam.IsReadOnly)
                                {
                                    occupancyParam.Set(data.TrayOccupancy);
                                }

                                Parameter weightParam = element.get_Parameter(rtCablesWeightGuid);
                                if (weightParam != null && !weightParam.IsReadOnly)
                                {
                                    weightParam.Set(data.CablesWeight);
                                }

                                Parameter minSizeParam = element.get_Parameter(rtTrayMinSizeGuid);
                                if (minSizeParam != null && !minSizeParam.IsReadOnly)
                                {
                                    minSizeParam.Set(data.TrayMinSize);
                                }
                            }
                        }
                    }

                    HashSet<string> currentUniqueCablesOnElement = new HashSet<string>();
                    foreach (var guid in cableGuids)
                    {
                        Parameter cableParam = element.get_Parameter(guid);
                        if (cableParam != null && cableParam.HasValue)
                        {
                            string originalValue = cableParam.AsString();
                            if (!string.IsNullOrEmpty(originalValue))
                            {
                                currentUniqueCablesOnElement.Add(originalValue);
                            }
                        }
                    }

                    List<string> cablesToAssign = currentUniqueCablesOnElement.OrderBy(s => s).ToList();

                    for (int i = 0; i < cableGuids.Count; i++)
                    {
                        Parameter cableParam = element.get_Parameter(cableGuids[i]);
                        if (cableParam != null && !cableParam.IsReadOnly)
                        {
                            if (i < cablesToAssign.Count)
                            {
                                if (cableParam.AsString() != cablesToAssign[i])
                                {
                                    cableParam.Set(cablesToAssign[i]);
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    cableParam.Set(string.Empty);
                                }
                            }
                        }
                    }
                }

                tx.Commit();
                TaskDialog.Show("RT_TrayOccupancy", "Cable Tray data processed and updated successfully.");
            }
        }

        private void ExportTrayDataToCsv(List<TrayCableData> data, string filePath)
        {
            var sb = new StringBuilder();

            var headers = new List<string> { "RTS_ID" };
            for (int i = 1; i <= 30; i++)
            {
                headers.Add($"Cable_{i:D2}");
            }
            headers.Add("RT_Tray Occupancy");
            headers.Add("RT_Cables Weight");
            headers.Add("RT_Tray Min Size");
            headers.Add("RTS_Comment");

            sb.AppendLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var line = new List<string> { row.RtsId };
                line.AddRange(row.CableValues);
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

            var headers = new List<string> { "PC_Cable Reference", "From", "To", "PC_Cable Length", "Cable Length (m)", "Variance", "RTS_Comment" };
            sb.AppendLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var line = new List<string> {
                    row.PC_Cable_Reference,
                    row.From,
                    row.To,
                    row.PC_Cable_Length,
                    row.CableLengthFromCsv,
                    row.Variance,
                    row.RTS_Comment
                };
                var formattedLine = line.Select(val => val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r") ? $"\"{val.Replace("\"", "\"\"")}\"" : $"{(val ?? string.Empty).Trim()}");
                sb.AppendLine(string.Join(",", formattedLine));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        // *** FIX: Removed the local GetCableGuids() method. The code now uses SharedParameters.Cable.AllCableGuids directly. ***
        #endregion

        #region Data Classes
        private class TrayCableData
        {
            public string RtsId { get; set; }
            public List<string> CableValues { get; set; }
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
            public string CableLengthFromCsv { get; set; }
            public string Variance { get; set; }
            public string RTS_Comment { get; set; }
        }

        private class CableRunInfo
        {
            public List<double> Lengths { get; set; }
            public HashSet<string> HostElementCategories { get; set; }

            public CableRunInfo()
            {
                Lengths = new List<double>();
                HostElementCategories = new HashSet<string>();
            }
        }
        #endregion
    }

    public class MepElementSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category == null) return false;

#if REVIT2024_OR_GREATER
            long catId = elem.Category.Id.Value;
#else
            long catId = elem.Category.Id.IntegerValue;
#endif

            return catId == (long)BuiltInCategory.OST_LightingFixtures
                || catId == (long)BuiltInCategory.OST_ElectricalFixtures
                || catId == (long)BuiltInCategory.OST_ElectricalEquipment
                || catId == (long)BuiltInCategory.OST_CommunicationDevices
                || catId == (long)BuiltInCategory.OST_FireAlarmDevices
                || catId == (long)BuiltInCategory.OST_SecurityDevices
                || catId == (long)BuiltInCategory.OST_MechanicalEquipment
                || catId == (long)BuiltInCategory.OST_Sprinklers
                || catId == (long)BuiltInCategory.OST_PlumbingFixtures;
        }
        public bool AllowReference(Reference reference, XYZ position) { return false; }
    }
}
