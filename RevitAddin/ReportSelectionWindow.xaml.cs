//-----------------------------------------------------------------------------
// <copyright file="ReportSelectionWindow.xaml.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     This file contains the interaction logic for the ReportSelectionWindow.xaml.
//     It handles user selections for various report types and orchestrates
//     the data recall from Revit's extensible storage and subsequent export
//     to CSV format, particularly for the RSGx Cable Schedule.
// </summary>
//-----------------------------------------------------------------------------

/*
 * Change Log:
 *
 * Date       | Version | Author | Description
 * -----------|---------|--------|----------------------------------------------------------------------------------------------------
 * 2025-07-04 | 1.0.0   | Gemini | Initial implementation to address column and data mixing in RSGx report.
 * |         |         |        | - Introduced explicit ordering for RSGxCableData properties in ExportDataToCsvGeneric.
 * |         |         |        | - Added `orderedPropertiesToExport` to ensure CSV column order matches defined headers.
 * |         |         |        | - Included a check for missing properties in RSGxCableData to enhance robustness.
 * 2025-07-04 | 1.0.1   | Gemini | Added file header comments and a change log as requested.
 * |         |         |        | - Included detailed description of file purpose and function.
 * 2025-07-07 | 1.1.0   | Gemini | Replaced WindowsAPICodePack folder browser with standard System.Windows.Forms.FolderBrowserDialog.
 * |         |         |        | - This removes the final dependency on the conflicting library to resolve build and runtime errors.
 * |         |         |        | - Added a Win32Window helper class to properly parent the dialog to the Revit window.
 * 2025-07-07 | 1.1.1   | Gemini | Resolved ambiguous reference error for IWin32Window by specifying the System.Windows.Forms namespace.
 * 2025-07-08 | 1.2.0   | Gemini | Updated RSGx Cable Schedule generation logic based on user feedback.
 * |         |         |        | - Relabeled 'Cable Type' header to 'Cores' and mapped it to the 'Cores' data field.
 * |         |         |        | - Implemented conditional logic for 'Fire Rating' based on the cable's 'Type' property.
 * |         |         |        | - Implemented conditional logic for 'Update Summary' based on cable length differences.
 * |         |         |        | - Removed 'Installation Configuration' column from the report.
 * 2025-07-08 | 1.3.0   | Gemini | Enhanced RSGx report data formatting.
 * |         |         |        | - 'Maximum Length' is now rounded down to the nearest integer.
 * |         |         |        | - 'Cable Description' is now a concatenation of key cable properties, omitting nulls.
 * 2025-07-08 | 1.4.0   | Gemini | Refined RSGx report data formatting.
 * |         |         |        | - Appended "mm²" unit to 'Active Cable Size' in the 'Cable Description'.
 * |         |         |        | - Set 'Voltage Rating' to null if 'Destination Device (ID)' is not available.
 * 2025-07-08 | 1.5.0   | Gemini | Added "Load (A)" column to the RSGx Cable Schedule report.
 * |         |         |        | - Data is sourced from the new LoadA property in the primary data store.
 * 2025-07-08 | 1.6.0   | Gemini | Added Maximum Demand (MD) comparison to the "Update Summary" field.
 * |         |         |        | - Compares "Load (A)" between primary and consultant data.
 * 2025-07-08 | 1.7.0   | Gemini | Added "Number of Earth Cables" column to the RSGx report.
 * |         |         |        | - Column inserted before "Earth Size (mm2)".
 * 2025-07-08 | 1.8.0   | Gemini | Removed the "No." column (related to neutral cables) from the RSGx report.
 */

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible;
using System.Text.RegularExpressions;
using System.Windows.Forms; // For FolderBrowserDialog
using System.Windows.Interop; // For WindowInteropHelper
#endregion

namespace RTS_Reports
{
    /// <summary>
    /// Interaction logic for ReportSelectionWindow.xaml
    /// This class now also contains the logic for generating reports.
    /// </summary>
    public partial class ReportSelectionWindow : Window
    {
        private Document _doc;
        private PC_ExtensibleClass _pcExtensible;

        public ReportSelectionWindow(ExternalCommandData commandData)
        {
            InitializeComponent();
            _doc = commandData.Application.ActiveUIDocument.Document;
            _pcExtensible = new PC_ExtensibleClass();
        }

        private void ExportMyPowerCADDataButton_Click(object sender, RoutedEventArgs e)
        {
            PerformExport<PC_ExtensibleClass.CableData>(
                PC_ExtensibleClass.PrimarySchemaGuid,
                PC_ExtensibleClass.PrimarySchemaName,
                PC_ExtensibleClass.PrimaryFieldName,
                PC_ExtensibleClass.PrimaryDataStorageElementName,
                "My_PowerCAD_Data_Report.csv",
                "My PowerCAD Data"
            );
        }

        private void ExportConsultantPowerCADDataButton_Click(object sender, RoutedEventArgs e)
        {
            PerformExport<PC_ExtensibleClass.CableData>(
                PC_ExtensibleClass.ConsultantSchemaGuid,
                PC_ExtensibleClass.ConsultantSchemaName,
                PC_ExtensibleClass.ConsultantFieldName,
                PC_ExtensibleClass.ConsultantDataStorageElementName,
                "Consultant_PowerCAD_Data_Report.csv",
                "Consultant PowerCAD Data"
            );
        }

        private void ExportModelGeneratedDataButton_Click(object sender, RoutedEventArgs e)
        {
            PerformExport<PC_ExtensibleClass.ModelGeneratedData>(
                PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                PC_ExtensibleClass.ModelGeneratedSchemaName,
                PC_ExtensibleClass.ModelGeneratedFieldName,
                PC_ExtensibleClass.ModelGeneratedDataStorageElementName,
                "Model_Generated_Data_Report.csv",
                "Model Generated Data"
            );
        }

        private void GenerateRSGxCableScheduleCsvReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

            string outputFolderPath = GetOutputFolderPath();
            if (string.IsNullOrEmpty(outputFolderPath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output folder not selected. RSGx Cable Schedule export cancelled.");
                return;
            }

            string filePath = System.IO.Path.Combine(outputFolderPath, "RSGx Cable Schedule.csv");

            try
            {
                List<PC_ExtensibleClass.CableData> primaryData = _pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    _doc, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName
                );

                List<PC_ExtensibleClass.CableData> consultantData = _pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    _doc, PC_ExtensibleClass.ConsultantSchemaGuid, PC_ExtensibleClass.ConsultantSchemaName,
                    PC_ExtensibleClass.ConsultantFieldName, PC_ExtensibleClass.ConsultantDataStorageElementName
                );

                List<PC_ExtensibleClass.ModelGeneratedData> modelGeneratedData = _pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                    _doc, PC_ExtensibleClass.ModelGeneratedSchemaGuid, PC_ExtensibleClass.ModelGeneratedSchemaName,
                    PC_ExtensibleClass.ModelGeneratedFieldName, PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                );

                var primaryDict = primaryData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var consultantDict = consultantData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var modelGeneratedDict = modelGeneratedData.ToDictionary(m => m.CableReference, m => m, StringComparer.OrdinalIgnoreCase);

                var allCableRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cable in primaryData) allCableRefs.Add(cable.CableReference);
                foreach (var cable in consultantData) allCableRefs.Add(cable.CableReference);
                foreach (var cable in modelGeneratedData) allCableRefs.Add(cable.CableReference);
                allCableRefs.RemoveWhere(string.IsNullOrEmpty);

                var sortedUniqueCableRefs = allCableRefs.OrderBy(cr => cr).ToList();
                var reportData = new List<RSGxCableData>();
                int rowNum = 1;

                foreach (string cableRef in sortedUniqueCableRefs)
                {
                    primaryDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData primaryInfo);
                    consultantDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData consultantInfo);
                    modelGeneratedDict.TryGetValue(cableRef, out PC_ExtensibleClass.ModelGeneratedData modelInfo);

                    string originDeviceID = primaryInfo?.From ?? consultantInfo?.From ?? "N/A";
                    string destinationDeviceID = primaryInfo?.To ?? consultantInfo?.To ?? "N/A";
                    string rsgxRouteLength = primaryInfo?.CableLength ?? "N/A";
                    string djvDesignLength = consultantInfo?.CableLength ?? "N/A";

                    string cableLengthDifference = "N/A";
                    if (double.TryParse(rsgxRouteLength, out double rsgxLen) && double.TryParse(djvDesignLength, out double djvLen))
                    {
                        cableLengthDifference = (rsgxLen - djvLen).ToString("F1");
                    }

                    string cableSizeChangeYN = "N/A";
                    if (double.TryParse(rsgxRouteLength, out rsgxLen) && double.TryParse(djvDesignLength, out djvLen))
                    {
                        cableSizeChangeYN = (rsgxLen == djvLen) ? "N" : "Y";
                    }

                    string earthIncludedRaw = primaryInfo?.SeparateEarthForMulticore;
                    string earthIncludedFormatted = string.Equals(earthIncludedRaw, "No", StringComparison.OrdinalIgnoreCase) ? "Yes" : "No";

                    string fireRatingValue = "";
                    if (primaryInfo?.CableType != null && primaryInfo.CableType.IndexOf("Fire", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        fireRatingValue = "WS52W";
                    }

                    var summaryParts = new List<string>();
                    if (double.TryParse(cableLengthDifference, out double diff))
                    {
                        if (diff > 0) summaryParts.Add("Length increased.");
                        else if (diff < 0) summaryParts.Add("Length decreased.");
                        else summaryParts.Add("Length consistent.");
                    }

                    string primaryLoadStr = primaryInfo?.LoadA;
                    string consultantLoadStr = consultantInfo?.LoadA;
                    if (double.TryParse(primaryLoadStr, out double primaryLoad) && double.TryParse(consultantLoadStr, out double consultantLoad))
                    {
                        if (primaryLoad < consultantLoad) summaryParts.Add("MD decreased");
                        else if (primaryLoad > consultantLoad) summaryParts.Add("MD increased");
                        else summaryParts.Add("MD un-changed");
                    }
                    else
                    {
                        summaryParts.Add("N/A");
                    }
                    string updateSummaryValue = string.Join(" | ", summaryParts.Where(s => !string.IsNullOrEmpty(s)));


                    string maxLengthPermissible = primaryInfo?.CableMaxLengthM ?? "N/A";
                    if (double.TryParse(maxLengthPermissible, out double maxLen))
                    {
                        maxLengthPermissible = Math.Floor(maxLen).ToString();
                    }
                    else
                    {
                        maxLengthPermissible = "N/A";
                    }

                    string voltageRating = "0.6/1kV";
                    if (destinationDeviceID == "N/A")
                    {
                        voltageRating = ""; // Set to null/empty
                    }

                    var descriptionParts = new List<string>();
                    string activeCableSize = primaryInfo?.ActiveCableSize ?? "N/A";
                    string cores = primaryInfo?.Cores ?? "N/A";
                    string type = primaryInfo?.CableType ?? "N/A";
                    string sheath = primaryInfo?.Sheath ?? "N/A";
                    string insulation = primaryInfo?.Insulation ?? "N/A";

                    if (!string.IsNullOrEmpty(activeCableSize) && activeCableSize != "N/A") descriptionParts.Add($"{activeCableSize}mm²");
                    if (!string.IsNullOrEmpty(cores) && cores != "N/A") descriptionParts.Add(cores);
                    if (!string.IsNullOrEmpty(voltageRating) && voltageRating != "N/A") descriptionParts.Add(voltageRating);
                    if (!string.IsNullOrEmpty(type) && type != "N/A") descriptionParts.Add(type);
                    if (!string.IsNullOrEmpty(sheath) && sheath != "N/A") descriptionParts.Add(sheath);
                    if (!string.IsNullOrEmpty(insulation) && insulation != "N/A") descriptionParts.Add(insulation);
                    if (!string.IsNullOrEmpty(fireRatingValue) && fireRatingValue != "N/A") descriptionParts.Add(fireRatingValue);
                    string cableDescriptionValue = string.Join(" | ", descriptionParts);

                    reportData.Add(new RSGxCableData
                    {
                        RowNumber = (rowNum++).ToString(),
                        CableTag = cableRef,
                        OriginDeviceID = originDeviceID,
                        DestinationDeviceID = destinationDeviceID,
                        RSGxRouteLengthM = rsgxRouteLength,
                        DJVDesignLengthM = djvDesignLength,
                        CableLengthDifferenceM = cableLengthDifference,
                        MaxLengthPermissibleForCableSizeM = maxLengthPermissible,
                        ActiveCableSizeMM2 = activeCableSize,
                        NoOfSets = primaryInfo?.NumberOfActiveCables ?? "N/A",
                        NeutralCableSizeMM2 = primaryInfo?.NeutralCableSize ?? "N/A",
                        Cores = cores,
                        ConductorType = primaryInfo?.ConductorActive ?? "N/A",
                        CableSizeChangeFromDesignYN = cableSizeChangeYN,
                        PreviousDesignSize = consultantInfo?.ActiveCableSize ?? "N/A",
                        EarthIncludedYesNo = earthIncludedFormatted,
                        NumberOfEarthCables = primaryInfo?.NumberOfEarthCables ?? "N/A",
                        EarthSizeMM2 = primaryInfo?.EarthCableSize ?? "N/A",
                        Voltage = primaryInfo?.VoltageVac ?? "N/A",
                        VoltageRating = voltageRating,
                        Type = type,
                        SheathConstruction = sheath,
                        InsulationConstruction = insulation,
                        FireRating = fireRatingValue,
                        LoadA = primaryInfo?.LoadA ?? "N/A",
                        CableDescription = cableDescriptionValue,
                        Comments = modelInfo?.Comment ?? "N/A",
                        UpdateSummary = updateSummaryValue
                    });
                }

                if (!reportData.Any())
                {
                    Autodesk.Revit.UI.TaskDialog.Show("No Data", "No unique cable references found to generate the RSGx Cable Schedule.");
                }

                ExportDataToCsvGeneric(reportData, filePath, "RSGx Cable Schedule");
                Autodesk.Revit.UI.TaskDialog.Show("Export Complete", $"RSGx Cable Schedule successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to export RSGx Cable Schedule: {ex.Message}");
            }
        }

        private void PerformExport<T>(Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName, string defaultFileName, string dataTypeName) where T : class, new()
        {
            this.Close();

            List<T> dataToExport = _pcExtensible.RecallDataFromExtensibleStorage<T>(
                _doc, schemaGuid, schemaName, fieldName, dataStorageElementName);

            if (dataToExport == null || !dataToExport.Any())
            {
                Autodesk.Revit.UI.TaskDialog.Show("No Data", $"No {dataTypeName} found to export.");
                return;
            }

            string outputFolderPath = GetOutputFolderPath();
            if (string.IsNullOrEmpty(outputFolderPath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output folder not selected.");
                return;
            }

            string filePath = System.IO.Path.Combine(outputFolderPath, defaultFileName);

            try
            {
                ExportDataToCsvGeneric(dataToExport, filePath, dataTypeName);
                Autodesk.Revit.UI.TaskDialog.Show("Export Complete", $"{dataTypeName} successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to export {dataTypeName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Prompts the user to select an output folder using the standard System.Windows.Forms.FolderBrowserDialog.
        /// </summary>
        private string GetOutputFolderPath()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a Folder to Save the Exported Reports";
                dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                dialog.ShowNewFolderButton = true;

                // Create a wrapper for the Revit window handle to properly parent the dialog
                var helper = new WindowInteropHelper(this);
                System.Windows.Forms.DialogResult result = dialog.ShowDialog(new Win32Window(helper.Handle));

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    return dialog.SelectedPath;
                }
            }
            return null; // User cancelled
        }

        /// <summary>
        /// A helper class to wrap a WPF window handle for use with WinForms dialogs.
        /// </summary>
        public class Win32Window : System.Windows.Forms.IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public Win32Window(IntPtr handle)
            {
                Handle = handle;
            }
        }

        private void ExportDataToCsvGeneric<T>(List<T> data, string filePath, string dataTypeName) where T : class
        {
            var sb = new StringBuilder();
            PropertyInfo[] allTypeProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            List<string> headers = new List<string>();
            List<PropertyInfo> orderedPropertiesToExport = new List<PropertyInfo>();

            if (typeof(T) == typeof(RSGxCableData))
            {
                headers.AddRange(new string[] {
                    "Row Number", "Cable Tag", "Origin Device (ID)", "Destination Device (ID)",
                    "RSGx Route Length (m)", "DJV Design Length (m)", "Cable Length Difference (m)",
                    "Maximum Length permissible for Cable Size (m)", "Active Cable Size (mm²)", "No. of Sets",
                    "Neutral Cable Size (mm²)", "Cores", "Conductor Type",
                    "Cable size Change from Design (Y/N)", "Previous Design Size", "Earth Included (Yes / No)",
                    "Number of Earth Cables", "Earth Size (mm2)", "Voltage", "Voltage Rating", "Type", "Sheath Construction",
                    "Insulation Construction", "Fire Rating", "Load (A)",
                    "Cable Description", "Comments", "Update Summary"
                });

                var rsgxPropertyNamesInOrder = new List<string> {
                    "RowNumber", "CableTag", "OriginDeviceID", "DestinationDeviceID",
                    "RSGxRouteLengthM", "DJVDesignLengthM", "CableLengthDifferenceM",
                    "MaxLengthPermissibleForCableSizeM", "ActiveCableSizeMM2", "NoOfSets",
                    "NeutralCableSizeMM2", "Cores", "ConductorType",
                    "CableSizeChangeFromDesignYN", "PreviousDesignSize", "EarthIncludedYesNo",
                    "NumberOfEarthCables", "EarthSizeMM2", "Voltage", "VoltageRating", "Type", "SheathConstruction",
                    "InsulationConstruction", "FireRating", "LoadA",
                    "CableDescription", "Comments", "UpdateSummary"
                };

                foreach (var propName in rsgxPropertyNamesInOrder)
                {
                    var propInfo = typeof(T).GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                    if (propInfo != null)
                    {
                        orderedPropertiesToExport.Add(propInfo);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Property '{propName}' not found in RSGxCableData class.");
                    }
                }
            }
            else
            {
                headers = allTypeProperties.Select(p => Regex.Replace(p.Name, "([a-z])([A-Z])", "$1 $2")).ToList();
                orderedPropertiesToExport = allTypeProperties.ToList();
            }

            sb.AppendLine(string.Join(",", headers));

            foreach (var item in data)
            {
                string[] values = orderedPropertiesToExport.Select(p =>
                {
                    object value = p.GetValue(item);
                    string stringValue = value?.ToString() ?? string.Empty;
                    if (stringValue.Contains(",") || stringValue.Contains("\"") || stringValue.Contains("\n") || stringValue.Contains("\r"))
                    {
                        return $"\"{stringValue.Replace("\"", "\"\"")}\"";
                    }
                    return stringValue.Trim();
                }).ToArray();
                sb.AppendLine(string.Join(",", values));
            }

            System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public class RSGxCableData
        {
            public string RowNumber { get; set; } = "";
            public string CableTag { get; set; } = "";
            public string OriginDeviceID { get; set; } = "";
            public string DestinationDeviceID { get; set; } = "";
            public string RSGxRouteLengthM { get; set; } = "";
            public string DJVDesignLengthM { get; set; } = "";
            public string CableLengthDifferenceM { get; set; } = "";
            public string MaxLengthPermissibleForCableSizeM { get; set; } = "";
            public string ActiveCableSizeMM2 { get; set; } = "";
            public string NoOfSets { get; set; } = "";
            public string NeutralCableSizeMM2 { get; set; } = "";
            public string Cores { get; set; } = "";
            public string ConductorType { get; set; } = "";
            public string CableSizeChangeFromDesignYN { get; set; } = "";
            public string PreviousDesignSize { get; set; } = "";
            public string EarthIncludedYesNo { get; set; } = "";
            public string NumberOfEarthCables { get; set; } = "";
            public string EarthSizeMM2 { get; set; } = "";
            public string Voltage { get; set; } = "";
            public string VoltageRating { get; set; } = "";
            public string Type { get; set; } = "";
            public string SheathConstruction { get; set; } = "";
            public string InsulationConstruction { get; set; } = "";
            public string FireRating { get; set; } = "";
            public string LoadA { get; set; } = "";
            public string CableDescription { get; set; } = "";
            public string Comments { get; set; } = "";
            public string UpdateSummary { get; set; } = "";
        }
    }
}
