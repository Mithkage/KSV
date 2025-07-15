//-----------------------------------------------------------------------------
// <copyright file="ReportSelectionWindow.xaml.cs" company="RTS Reports">
//      Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//      This file contains the interaction logic for the ReportSelectionWindow.xaml.
//      It handles user selections for various report types and orchestrates
//      the data recall from Revit's extensible storage and subsequent export
//      to CSV format, particularly for the RSGx Cable Schedule.
// </summary>
//-----------------------------------------------------------------------------

/*
 * Change Log:
 *
 * Date       | Version | Author | Description
 * ===========|=========|========|====================================================================================================
 * 2025-07-04 | 1.0.0    | Gemini | Initial implementation to address column and data mixing in RSGx report.
 * |         |        | - Introduced explicit ordering for RSGxCableData properties in ExportDataToCsvGeneric.
 * |         |        | - Added `orderedPropertiesToExport` to ensure CSV column order matches defined headers.
 * |         |        | - Included a check for missing properties in RSGxCableData to enhance robustness.
 * 2025-07-04 | 1.0.1    | Gemini | Added file header comments and a change log as requested.
 * |         |        | - Included detailed description of file purpose and function.
 * 2025-07-07 | 1.1.0    | Gemini | Replaced WindowsAPICodePack folder browser with standard System.Windows.Forms.FolderBrowserDialog.
 * |         |        | - This removes the final dependency on the conflicting library to resolve build and runtime errors.
 * |         |        | - Added a Win32Window helper class to properly parent the dialog to the Revit window.
 * 2025-07-07 | 1.1.1    | Gemini | Resolved ambiguous reference error for IWin32Window by specifying the System.Windows.Forms namespace.
 * 2025-07-08 | 1.2.0    | Gemini | Updated RSGx Cable Schedule generation logic based on user feedback.
 * |         |        | - Relabeled 'Cable Type' header to 'Cores' and mapped it to the 'Cores' data field.
 * |         |        | - Implemented conditional logic for 'Fire Rating' based on the cable's 'Type' property.
 * |         |        | - Implemented conditional logic for 'Update Summary' based on cable length differences.
 * |         |        | - Removed 'Installation Configuration' column from the report.
 * 2025-07-08 | 1.3.0    | Gemini | Enhanced RSGx report data formatting.
 * |         |        | - 'Maximum Length' is now rounded down to the nearest integer.
 * |         |        | - 'Cable Description' is now a concatenation of key cable properties, omitting nulls.
 * 2025-07-08 | 1.4.0    | Gemini | Refined RSGx report data formatting.
 * |         |        | - Appended "mm²" unit to 'Active Cable Size' in the 'Cable Description'.
 * |         |        | - Set 'Voltage Rating' to null if 'Destination Device (ID)' is not available.
 * 2025-07-08 | 1.5.0    | Gemini | Added "Load (A)" column to the RSGx Cable Schedule report.
 * |         |        | - Data is sourced from the new LoadA property in the primary data store.
 * 2025-07-08 | 1.6.0    | Gemini | Added Maximum Demand (MD) comparison to the "Update Summary" field.
 * |         |        | - Compares "Load (A)" between primary and consultant data.
 * 2025-07-08 | 1.7.0    | Gemini | Added "Number of Earth Cables" column to the RSGx report.
 * |         |        | - Column inserted before "Earth Size (mm2)".
 * 2025-07-08 | 1.8.0    | Gemini | Removed the "No." column (related to neutral cables) from the RSGx report.
 * 2025-07-10 | 1.9.0    | Gemini | Updated RSGx report sorting logic.
 * |         |        | - Removed the general alphabetical re-ordering of all cable references.
 * |         |        | - Report now sorts primarily by Cable Reference order from Primary Storage.
 * |         |        | - Cable References found only in Consultant Storage (or Model Generated)
 * |         |        |   are moved to the end, sorted alphabetically by Cable Reference.
 * 2025-07-10 | 2.0.0    | Gemini | Refined RSGx report data generation logic.
 * |         |        | - Removed actual 'Comments' data, keeping the column as empty.
 * |         |        | - Implemented conditional logic: if 'Cores' is "N/A", then 'Voltage Rating'
 * |         |        |   and 'Cable Description' are also set to "N/A".
 * 2025-07-10 | 2.0.1    | Gemini | Fixed CS0103 error for 'activeCableSize' by declaring it in a broader scope.
 * |         |        | - Also applied this fix to 'type', 'sheath', and 'insulation' for consistency.
 * 2025-07-10 | 2.1.0    | Gemini | Added new button handler for "RSGx Cable Summary (XLSX)" report.
 * |         |        | - Integrated call to `PC_Generate_MDClass.Execute` for Excel generation.
 * 2025-07-10 | 2.1.1    | Gemini | Corrected `using` directive for `PC_Generate_MDClass`.
 * |         |        | - Changed `using RTS;` to `using PC_Generate_MD;` to resolve CS0246 error.
 * 2025-07-10 | 2.1.2    | Gemini | Corrected instantiation of PC_Generate_MDClass.
 * |         |        | - Removed incorrect `using PC_Generate_MD;` directive.
 * |         |        | - Changed `new PC_Generate_MD.PC_Generate_MDClass()` to `new PC_Generate_MDClass()`.
 * 2025-07-10 | 2.2.0    | Gemini | Updated "Cable size Change from Design (Y/N)" logic for RSGx report.
 * |         |        | - Now compares Primary "Active Cable Size (mm²)" with Consultant "Active Cable Size (mm²)".
 * |         |        | - Renamed "Previous Design Size" column to "DJV Active Cable Size (mm²)" and reordered.
 * 2025-07-10 | 2.3.0    | Gemini | Further refined RSGx report column order and naming.
 * |         |        | - Moved "DJV Active Cable Size (mm²)" to appear after "RSGx Active Cable Size (mm²)".
 * |         |        | - Renamed "Active Cable Size (mm²)" to "RSGx Active Cable Size (mm²)".
 * 2025-07-10 | 2.4.0    | Gemini | Reordered RSGx report columns.
 * |         |        | - Moved "Cable size Change from Design (Y/N)" to appear after "DJV Active Cable Size (mm²)".
 * 2025-07-10 | 2.5.0    | Gemini | Updated "Earth Included (Yes / No)" column logic.
 * |         |        | - Now outputs "Y" or "N" instead of "Yes" or "No".
 * 2025-07-10 | 2.6.0    | Gemini | Renamed "Earth Included (Yes / No)" column header to "Earth Included (Y/N)".
 * 2025-07-10 | 2.7.0    | Gemini | Updated "Type" column in RSGx report.
 * |         |        | - Appended ", LSZH" if the value is not "N/A".
 * 2025-07-15 | 2.8.0    | Gemini | Updated namespace from 'RTS_Reports' to 'RTS.UI' to align with project structure.
 * 2025-07-15 | 2.9.0    | Gemini | Corrected using directive for PC_ExtensibleClass to match its actual namespace.
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
using System.Text.RegularExpressions;
using System.Windows.Forms; // For FolderBrowserDialog
using System.Windows.Interop; // For WindowInteropHelper
using RTS.Commands; // Corrected using directive for PC_Generate_MDClass
using PC_Extensible; // UPDATED: Corrected using directive for PC_ExtensibleClass
#endregion

namespace RTS.UI // UPDATED: Changed from 'RTS_Reports' to 'RTS.UI'
{
    /// <summary>
    /// Interaction logic for ReportSelectionWindow.xaml
    /// This class now also contains the logic for generating reports.
    /// </summary>
    public partial class ReportSelectionWindow : Window
    {
        private Document _doc;
        private PC_ExtensibleClass _pcExtensible;
        private ExternalCommandData _commandData; // Stored to pass to Excel generation command

        public ReportSelectionWindow(ExternalCommandData commandData)
        {
            InitializeComponent();
            _doc = commandData.Application.ActiveUIDocument.Document;
            _pcExtensible = new PC_ExtensibleClass();
            _commandData = commandData; // Store the commandData
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

                // --- START New Sorting Logic for RSGx Cable Schedule ---
                var finalReportCableRefs = new List<string>();
                var processedCableRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // 1. Add cable references from primaryData, maintaining their retrieved order.
                // This ensures "Report by Cable Reference to match Primary Storage Cable Reference order".
                foreach (var primaryCable in primaryData)
                {
                    if (!string.IsNullOrEmpty(primaryCable.CableReference) && processedCableRefs.Add(primaryCable.CableReference))
                    {
                        finalReportCableRefs.Add(primaryCable.CableReference);
                    }
                }

                // 2. Add cable references from consultantData that are NOT in primaryData, sorted alphabetically.
                // These are "Cable References present in Consultant Storage but not in Primary Storage"
                // and are "to be moved to the end after the cables are sorted by Primary Data Cable Reference".
                var consultantOnlyRefs = consultantData
                    .Where(c => !string.IsNullOrEmpty(c.CableReference) && !processedCableRefs.Contains(c.CableReference))
                    .Select(c => c.CableReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(cr => cr, StringComparer.OrdinalIgnoreCase) // Sort consultant-only cables alphabetically
                    .ToList();

                foreach (var conRef in consultantOnlyRefs)
                {
                    if (processedCableRefs.Add(conRef))
                    {
                        finalReportCableRefs.Add(conRef);
                    }
                }

                // 3. Add cable references from modelGeneratedData that are NOT in primaryData or consultantData, sorted alphabetically.
                // This ensures any model-generated only cables are also included at the very end.
                var modelGeneratedOnlyRefs = modelGeneratedData
                    .Where(m => !string.IsNullOrEmpty(m.CableReference) && !processedCableRefs.Contains(m.CableReference))
                    .Select(m => m.CableReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(cr => cr, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var modelRef in modelGeneratedOnlyRefs)
                {
                    if (processedCableRefs.Add(modelRef))
                    {
                        finalReportCableRefs.Add(modelRef);
                    }
                }
                // --- END New Sorting Logic ---

                var reportData = new List<RSGxCableData>();
                int rowNum = 1;

                foreach (string cableRef in finalReportCableRefs) // Use the newly sorted/ordered list
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

                    // --- START Updated Logic for Cable size Change from Design (Y/N) ---
                    string cableSizeChangeYN;
                    string primaryActiveSize = primaryInfo?.ActiveCableSize;
                    string consultantActiveSize = consultantInfo?.ActiveCableSize;

                    if (string.IsNullOrEmpty(primaryActiveSize) || primaryActiveSize == "N/A" ||
                        string.IsNullOrEmpty(consultantActiveSize) || consultantActiveSize == "N/A")
                    {
                        cableSizeChangeYN = "N/A";
                    }
                    else
                    {
                        cableSizeChangeYN = (string.Equals(primaryActiveSize, consultantActiveSize, StringComparison.OrdinalIgnoreCase)) ? "N" : "Y";
                    }
                    // --- END Updated Logic for Cable size Change from Design (Y/N) ---

                    string earthIncludedRaw = primaryInfo?.SeparateEarthForMulticore;
                    // Updated "Earth Included (Yes / No)" to be "Y" or "N"
                    string earthIncludedFormatted = string.Equals(earthIncludedRaw, "No", StringComparison.OrdinalIgnoreCase) ? "Y" : "N";

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

                    string voltageRating;
                    string cableDescriptionValue;
                    string cores = primaryInfo?.Cores ?? "N/A"; // Get cores here for the conditional logic

                    // Declare these variables outside the if/else to ensure they are always in scope
                    string activeCableSize = primaryInfo?.ActiveCableSize ?? "N/A";
                    string type = primaryInfo?.CableType ?? "N/A";
                    string sheath = primaryInfo?.Sheath ?? "N/A";
                    string insulation = primaryInfo?.Insulation ?? "N/A";

                    // Conditional logic for Voltage Rating and Cable Description based on Cores
                    if (cores == "N/A")
                    {
                        voltageRating = "N/A";
                        cableDescriptionValue = "N/A";
                    }
                    else
                    {
                        voltageRating = "0.6/1kV";
                        if (destinationDeviceID == "N/A")
                        {
                            voltageRating = ""; // Set to null/empty if destination device is N/A
                        }

                        var descriptionParts = new List<string>();

                        if (!string.IsNullOrEmpty(activeCableSize) && activeCableSize != "N/A") descriptionParts.Add($"{activeCableSize}mm²");
                        if (!string.IsNullOrEmpty(cores) && cores != "N/A") descriptionParts.Add(cores);
                        if (!string.IsNullOrEmpty(voltageRating) && voltageRating != "N/A") descriptionParts.Add(voltageRating);
                        if (!string.IsNullOrEmpty(type) && type != "N/A") descriptionParts.Add(type);
                        if (!string.IsNullOrEmpty(sheath) && sheath != "N/A") descriptionParts.Add(sheath);
                        if (!string.IsNullOrEmpty(insulation) && insulation != "N/A") descriptionParts.Add(insulation);
                        if (!string.IsNullOrEmpty(fireRatingValue) && fireRatingValue != "N/A") descriptionParts.Add(fireRatingValue);
                        cableDescriptionValue = string.Join(" | ", descriptionParts);
                    }

                    // Append ", LSZH" to Type if it's not "N/A"
                    if (!string.IsNullOrEmpty(type) && !string.Equals(type, "N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        type += ", LSZH";
                    }


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
                        PreviousDesignSize = consultantInfo?.ActiveCableSize ?? "N/A", // This is now "DJV Active Cable Size (mm²)"
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
                        Comments = "", // Set Comments to empty string
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

        /// <summary>
        /// Handles the click event for the "RSGx Cable Summary (XLSX)" button.
        /// Initiates the generation of the Excel report using PC_Generate_MDClass.
        /// </summary>
        private void GenerateRSGxCableSummaryXlsxReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close(); // Close the selection window

            string message = "";
            // ElementSet is typically used for selected elements, but for a general report, an empty one is fine.
            ElementSet elements = new ElementSet();

            // Corrected instantiation: PC_Generate_MDClass is directly accessible via 'using RTS.Commands;'
            PC_Generate_MDClass excelGenerator = new PC_Generate_MDClass();
            Result result = excelGenerator.Execute(_commandData, ref message, elements);

            // The PC_Generate_MDClass.Execute method already handles success/failure messages via MessageBox.Show
            // However, if it returns Result.Failed, we can log or show an additional message if needed.
            if (result == Result.Failed)
            {
                // This TaskDialog will only show if the Excel generator itself didn't show a more specific error.
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to generate RSGx Cable Summary (XLSX): {message}");
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
                    "Maximum Length permissible for Cable Size (m)",
                    "RSGx Active Cable Size (mm²)",
                    "DJV Active Cable Size (mm²)",
                    "Cable size Change from Design (Y/N)",
                    "No. of Sets",
                    "Neutral Cable Size (mm²)", "Cores", "Conductor Type",
                    "Earth Included (Y/N)",
                    "Number of Earth Cables", "Earth Size (mm2)", "Voltage", "Voltage Rating", "Type", "Sheath Construction",
                    "Insulation Construction", "Fire Rating", "Load (A)",
                    "Cable Description", "Comments", "Update Summary"
                });

                var rsgxPropertyNamesInOrder = new List<string> {
                    "RowNumber", "CableTag", "OriginDeviceID", "DestinationDeviceID",
                    "RSGxRouteLengthM", "DJVDesignLengthM", "CableLengthDifferenceM",
                    "MaxLengthPermissibleForCableSizeM",
                    "ActiveCableSizeMM2",
                    "PreviousDesignSize",
                    "CableSizeChangeFromDesignYN",
                    "NoOfSets",
                    "NeutralCableSizeMM2", "Cores", "ConductorType",
                    "EarthIncludedYesNo",
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
