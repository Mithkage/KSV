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
 * |         |        | - Introduced explicit ordering for RSGxCableData properties in ExportDataToCsvGeneric.
 * |         |        | - Added `orderedPropertiesToExport` to ensure CSV column order matches defined headers.
 * |         |        | - Included a check for missing properties in RSGxCableData to enhance robustness.
 * 2025-07-04 | 1.0.1   | Gemini | Added file header comments and a change log as requested.
 * |         |        | - Included detailed description of file purpose and function.
 */

#region Namespaces
using System;
using System.Collections.Generic;
// using System.IO; // REMOVED: To resolve ambiguity with System.Windows.Shapes.Path
using System.Linq;
using System.Reflection; // Required for dynamic property access (reflection)
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PC_Extensible; // IMPORTANT: Add reference to PC_Extensible project and this using directive
using System.Text.RegularExpressions; // For regex in CSV export helper

// ADDED: For WindowsAPICodePack.Dialogs
using Microsoft.WindowsAPICodePack.Dialogs;
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

            // EPPlus License setting removed as EPPlus is no longer used for this report.
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

        // Renamed from GenerateRSGxCableScheduleExcelReport_Click
        private void GenerateRSGxCableScheduleCsvReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close(); // Close the window immediately to allow user interaction for folder selection

            string outputFolderPath = GetOutputFolderPath();
            if (string.IsNullOrEmpty(outputFolderPath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output folder not selected. RSGx Cable Schedule export cancelled.");
                return;
            }

            string filePath = System.IO.Path.Combine(outputFolderPath, "RSGx Cable Schedule.csv"); // Ensures .csv extension

            try
            {
                // --- 1. RECALL DATA FROM ALL RELEVANT EXTENSIBLE STORAGES ---
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

                // --- 2. CREATE LOOKUP DICTIONARIES ---
                var primaryDict = primaryData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var consultantDict = consultantData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var modelGeneratedDict = modelGeneratedData.ToDictionary(m => m.CableReference, m => m, StringComparer.OrdinalIgnoreCase);

                // --- 3. COLLECT ALL UNIQUE CABLE REFERENCES ---
                var allCableRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cable in primaryData) allCableRefs.Add(cable.CableReference);
                foreach (var cable in consultantData) allCableRefs.Add(cable.CableReference);
                foreach (var cable in modelGeneratedData) allCableRefs.Add(cable.CableReference);

                // Remove any null/empty cable references that might have slipped in
                allCableRefs.RemoveWhere(string.IsNullOrEmpty);

                // --- 4. SORT UNIQUE CABLE REFERENCES ALPHABETICALLY ---
                var sortedUniqueCableRefs = allCableRefs.OrderBy(cr => cr).ToList();

                // --- 5. POPULATE RSGxCableData LIST ---
                var reportData = new List<RSGxCableData>();
                int rowNum = 1;

                foreach (string cableRef in sortedUniqueCableRefs)
                {
                    primaryDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData primaryInfo);
                    consultantDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData consultantInfo);
                    modelGeneratedDict.TryGetValue(cableRef, out PC_ExtensibleClass.ModelGeneratedData modelInfo);

                    // Determine From/To with Primary precedence
                    string originDeviceID = primaryInfo?.From ?? consultantInfo?.From ?? "N/A";
                    string destinationDeviceID = primaryInfo?.To ?? consultantInfo?.To ?? "N/A";

                    // Get RSGx and DJV lengths
                    string rsgxRouteLength = primaryInfo?.CableLength ?? "N/A";
                    string djvDesignLength = consultantInfo?.CableLength ?? "N/A";

                    // Calculate Cable Length Difference
                    string cableLengthDifference = "N/A";
                    if (double.TryParse(rsgxRouteLength, out double rsgxLen) && double.TryParse(djvDesignLength, out double djvLen))
                    {
                        cableLengthDifference = (rsgxLen - djvLen).ToString("F1"); // Format to 1 decimal place
                    }

                    // Calculate Cable size Change from Design (Y/N)
                    string cableSizeChangeYN = "N/A"; // Default
                    if (double.TryParse(rsgxRouteLength, out rsgxLen) && double.TryParse(djvDesignLength, out djvLen))
                    {
                        if (rsgxLen == djvLen)
                        {
                            cableSizeChangeYN = "N";
                        }
                        else
                        {
                            cableSizeChangeYN = "Y";
                        }
                    }

                    // Determine "Earth Included (Yes / No)"
                    string earthIncludedRaw = primaryInfo?.SeparateEarthForMulticore;
                    string earthIncludedFormatted;
                    if (string.Equals(earthIncludedRaw, "Yes", StringComparison.OrdinalIgnoreCase))
                    {
                        earthIncludedFormatted = "No"; // Invert: If raw is "Yes", output "No"
                    }
                    else if (string.Equals(earthIncludedRaw, "No", StringComparison.OrdinalIgnoreCase))
                    {
                        earthIncludedFormatted = "Yes"; // Invert: If raw is "No", output "Yes"
                    }
                    else
                    {
                        earthIncludedFormatted = "No"; // Default to "No" for null/empty/other values
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

                        // Mapped fields from Primary Data:
                        MaxLengthPermissibleForCableSizeM = primaryInfo?.CableMaxLengthM ?? "N/A", // From new CableMaxLengthM
                        ActiveCableSizeMM2 = primaryInfo?.ActiveCableSize ?? "N/A",
                        NoOfSets = primaryInfo?.NumberOfActiveCables ?? "N/A",
                        NeutralCableSizeMM2 = primaryInfo?.NeutralCableSize ?? "N/A",
                        No = primaryInfo?.NumberOfNeutralCables ?? "N/A", // This is the "No." under Neutral Cable Size
                        CableType = primaryInfo?.CableType ?? "N/A", // From Primary Data's CableType
                        ConductorType = primaryInfo?.ConductorActive ?? "N/A", // Conductor Type from Primary's ConductorActive
                        CableSizeChangeFromDesignYN = cableSizeChangeYN, // Populated based on logic above
                        PreviousDesignSize = consultantInfo?.ActiveCableSize ?? "N/A", // From Consultant Data's ActiveCableSize
                        EarthIncludedYesNo = earthIncludedFormatted, // Populated with inverted logic
                        EarthSizeMM2 = primaryInfo?.EarthCableSize ?? "N/A",
                        // EarthSheath property is removed from RSGxCableData class
                        Voltage = primaryInfo?.VoltageVac ?? "N/A", // From Primary Data's VoltageVac
                        VoltageRating = "0.6/1kV", // UPDATED: Set to constant value
                        Type = primaryInfo?.CableType ?? "N/A", // Reusing CableType if 'Type' refers to cable type description
                        SheathConstruction = primaryInfo?.Sheath ?? "N/A", // From Primary Data's Sheath (RSGx Sheath Construction)
                        InsulationConstruction = primaryInfo?.Insulation ?? "N/A", // From Primary Data's Insulation (RSGx Insulation Construction)
                        FireRating = "WS52W", // UPDATED: Set to constant value
                        InstallationConfiguration = "N/A", // Not yet mapped
                        CableDescription = "N/A", // Not yet mapped
                        Comments = modelInfo?.Comment ?? "N/A", // Comments from Model Generated Data
                        UpdateSummary = "N/A" // Not yet mapped
                    });
                }

                // Display message if no data to export
                if (!reportData.Any())
                {
                    Autodesk.Revit.UI.TaskDialog.Show("No Data", "No unique cable references found across all specified extensible storages to generate the RSGx Cable Schedule. The report will be empty.");
                    // Optionally, you might want to create an empty CSV with just headers.
                }

                // Call the generic CSV exporter for RSGx Cable Schedule data
                ExportDataToCsvGeneric(reportData, filePath, "RSGx Cable Schedule");
                Autodesk.Revit.UI.TaskDialog.Show("Export Complete", $"RSGx Cable Schedule successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to export RSGx Cable Schedule: {ex.Message}");
            }
        }


        /// <summary>
        /// Handles the common logic for recalling data and exporting it to CSV.
        /// </summary>
        /// <typeparam name="T">The type of data to recall and export (e.g., CableData, ModelGeneratedData).</typeparam>
        private void PerformExport<T>(
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName,
            string defaultFileName,
            string dataTypeName) where T : class, new()
        {
            this.Close(); // Close the current window to allow Revit to regain focus during folder selection/export.

            List<T> dataToExport = _pcExtensible.RecallDataFromExtensibleStorage<T>(
                _doc, schemaGuid, schemaName, fieldName, dataStorageElementName);

            if (dataToExport == null || !dataToExport.Any())
            {
                Autodesk.Revit.UI.TaskDialog.Show("No Data", $"No {dataTypeName} found in project's extensible storage to export.");
                return; // Exit as there's no data
            }

            string outputFolderPath = GetOutputFolderPath();
            if (string.IsNullOrEmpty(outputFolderPath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output folder not selected. Export operation cancelled.");
                return; // Exit if folder not selected
            }

            // Use System.IO.Path explicitly
            string filePath = System.IO.Path.Combine(outputFolderPath, defaultFileName);

            try
            {
                ExportDataToCsvGeneric(dataToExport, filePath, dataTypeName); // Call method within this class
                Autodesk.Revit.UI.TaskDialog.Show("Export Complete", $"{dataTypeName} successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to export {dataTypeName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Prompts the user to select an output folder using Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog.
        /// This provides a modern Windows dialog for folder selection.
        /// </summary>
        private string GetOutputFolderPath()
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Title = "Select a Folder to Save the Exported Reports";
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                CommonFileDialogResult result = dialog.ShowDialog();

                if (result == CommonFileDialogResult.Ok) // Check against CommonFileDialogResult.Ok
                {
                    return dialog.FileName; // FileName property contains the selected folder path when IsFolderPicker is true
                }
            }
            return null; // User cancelled
        }

        /// <summary>
        /// Exports a list of generic data objects to a CSV file dynamically.
        /// This method can now export RSGxCableData as well.
        /// </summary>
        /// <typeparam name="T">The type of objects in the list (e.g., CableData, ModelGeneratedData, RSGxCableData).</typeparam>
        /// <param name="data">The list of data objects to export.</param>
        /// <param name="filePath">The full path to the output CSV file.</param>
        /// <param name="dataTypeName">A friendly name for the data being exported (for messages).</param>
        private void ExportDataToCsvGeneric<T>(List<T> data, string filePath, string dataTypeName) where T : class
        {
            var sb = new StringBuilder();

            // Get properties of type T using reflection
            PropertyInfo[] allTypeProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            List<string> headers = new List<string>();
            List<PropertyInfo> orderedPropertiesToExport = new List<PropertyInfo>(); // This will hold the ordered properties

            if (typeof(T) == typeof(RSGxCableData))
            {
                // Explicitly define the single header row for RSGx Cable Schedule CSV
                // This combines the two Excel rows into one logical CSV header.
                headers.Add("Row Number");
                headers.Add("Cable Tag");
                headers.Add("Origin Device (ID)");
                headers.Add("Destination Device (ID)");
                headers.Add("RSGx Route Length (m)");
                headers.Add("DJV Design Length (m)");
                headers.Add("Cable Length Difference (m)");
                headers.Add("Maximum Length permissible for Cable Size (m)");
                headers.Add("Active Cable Size (mm²)");
                headers.Add("No. of Sets");
                headers.Add("Neutral Cable Size (mm²)");
                headers.Add("No.");
                headers.Add("Cable Type");
                headers.Add("Conductor Type");
                headers.Add("Cable size Change from Design (Y/N)");
                headers.Add("Previous Design Size");
                headers.Add("Earth Included (Yes / No)");
                headers.Add("Earth Size (mm2)");
                headers.Add("Voltage");
                headers.Add("Voltage Rating");
                headers.Add("Type");
                headers.Add("Sheath Construction");
                headers.Add("Insulation Construction");
                headers.Add("Fire Rating");
                headers.Add("Installation Configuration");
                headers.Add("Cable Description");
                headers.Add("Comments");
                headers.Add("Update Summary");

                // Manually order properties to match the headers for RSGxCableData
                // Ensure property names here match the actual property names in RSGxCableData class
                var rsgxPropertyNamesInOrder = new List<string>
                {
                    "RowNumber", "CableTag", "OriginDeviceID", "DestinationDeviceID",
                    "RSGxRouteLengthM", "DJVDesignLengthM", "CableLengthDifferenceM",
                    "MaxLengthPermissibleForCableSizeM", "ActiveCableSizeMM2", "NoOfSets",
                    "NeutralCableSizeMM2", "No", "CableType", "ConductorType",
                    "CableSizeChangeFromDesignYN", "PreviousDesignSize", "EarthIncludedYesNo",
                    "EarthSizeMM2", "Voltage", "VoltageRating", "Type", "SheathConstruction",
                    "InsulationConstruction", "FireRating", "InstallationConfiguration",
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
                        // This is a critical error: a property expected in the header list is missing from the class.
                        // You should log this or throw an exception during development.
                        throw new InvalidOperationException($"Property '{propName}' not found in RSGxCableData class, but listed in CSV header order.");
                    }
                }
            }
            else // For CableData, ModelGeneratedData etc., use generic header generation
            {
                // For generic types, maintain current behavior of generating headers and properties
                // based on reflection order (which might still be inconsistent for non-RSGx reports,
                // but that's a separate concern if it becomes an issue for those reports).
                headers = allTypeProperties.Select(p =>
                {
                    // Custom formatting for specific property names to make headers user-friendly
                    if (p.Name == "CableLengthM") return "Cable Length (m)";
                    if (p.Name == "CablesKgPerM") return "Cables kg per m";
                    if (p.Name == "NominalOverallDiameter") return "Nominal Overall Diameter (mm)";
                    if (p.Name == "AsNsz3008CableDeratingFactor") return "AS/NSZ 3008 Cable Derating Factor";
                    if (p.Name == "ConductorActive") return "Conductor (Active)";
                    if (p.Name == "ConductorEarth") return "Conductor (Earth)";
                    if (p.Name == "SeparateEarthForMulticore") return "Separate Earth for Multicore";
                    if (p.Name == "TotalCableRunWeight") return "Total Cable Run Weight (Incl. N & E) (kg)";
                    if (p.Name == "NumberOfActiveCables") return "Number of Active Cables";
                    if (p.Name == "ActiveCableSize") return "Active Cable Size (mm\u00B2)";
                    if (p.Name == "NumberOfNeutralCables") return "Number of Neutral Cables";
                    if (p.Name == "NeutralCableSize") return "Neutral Cable Size (mm\u00B2)";
                    if (p.Name == "NumberOfEarthCables") return "Number of Earth Cables";
                    if (p.Name == "EarthCableSize") return "Earth Cable Size (mm\u00B2)";
                    if (p.Name == "OriginDeviceID") return "Origin Device (ID)";
                    if (p.Name == "DestinationDeviceID") return "Destination Device (ID)";
                    if (p.Name == "RSGxRouteLengthM") return "RSGx Route Length (m)";
                    if (p.Name == "DJVDesignLengthM") return "DJV Design Length (m)";
                    if (p.Name == "CableLengthDifferenceM") return "Cable Length Difference (m)";
                    if (p.Name == "MaxLengthPermissibleForCableSizeM") return "Maximum Length permissible for Cable Size (m)";
                    if (p.Name == "ActiveCableSizeMM2") return "Active Cable Size (mm²)";
                    if (p.Name == "NeutralCableSizeMM2") return "Neutral Cable Size (mm²)";
                    if (p.Name == "EarthSizeMM2") return "Earth Size (mm2)";
                    if (p.Name == "EarthIncludedYesNo") return "Earth Included (Yes / No)";
                    if (p.Name == "NoOfSets") return "No. of Sets";
                    if (p.Name == "No") return "No.";
                    if (p.Name == "VoltageRating") return "Voltage Rating";
                    if (p.Name == "SheathConstruction") return "Sheath Construction";
                    if (p.Name == "InsulationConstruction") return "Insulation Construction";
                    if (p.Name == "FireRating") return "Fire Rating";
                    if (p.Name == "InstallationConfiguration") return "Installation Configuration";
                    if (p.Name == "CableDescription") return "Cable Description";
                    if (p.Name == "CableSizeChangeFromDesignYN") return "Cable size Change from Design (Y/N)";
                    if (p.Name == "PreviousDesignSize") return "Previous Design Size";
                    if (p.Name == "UpdateSummary") return "Update Summary";
                    if (p.Name == "EarthSheath") return "Earth Sheath";
                    if (p.Name == "VoltageVac") return "Voltage (Vac)";

                    return Regex.Replace(p.Name, "([a-z])([A-Z])", "$1 $2"); // Convert CamelCase to "Camel Case"
                }).ToList();

                orderedPropertiesToExport = allTypeProperties.ToList(); // For generic types, just use the reflection order
            }

            sb.AppendLine(string.Join(",", headers));

            // Generate data rows using the determined order of properties
            foreach (var item in data)
            {
                string[] values = orderedPropertiesToExport.Select(p => // Use orderedPropertiesToExport here
                {
                    object value = p.GetValue(item);
                    string stringValue = value?.ToString() ?? string.Empty;

                    // Apply CSV escaping rules
                    if (stringValue.Contains(",") || stringValue.Contains("\"") || stringValue.Contains("\n") || stringValue.Contains("\r"))
                    {
                        return $"\"{stringValue.Replace("\"", "\"\"")}\"";
                    }
                    return stringValue.Trim(); // Trim whitespace from values
                }).ToArray();
                sb.AppendLine(string.Join(",", values));
            }

            System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        // --- NEW DATA CLASS FOR RSGX CABLE SCHEDULE ---
        /// <summary>
        /// Data class to hold the structure for the RSGx Cable Schedule Excel report.
        /// Properties match the detailed column headers from the sample provided.
        /// </summary>
        public class RSGxCableData
        {
            public string RowNumber { get; set; }
            public string CableTag { get; set; }
            public string OriginDeviceID { get; set; }
            public string DestinationDeviceID { get; set; }
            public string RSGxRouteLengthM { get; set; }
            public string DJVDesignLengthM { get; set; }
            public string CableLengthDifferenceM { get; set; }
            public string MaxLengthPermissibleForCableSizeM { get; set; }
            public string ActiveCableSizeMM2 { get; set; }
            public string NoOfSets { get; set; }
            public string NeutralCableSizeMM2 { get; set; }
            public string No { get; set; }
            public string CableType { get; set; } // These 4 were implicitly single headers in row 1
            public string ConductorType { get; set; }
            public string CableSizeChangeFromDesignYN { get; set; }
            public string PreviousDesignSize { get; set; }
            public string EarthIncludedYesNo { get; set; }
            public string EarthSizeMM2 { get; set; }
            // REMOVED: public string EarthSheath { get; set; } // Removed as requested
            public string Voltage { get; set; }
            public string VoltageRating { get; set; }
            public string Type { get; set; } // Renamed from original "Type" to distinguish from System.Type
            public string SheathConstruction { get; set; }
            public string InsulationConstruction { get; set; }
            public string FireRating { get; set; }
            public string InstallationConfiguration { get; set; }
            public string CableDescription { get; set; }
            public string Comments { get; set; }
            public string UpdateSummary { get; set; }

            // Constructor to initialize with empty strings for dummy data
            public RSGxCableData()
            {
                // Initialize all properties to empty string to prevent null reference issues
                // and for cleaner Excel output when data is missing.
                foreach (PropertyInfo prop in typeof(RSGxCableData).GetProperties())
                {
                    if (prop.PropertyType == typeof(string) && prop.CanWrite)
                    {
                        prop.SetValue(this, string.Empty);
                    }
                }
            }
        }
    }
}
