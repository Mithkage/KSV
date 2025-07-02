// ReportSelectionWindow.xaml.cs
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
using System.Windows.Forms; // ADDED: For FolderBrowserDialog (Windows Forms)

// OfficeOpenXml references are entirely removed
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
                TaskDialog.Show("Export Cancelled", "Output folder not selected. RSGx Cable Schedule export cancelled.");
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
                    string djvDesignLength = consultantInfo?.CableLength ?? "N/A"; // Changed to use consultantInfo?.CableLength

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


                    reportData.Add(new RSGxCableData
                    {
                        RowNumber = (rowNum++).ToString(),
                        CableTag = cableRef,
                        OriginDeviceID = originDeviceID,
                        DestinationDeviceID = destinationDeviceID,
                        RSGxRouteLengthM = rsgxRouteLength,
                        DJVDesignLengthM = djvDesignLength,
                        CableLengthDifferenceM = cableLengthDifference,
                        Comments = modelInfo?.Comment ?? "N/A", // Comments from Model Generated Data
                        CableSizeChangeFromDesignYN = cableSizeChangeYN, // Populated based on new logic

                        // Remaining fields are "N/A" as per current instruction unless specified otherwise
                        MaxLengthPermissibleForCableSizeM = "N/A",
                        ActiveCableSizeMM2 = "N/A",
                        NoOfSets = "N/A",
                        NeutralCableSizeMM2 = "N/A",
                        No = "N/A",
                        EarthIncludedYesNo = "N/A",
                        EarthSizeMM2 = "N/A",
                        Voltage = "N/A",
                        VoltageRating = "N/A",
                        Type = "N/A",
                        SheathConstruction = "N/A",
                        InsulationConstruction = "N/A",
                        FireRating = "N/A",
                        InstallationConfiguration = "N/A",
                        CableDescription = "N/A",
                        CableType = "N/A",
                        ConductorType = "N/A",
                        PreviousDesignSize = "N/A",
                        UpdateSummary = "N/A"
                    });
                }

                // Display message if no data to export
                if (!reportData.Any())
                {
                    TaskDialog.Show("No Data", "No unique cable references found across all specified extensible storages to generate the RSGx Cable Schedule. The report will be empty.");
                    // Optionally, you might want to create an empty CSV with just headers.
                }

                // Call the generic CSV exporter for RSGx Cable Schedule data
                ExportDataToCsvGeneric(reportData, filePath, "RSGx Cable Schedule");
                TaskDialog.Show("Export Complete", $"RSGx Cable Schedule successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Failed to export RSGx Cable Schedule: {ex.Message}");
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
                TaskDialog.Show("No Data", $"No {dataTypeName} found in project's extensible storage to export.");
                return; // Exit as there's no data
            }

            string outputFolderPath = GetOutputFolderPath();
            if (string.IsNullOrEmpty(outputFolderPath))
            {
                TaskDialog.Show("Export Cancelled", "Output folder not selected. Export operation cancelled.");
                return; // Exit if folder not selected
            }

            // Use System.IO.Path explicitly
            string filePath = System.IO.Path.Combine(outputFolderPath, defaultFileName);

            try
            {
                ExportDataToCsvGeneric(dataToExport, filePath, dataTypeName); // Call method within this class
                TaskDialog.Show("Export Complete", $"{dataTypeName} successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Failed to export {dataTypeName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Prompts the user to select an output folder.
        /// </summary>
        private string GetOutputFolderPath()
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select a Folder to Save the Exported Reports";
                folderBrowserDialog.ShowNewFolderButton = true;
                folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // Default to My Documents

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return null;
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
            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Generate headers from property names
            // For RSGxCableData, use a specific header definition that flattens the two-row Excel concept.
            List<string> headers = new List<string>();
            if (typeof(T) == typeof(RSGxCableData))
            {
                // Explicitly define the single header row for RSGx Cable Schedule CSV
                // This combines the two Excel rows into one logical CSV header.
                headers.Add("Row Number"); headers.Add("Cable Tag");
                headers.Add("Origin Device (ID)"); headers.Add("Destination Device (ID)");
                headers.Add("RSGx Route Length (m)"); headers.Add("DJV Design Length (m)");
                headers.Add("Cable Length Difference (m)"); headers.Add("Maximum Length permissible for Cable Size (m)");
                headers.Add("Active Cable Size (mm²)"); headers.Add("No. of Sets");
                headers.Add("Neutral Cable Size (mm²)"); headers.Add("No.");
                headers.Add("Earth Included (Yes / No)"); headers.Add("Earth Size (mm2)");
                headers.Add("Voltage"); headers.Add("Voltage Rating");
                headers.Add("Type"); headers.Add("Sheath Construction");
                headers.Add("Insulation Construction"); headers.Add("Fire Rating");
                headers.Add("Installation Configuration"); headers.Add("Cable Description");
                headers.Add("Comments");
                headers.Add("Cable Type"); headers.Add("Conductor Type");
                headers.Add("Cable size Change from Design (Y/N)"); headers.Add("Previous Design Size");
                headers.Add("Update Summary");
            }
            else // For CableData, ModelGeneratedData etc., use generic header generation
            {
                headers = properties.Select(p =>
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

                    return Regex.Replace(p.Name, "([a-z])([A-Z])", "$1 $2"); // Convert CamelCase to "Camel Case"
                }).ToList();
            }

            sb.AppendLine(string.Join(",", headers));

            // Generate data rows
            foreach (var item in data)
            {
                string[] values = properties.Select(p =>
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
            public string EarthIncludedYesNo { get; set; }
            public string EarthSizeMM2 { get; set; }
            public string Voltage { get; set; }
            public string VoltageRating { get; set; }
            public string Type { get; set; } // Renamed from original "Type" to distinguish from System.Type
            public string SheathConstruction { get; set; }
            public string InsulationConstruction { get; set; }
            public string FireRating { get; set; }
            public string InstallationConfiguration { get; set; }
            public string CableDescription { get; set; }
            public string Comments { get; set; }
            public string CableType { get; set; } // These 4 were implicitly single headers in row 1
            public string ConductorType { get; set; }
            public string CableSizeChangeFromDesignYN { get; set; }
            public string PreviousDesignSize { get; set; }
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