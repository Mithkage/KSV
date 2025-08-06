//-----------------------------------------------------------------------------
// <copyright file="ReportSelectionWindow.xaml.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     This file contains the interaction logic for the ReportSelectionWindow.xaml.
//     It handles user selections for various report types and orchestrates
//     the data recall from Revit's extensible storage and subsequent export
//     to CSV format, particularly for the RSGx Cable Schedule.
//     UPDATED: Implements right-click context menu for File Name column.
// </summary>
//-----------------------------------------------------------------------------

/*
 * Change Log:
 *
 * Date         | Version | Author | Description
 * =============|=========|========|====================================================================================================
 * ... (Previous Change Log Entries) ...
 * 2025-08-06 | 4.3.0   | Gemini | Enhanced Routing Sequence Report logic.
 * |         |         |        | - Updated endpoint search to match 'From' field on 'Panel Name' and 'RTS_ID'.
 * |         |         |        | - Status column now differentiates between fully and partially confirmed routes.
 * |         |         |        | - Partially confirmed routes now map all available disconnected segments.
 * 2025-08-06 | 4.3.1   | Gemini | Refactored Routing Sequence logic to prevent KeyNotFoundException.
 * |         |         |        | - Changed adjacency graph and pathfinding to use ElementId instead of Element as dictionary keys for improved stability.
 * 2025-08-06 | 4.4.0   | Gemini | Added logic to handle duplicate equipment names/IDs.
 * |         |         |        | - FindMatchingEquipment now returns all potential candidates, prioritizing RTS_ID over Panel Name.
 * |         |         |        | - When duplicates are found, the system now calculates all possible routes and selects the one with the greatest physical length.
 * 2025-08-06 | 4.5.0   | Gemini | Added Branch Sequencing column and logic.
 * |         |         |        | - New column "Branch Sequencing" added to the report.
 * |         |         |        | - Extracts unique branch numbers from the final routing sequence and formats them into a comma-separated string.
 * 2025-08-06 | 4.5.1   | Gemini | Updated Branch Sequencing logic.
 * |         |         |        | - Branch number is now the last 4 characters of the RTS_ID.
 * |         |         |        | - Separator updated to ", " for improved readability.
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
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;
#endregion

namespace RTS.UI
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

        // Assume you have a collection of LinkViewModel for your grid
        public List<ReportLinkViewModel> Links { get; set; }

        public ReportSelectionWindow(ExternalCommandData commandData)
        {
            InitializeComponent();
            _doc = commandData.Application.ActiveUIDocument.Document;
            _pcExtensible = new PC_ExtensibleClass();
            _commandData = commandData; // Store the commandData
            Links = new List<ReportLinkViewModel>();
            // TODO: Load your Links collection here
        }

        // --- Context Menu Implementation for File Name Column ---

        // Attach this handler to your DataGrid's PreviewMouseRightButtonDown event
        private void DataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var dataGrid = sender as System.Windows.Controls.DataGrid;
            var depObj = e.OriginalSource as DependencyObject;
            var cell = FindParent<System.Windows.Controls.DataGridCell>(depObj);
            if (cell == null) return;

            var column = cell.Column;
            if (column == null || (column.Header?.ToString() ?? "") != "File Name") return;

            var row = FindParent<DataGridRow>(cell);
            if (row == null) return;

            var link = row.Item as ReportLinkViewModel;
            if (link == null) return;

            var menu = new System.Windows.Controls.ContextMenu();

            var unloadItem = new System.Windows.Controls.MenuItem
            {
                Header = "Unload",
                IsEnabled = link.IsLoaded && !link.IsPlaceholder
            };
            unloadItem.Click += (s, args) => UnloadLink(link);
            menu.Items.Add(unloadItem);

            var reloadItem = new System.Windows.Controls.MenuItem
            {
                Header = "Reload",
                IsEnabled = link.IsLoaded
            };
            reloadItem.Click += (s, args) => ReloadLink(link);
            menu.Items.Add(reloadItem);

            var reloadFromItem = new System.Windows.Controls.MenuItem
            {
                Header = "Reload From"
            };
            reloadFromItem.Click += (s, args) => ReloadFromLink(link);
            menu.Items.Add(reloadFromItem);

            var openLocationItem = new System.Windows.Controls.MenuItem
            {
                Header = "Open Location"
            };
            openLocationItem.Click += (s, args) => OpenLocation(link);
            menu.Items.Add(openLocationItem);

            var removePlaceholderItem = new System.Windows.Controls.MenuItem
            {
                Header = "Remove Placeholder",
                IsEnabled = link.IsPlaceholder
            };
            removePlaceholderItem.Click += (s, args) => RemovePlaceholder(link);
            menu.Items.Add(removePlaceholderItem);

            cell.ContextMenu = menu;
            menu.IsOpen = true;
            e.Handled = true;
        }

        // --- Action Logic ---

        private void UnloadLink(ReportLinkViewModel link)
        {
            if (!link.IsLoaded || link.IsPlaceholder) return;
            using (var tx = new Transaction(_doc, $"Unload Link {link.FileName}"))
            {
                tx.Start();
                try
                {
                    var linkType = FindRevitLinkType(link.FileName);
                    if (linkType != null)
                    {
                        linkType.Unload(null);
                        tx.Commit();
                        link.IsLoaded = false;
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    ShowError($"Failed to unload link: {ex.Message}");
                }
            }
        }

        private void ReloadLink(ReportLinkViewModel link)
        {
            if (!link.IsLoaded) return;
            using (var tx = new Transaction(_doc, $"Reload Link {link.FileName}"))
            {
                tx.Start();
                try
                {
                    var linkType = FindRevitLinkType(link.FileName);
                    if (linkType != null)
                    {
                        linkType.Reload();
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    ShowError($"Failed to reload link: {ex.Message}");
                }
            }
        }

        private void ReloadFromLink(ReportLinkViewModel link)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select New Link Source",
                Filter = "Revit Files (*.rvt)|*.rvt|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                string newPath = dlg.FileName;
                using (var tx = new Transaction(_doc, $"Reload From {link.FileName}"))
                {
                    tx.Start();
                    try
                    {
                        var linkType = FindRevitLinkType(link.FileName);
                        if (linkType != null)
                        {
                            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath);
                            linkType.LoadFrom(modelPath, new WorksetConfiguration());
                        }
                        tx.Commit();
                        link.FilePath = newPath;
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        ShowError($"Failed to reload from new source: {ex.Message}");
                    }
                }
            }
        }

        private void OpenLocation(ReportLinkViewModel link)
        {
            if (!string.IsNullOrEmpty(link.FilePath) && System.IO.File.Exists(link.FilePath))
            {
                Process.Start("explorer.exe", $"/select,\"{link.FilePath}\"");
            }
            else
            {
                ShowError("File path not found or file does not exist.");
            }
        }

        private void RemovePlaceholder(ReportLinkViewModel link)
        {
            if (!link.IsPlaceholder) return;
            Links.Remove(link);
        }

        // --- Helper Methods ---

        private RevitLinkType FindRevitLinkType(string fileName)
        {
            var allLinkTypes = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkType))
                .Cast<RevitLinkType>();
            foreach (var linkType in allLinkTypes)
            {
                var extRef = linkType.GetExternalFileReference();
                if (extRef != null)
                {
                    var path = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetPath());
                    if (System.IO.Path.GetFileName(path).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        return linkType;
                }
            }
            return null;
        }

        private void ShowError(string message)
        {
            Autodesk.Revit.UI.TaskDialog.Show("Error", message);
        }

        public static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = System.Windows.Media.VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindParent<T>(parentObject);
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

            string filePath = GetOutputFilePath("RSGx Cable Schedule.csv", "Save RSGx Cable Schedule Report");
            if (string.IsNullOrEmpty(filePath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output file not selected. RSGx Cable Schedule export cancelled.");
                return;
            }

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

                var finalReportCableRefs = new List<string>();
                var processedCableRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var primaryCable in primaryData)
                {
                    if (!string.IsNullOrEmpty(primaryCable.CableReference) && processedCableRefs.Add(primaryCable.CableReference))
                    {
                        finalReportCableRefs.Add(primaryCable.CableReference);
                    }
                }

                var consultantOnlyRefs = consultantData
                    .Where(c => !string.IsNullOrEmpty(c.CableReference) && !processedCableRefs.Contains(c.CableReference))
                    .Select(c => c.CableReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(cr => cr, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var conRef in consultantOnlyRefs)
                {
                    if (processedCableRefs.Add(conRef))
                    {
                        finalReportCableRefs.Add(conRef);
                    }
                }

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

                var reportData = new List<RSGxCableData>();
                int rowNum = 1;

                foreach (string cableRef in finalReportCableRefs)
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

                    string earthIncludedRaw = primaryInfo?.SeparateEarthForMulticore;
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
                    string cores = primaryInfo?.Cores ?? "N/A";

                    string activeCableSize = primaryInfo?.ActiveCableSize ?? "N/A";
                    string type = primaryInfo?.CableType ?? "N/A";
                    string sheath = primaryInfo?.Sheath ?? "N/A";
                    string insulation = primaryInfo?.Insulation ?? "N/A";

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
                            voltageRating = "";
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
                        Comments = "",
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

        private void GenerateRSGxCableSummaryXlsxReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

            string message = "";
            ElementSet elements = new ElementSet();

            PC_Generate_MDClass excelGenerator = new PC_Generate_MDClass();
            Result result = excelGenerator.Execute(_commandData, ref message, elements);

            if (result == Result.Failed)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to generate RSGx Cable Summary (XLSX): {message}");
            }
        }

        #region --- Routing Sequence Report (Handles Duplicates) ---

        private void GenerateRoutingSequenceReportButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

            string filePath = GetOutputFilePath("Routing_Sequence.csv", "Save Routing Sequence Report");
            if (string.IsNullOrEmpty(filePath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output file not selected. Routing Sequence export cancelled.");
                return;
            }

            var projectInfo = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_ProjectInformation)
                .WhereElementIsNotElementType()
                .Cast<ProjectInfo>()
                .FirstOrDefault();

            string projectName = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NAME)?.AsString() ?? "N/A";
            string projectNumber = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NUMBER)?.AsString() ?? "N/A";

            var sb = new StringBuilder();
            sb.AppendLine("Cable Routing Sequence");
            sb.AppendLine($"Project:,{projectName}");
            sb.AppendLine($"Project No:,{projectNumber}");
            sb.AppendLine($"Date:,{DateTime.Now:dd/MM/yyyy}");
            sb.AppendLine();
            sb.AppendLine("Separator Legend:,");
            sb.AppendLine(",, (Comma) = Direct connection between same containment types");
            sb.AppendLine(",, + (Plus) = Transition between different containment types");
            sb.AppendLine(",, || (Double Pipe) = Jump between non-connected containment of the same type");
            sb.AppendLine(",, >> (Double Arrow) = Separates disconnected routing segments");
            sb.AppendLine();
            sb.AppendLine("Cable Reference,From,To,Status,Branch Sequencing,Routing Sequence");

            List<PC_ExtensibleClass.CableData> primaryData = _pcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                _doc, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName);

            if (primaryData == null || !primaryData.Any())
            {
                Autodesk.Revit.UI.TaskDialog.Show("No Data", "No primary cable data found.");
                return;
            }

            var cableLookup = primaryData
                .Where(c => !string.IsNullOrWhiteSpace(c.CableReference))
                .GroupBy(c => c.CableReference)
                .ToDictionary(g => g.Key, g => g.First());

            var rtsIdGuid = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");

            var allConduits = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Conduit)
                .WhereElementIsNotElementType()
                .Where(e => e.get_Parameter(rtsIdGuid)?.HasValue ?? false)
                .ToList();

            var allCableTrays = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_CableTray)
                .WhereElementIsNotElementType()
                .Where(e => e.get_Parameter(rtsIdGuid)?.HasValue ?? false)
                .ToList();

            var allContainmentElements = allConduits.Concat(allCableTrays).ToList();

            var allEquipment = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .WhereElementIsNotElementType().ToList();

            var allFixtures = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
                .WhereElementIsNotElementType().ToList();

            var allCandidateEquipment = allEquipment.Concat(allFixtures).ToList();

            var cableGuids = new List<Guid>
            {
                new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), // Cable_01...
                new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"),
                new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"),
                new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"),
                new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"),
                new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"),
                new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"),
                new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"),
                new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"),
                new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"),
                new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"),
                new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"),
                new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"),
                new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"),
                new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"),
                new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b")
            };

            var progressWindow = new RoutingReportProgressBarWindow
            {
                Owner = System.Windows.Application.Current?.MainWindow
            };
            progressWindow.Show();

            try
            {
                int totalCables = cableLookup.Count;
                int currentCable = 0;

                foreach (var cableRef in cableLookup.Keys)
                {
                    currentCable++;
                    var cableInfo = cableLookup[cableRef];

                    progressWindow.UpdateProgress(currentCable, totalCables, cableRef,
                        cableInfo?.From ?? "N/A", cableInfo?.To ?? "N/A",
                        (double)currentCable / totalCables * 100.0, "Building routing sequence...");
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);

                    if (progressWindow.IsCancelled)
                    {
                        progressWindow.TaskDescriptionText.Text = "Operation cancelled by user.";
                        break;
                    }

                    string status = "Error";
                    string routingSequence = "Could not determine route.";
                    string branchSequence = "N/A";
                    List<ElementId> bestPath = null;

                    try
                    {
                        var cableContainmentElements = new List<Element>();
                        foreach (Element elem in allContainmentElements)
                        {
                            foreach (var guid in cableGuids)
                            {
                                Parameter cableParam = elem.get_Parameter(guid);
                                string paramValue = cableParam?.AsString();
                                if (string.Equals(CleanCableReference(paramValue), cableRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    cableContainmentElements.Add(elem);
                                    break;
                                }
                            }
                        }

                        if (!cableContainmentElements.Any())
                        {
                            status = "Pending";
                            routingSequence = "No cable containment assigned in model.";
                        }
                        else
                        {
                            var rtsIdToElementMap = new Dictionary<string, Element>();
                            foreach (var elem in cableContainmentElements)
                            {
                                string rtsId = elem.get_Parameter(rtsIdGuid)?.AsString();
                                if (!string.IsNullOrWhiteSpace(rtsId) && !rtsIdToElementMap.ContainsKey(rtsId))
                                {
                                    rtsIdToElementMap.Add(rtsId, elem);
                                }
                            }

                            var adjacencyGraph = BuildAdjacencyGraph(cableContainmentElements);

                            List<Element> startCandidates = FindMatchingEquipment(cableInfo.From, allCandidateEquipment, rtsIdGuid);
                            List<Element> endCandidates = FindMatchingEquipment(cableInfo.To, allCandidateEquipment, rtsIdGuid);

                            if (startCandidates.Any() && endCandidates.Any())
                            {
                                double maxLength = -1.0;

                                foreach (var startCandidate in startCandidates)
                                {
                                    foreach (var endCandidate in endCandidates)
                                    {
                                        var currentPath = FindConfirmedPath(startCandidate, endCandidate, cableContainmentElements, adjacencyGraph);
                                        if (currentPath != null && currentPath.Any())
                                        {
                                            double currentLength = GetPathLength(currentPath);
                                            if (currentLength > maxLength)
                                            {
                                                maxLength = currentLength;
                                                bestPath = currentPath;
                                            }
                                        }
                                    }
                                }

                                if (bestPath != null)
                                {
                                    status = "Confirmed";
                                    routingSequence = FormatPath(bestPath, adjacencyGraph, rtsIdToElementMap);
                                }
                                else
                                {
                                    status = "Unconfirmed";
                                    var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                    routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                                }
                            }
                            else if (startCandidates.Any())
                            {
                                status = "Confirmed (From)";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }
                            else if (endCandidates.Any())
                            {
                                status = "Confirmed (To)";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }
                            else
                            {
                                status = "Unconfirmed";
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                routingSequence = FormatDisconnectedPaths(disconnectedPaths, adjacencyGraph, rtsIdToElementMap);
                            }

                            if (bestPath != null && bestPath.Any())
                            {
                                branchSequence = FormatBranchSequence(bestPath, rtsIdToElementMap);
                            }
                            else if (status.Contains("Unconfirmed") || status.Contains("Confirmed (From)") || status.Contains("Confirmed (To)"))
                            {
                                // For partially confirmed or unconfirmed routes, compile from all segments
                                var disconnectedPaths = FindDisconnectedPaths(cableContainmentElements, adjacencyGraph);
                                var allPathIds = disconnectedPaths.SelectMany(p => p).ToList();
                                branchSequence = FormatBranchSequence(allPathIds, rtsIdToElementMap);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        status = "Error";
                        routingSequence = $"Processing failed: {ex.Message}";
                        branchSequence = "Error";
                    }

                    sb.AppendLine($"\"{cableRef}\",\"{cableInfo.From ?? "N/A"}\",\"{cableInfo.To ?? "N/A"}\",\"{status}\",\"{branchSequence}\",\"{routingSequence}\"");
                }

                if (!progressWindow.IsCancelled)
                {
                    System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                    progressWindow.UpdateProgress(totalCables, totalCables, "Complete", "N/A", "N/A", 100, "Exporting file...");
                    progressWindow.TaskDescriptionText.Text = "Export complete.";
                    Autodesk.Revit.UI.TaskDialog.Show("Export Complete", $"Routing Sequence report successfully exported to:\n{filePath}");
                }
            }
            catch (Exception ex)
            {
                progressWindow.ShowError($"An unexpected error occurred: {ex.Message}");
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"Failed to export Routing Sequence report: {ex.Message}");
            }
            finally
            {
                progressWindow.Close();
            }
        }

        #region Routing Sequence Helper Methods (Using ElementId)

        /// <summary>
        /// Builds an adjacency graph using ElementId as keys for stability.
        /// </summary>
        private Dictionary<ElementId, List<ElementId>> BuildAdjacencyGraph(List<Element> elements)
        {
            var graph = new Dictionary<ElementId, List<ElementId>>();
            foreach (var elemA in elements)
            {
                if (!graph.ContainsKey(elemA.Id))
                {
                    graph[elemA.Id] = new List<ElementId>();
                }
                foreach (var elemB in elements)
                {
                    if (elemA.Id == elemB.Id) continue;
                    if (AreConnected(elemA, elemB))
                    {
                        graph[elemA.Id].Add(elemB.Id);
                    }
                }
            }
            return graph;
        }

        /// <summary>
        /// Finds a single path between two nodes using Breadth-First Search on ElementIds.
        /// </summary>
        private List<ElementId> FindSinglePath_BFS(ElementId startNodeId, ElementId endNodeId, Dictionary<ElementId, List<ElementId>> graph)
        {
            if (startNodeId == null || endNodeId == null || startNodeId == endNodeId)
            {
                return new List<ElementId> { startNodeId ?? endNodeId };
            }

            var queue = new Queue<ElementId>();
            var cameFrom = new Dictionary<ElementId, ElementId>();
            var visited = new HashSet<ElementId>();

            queue.Enqueue(startNodeId);
            visited.Add(startNodeId);

            while (queue.Count > 0)
            {
                var currentId = queue.Dequeue();

                if (currentId == endNodeId)
                {
                    var path = new List<ElementId>();
                    var atId = currentId;
                    while (atId != null && atId != ElementId.InvalidElementId)
                    {
                        path.Add(atId);
                        if (!cameFrom.TryGetValue(atId, out ElementId prevId)) break;
                        atId = prevId;
                    }
                    path.Reverse();
                    return path;
                }

                if (graph.TryGetValue(currentId, out List<ElementId> neighborIds))
                {
                    foreach (var neighborId in neighborIds)
                    {
                        if (!visited.Contains(neighborId))
                        {
                            visited.Add(neighborId);
                            cameFrom[neighborId] = currentId;
                            queue.Enqueue(neighborId);
                        }
                    }
                }
            }
            return null; // Path not found
        }

        /// <summary>
        /// Finds the path between two confirmed equipment endpoints.
        /// </summary>
        private List<ElementId> FindConfirmedPath(Element matchedStart, Element matchedEnd, List<Element> cableElements, Dictionary<ElementId, List<ElementId>> graph)
        {
            XYZ matchedStartPt = GetElementLocation(matchedStart);
            XYZ matchedEndPt = GetElementLocation(matchedEnd);
            if (matchedStartPt == null || matchedEndPt == null) return new List<ElementId>();

            Element startElem = cableElements
                .OrderBy(e => GetElementLocation(e)?.DistanceTo(matchedStartPt) ?? double.MaxValue)
                .FirstOrDefault();

            Element endElem = cableElements
                .OrderBy(e => GetElementLocation(e)?.DistanceTo(matchedEndPt) ?? double.MaxValue)
                .FirstOrDefault();

            if (startElem == null || endElem == null) return new List<ElementId>();

            return FindSinglePath_BFS(startElem.Id, endElem.Id, graph) ?? new List<ElementId>();
        }

        /// <summary>
        /// Finds all disconnected path segments in the graph.
        /// </summary>
        private List<List<ElementId>> FindDisconnectedPaths(List<Element> cableElements, Dictionary<ElementId, List<ElementId>> graph)
        {
            var allPaths = new List<List<ElementId>>();
            var visited = new HashSet<ElementId>();
            var elementIds = cableElements.Select(e => e.Id).ToList();

            foreach (var startNodeId in elementIds)
            {
                if (visited.Contains(startNodeId)) continue;

                var pathSegment = new List<ElementId>();
                var q = new Queue<ElementId>();

                q.Enqueue(startNodeId);
                visited.Add(startNodeId);
                pathSegment.Add(startNodeId);

                while (q.Count > 0)
                {
                    var currentId = q.Dequeue();
                    if (graph.TryGetValue(currentId, out List<ElementId> neighborIds))
                    {
                        foreach (var neighborId in neighborIds)
                        {
                            if (!visited.Contains(neighborId))
                            {
                                visited.Add(neighborId);
                                pathSegment.Add(neighborId);
                                q.Enqueue(neighborId);
                            }
                        }
                    }
                }
                allPaths.Add(pathSegment);
            }
            return allPaths;
        }

        /// <summary>
        /// Formats a list of ElementIds into the final report string.
        /// </summary>
        private string FormatPath(List<ElementId> path, Dictionary<ElementId, List<ElementId>> graph, Dictionary<string, Element> rtsIdToElementMap)
        {
            if (path == null || !path.Any()) return "Path not found";

            var elementIdToRtsIdMap = rtsIdToElementMap.ToDictionary(kvp => kvp.Value.Id, kvp => kvp.Key);
            var sequence = new List<string>();

            for (int i = 0; i < path.Count; i++)
            {
                var currentId = path[i];
                if (!elementIdToRtsIdMap.TryGetValue(currentId, out string rtsId)) continue;

                if (i > 0)
                {
                    var prevId = path[i - 1];
                    var current = _doc.GetElement(currentId);
                    var prev = _doc.GetElement(prevId);

                    if (current == null || prev == null) continue;

                    if (current.Category.Id != prev.Category.Id)
                    {
                        sequence.Add("+");
                    }
                    else if (graph.TryGetValue(prevId, out var neighbors) && !neighbors.Contains(currentId))
                    {
                        sequence.Add("||");
                    }
                    else
                    {
                        sequence.Add(",");
                    }
                }
                sequence.Add(rtsId);
            }
            return string.Join(" ", sequence).Replace(" , ", ", ").Replace(" + ", " + ").Replace(" || ", " || ");
        }

        /// <summary>
        /// Formats a list of disconnected paths into the final report string.
        /// </summary>
        private string FormatDisconnectedPaths(List<List<ElementId>> allPaths, Dictionary<ElementId, List<ElementId>> graph, Dictionary<string, Element> rtsIdMap)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < allPaths.Count; i++)
            {
                if (i > 0) sb.Append(" >> ");
                sb.Append(FormatPath(allPaths[i], graph, rtsIdMap));
            }
            return sb.ToString();
        }

        private bool AreConnected(Element a, Element b)
        {
            if (a == null || b == null) return false;
            try
            {
                var connectorsA = (a as MEPCurve)?.ConnectorManager?.Connectors;
                var connectorsB = (b as MEPCurve)?.ConnectorManager?.Connectors;

                if (connectorsA == null || connectorsB == null) return false;

                foreach (Connector cA in connectorsA)
                {
                    foreach (Connector cB in connectorsB)
                    {
                        if (cA.IsConnectedTo(cB) || cA.Origin.IsAlmostEqualTo(cB.Origin, 0.1))
                        {
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Finds all matching equipment elements, prioritizing RTS_ID over Panel Name.
        /// </summary>
        /// <returns>A list of all potential matching elements.</returns>
        private List<Element> FindMatchingEquipment(string idToMatch, List<Element> allCandidateEquipment, Guid rtsIdGuid)
        {
            if (string.IsNullOrWhiteSpace(idToMatch)) return new List<Element>();

            // First, try to find matches based on RTS_ID
            var rtsIdMatches = allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var rtsIdParam = e.get_Parameter(rtsIdGuid);
                var rtsIdVal = rtsIdParam?.AsString();
                return !string.IsNullOrWhiteSpace(rtsIdVal) && idToMatch.IndexOf(rtsIdVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            // If any RTS_ID matches are found, return them and ignore Panel Name matches
            if (rtsIdMatches.Any())
            {
                return rtsIdMatches;
            }

            // If no RTS_ID matches, fall back to finding matches by Panel Name
            var panelNameMatches = allCandidateEquipment.Where(e =>
            {
                if (e == null) return false;
                var panelNameParam = e.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME);
                var panelNameVal = panelNameParam?.AsString();
                return !string.IsNullOrWhiteSpace(panelNameVal) && idToMatch.IndexOf(panelNameVal, StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            return panelNameMatches;
        }

        private XYZ GetElementLocation(Element element)
        {
            if (element == null) return null;
            try
            {
                if (element.Location is LocationPoint locPoint) return locPoint.Point;
                if (element.Location is LocationCurve locCurve) return locCurve.Curve.GetEndPoint(0);
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Calculates the total physical length of a path of containment elements.
        /// </summary>
        private double GetPathLength(List<ElementId> pathIds)
        {
            if (pathIds == null || !pathIds.Any()) return 0.0;

            double totalLength = 0.0;
            foreach (var id in pathIds)
            {
                Element elem = _doc.GetElement(id);
                if (elem is MEPCurve curve)
                {
                    totalLength += curve.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                }
            }
            return totalLength;
        }

        /// <summary>
        /// Extracts the branch number from an RTS_ID string.
        /// </summary>
        private string GetBranchNumber(string rtsId)
        {
            if (string.IsNullOrWhiteSpace(rtsId) || rtsId.Length < 4) return null;

            // The branch number is the last 4 characters of the ID
            return rtsId.Substring(rtsId.Length - 4);
        }

        /// <summary>
        /// Formats the branch sequence string from a given path.
        /// </summary>
        private string FormatBranchSequence(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap)
        {
            if (path == null || !path.Any()) return "N/A";

            var elementIdToRtsIdMap = rtsIdToElementMap.ToDictionary(kvp => kvp.Value.Id, kvp => kvp.Key);
            var uniqueBranches = new List<string>();
            var seenBranches = new HashSet<string>();

            foreach (var elementId in path)
            {
                if (elementIdToRtsIdMap.TryGetValue(elementId, out string rtsId))
                {
                    string branchNumber = GetBranchNumber(rtsId);
                    if (!string.IsNullOrEmpty(branchNumber) && seenBranches.Add(branchNumber))
                    {
                        uniqueBranches.Add(branchNumber);
                    }
                }
            }
            return string.Join(", ", uniqueBranches);
        }

        private List<XYZ> GetCableEndpoints(List<Element> cableElements)
        {
            var endpoints = new List<XYZ>();
            foreach (var elem in cableElements)
            {
                if (elem?.Location is LocationCurve locCurve && locCurve.Curve != null)
                {
                    try
                    {
                        endpoints.Add(locCurve.Curve.GetEndPoint(0));
                        endpoints.Add(locCurve.Curve.GetEndPoint(1));
                    }
                    catch { }
                }
            }
            return endpoints;
        }
        #endregion

        #endregion

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

            string filePath = GetOutputFilePath(defaultFileName, $"Save {dataTypeName} Report");
            if (string.IsNullOrEmpty(filePath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Cancelled", "Output file not selected.");
                return;
            }

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

        private string GetOutputFilePath(string defaultFileName, string dialogTitle)
        {
            var dialog = new SaveFileDialog
            {
                Title = dialogTitle,
                FileName = defaultFileName,
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                OverwritePrompt = true
            };

            var helper = new WindowInteropHelper(this);
            var result = dialog.ShowDialog(new Win32Window(helper.Handle));
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                return dialog.FileName;
            }
            return null;
        }

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
                    "Number of Earth Cables", "Earth Size (mm2)", "Voltage", "VoltageRating", "Type", "Sheath Construction",
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

        private string CleanCableReference(string cableReference)
        {
            if (string.IsNullOrEmpty(cableReference)) return cableReference;
            string cleaned = cableReference.Trim();
            int openParenIndex = cleaned.IndexOf('(');
            if (openParenIndex != -1)
            {
                cleaned = cleaned.Substring(0, openParenIndex).Trim();
            }
            int firstSlashIndex = cleaned.IndexOf('/');
            if (firstSlashIndex != -1)
            {
                cleaned = cleaned.Substring(0, firstSlashIndex).Trim();
            }
            string[] parts = cleaned.Split('-');
            Regex prefixPattern = new Regex(@"^[A-Za-z]{2}\d{2}", RegexOptions.IgnoreCase);
            if (parts.Length >= 3 && prefixPattern.IsMatch(parts[0]))
            {
                if (!int.TryParse(parts[2], out _)) return $"{parts[0]}-{parts[1]}";
            }
            if (parts.Length >= 4 && prefixPattern.IsMatch(parts[0]))
            {
                if (int.TryParse(parts[2], out _) && !int.TryParse(parts[3], out _))
                    return $"{parts[0]}-{parts[1]}-{parts[2]}";
            }
            return cleaned;
        }
    }

    // --- LinkViewModel for grid row ---
    public class ReportLinkViewModel
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsLoaded { get; set; }
        public bool IsPlaceholder { get; set; }
        // Add other properties as needed
    }
}
