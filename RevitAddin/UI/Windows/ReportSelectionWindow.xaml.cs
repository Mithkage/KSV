//-----------------------------------------------------------------------------
// <copyright file="ReportSelectionWindow.xaml.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     This file contains the interaction logic for the ReportSelectionWindow.xaml.
//     It handles user selections for various report types and orchestrates
//     the generation of different report formats.
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
 * 2025-08-07 | 5.0.0   | Gemini | Major refactoring: Split report generators into separate files.
 * |         |         |        | - Created modular report generation system with base classes and specific generators.
 * |         |         |        | - Improved maintainability with single-responsibility classes.
 */

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Interop;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;
using RTS.Reports.Generators;
using RTS.Reports.Models;
using RTS.Reports.Utils;
using RTS.Commands.DataExchange.DataManagement;
#endregion

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for ReportSelectionWindow.xaml
    /// This class handles user selections for report generation.
    /// </summary>
    public partial class ReportSelectionWindow : Window
    {
        private Document _doc;
        private PC_ExtensibleClass _pcExtensible;
        private ExternalCommandData _commandData; // Stored to pass to report generators

        // Collection of LinkViewModel for the grid
        public List<ReportLinkViewModel> Links { get; set; }

        public ReportSelectionWindow(ExternalCommandData commandData)
        {
            InitializeComponent();
            _doc = commandData.Application.ActiveUIDocument.Document;
            _pcExtensible = new PC_ExtensibleClass();
            _commandData = commandData;
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

        // --- Report Generation Button Handlers ---

        private void ExportMyPowerCADDataButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new PowerCADDataReportGenerator(
                _doc,
                _commandData,
                _pcExtensible,
                PC_ExtensibleClass.PrimarySchemaGuid,
                PC_ExtensibleClass.PrimarySchemaName,
                PC_ExtensibleClass.PrimaryFieldName,
                PC_ExtensibleClass.PrimaryDataStorageElementName,
                "My_PowerCAD_Data_Report.csv",
                "My PowerCAD Data"
            );
            
            generator.GenerateReport();
        }

        private void ExportConsultantPowerCADDataButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new PowerCADDataReportGenerator(
                _doc,
                _commandData,
                _pcExtensible,
                PC_ExtensibleClass.ConsultantSchemaGuid,
                PC_ExtensibleClass.ConsultantSchemaName,
                PC_ExtensibleClass.ConsultantFieldName,
                PC_ExtensibleClass.ConsultantDataStorageElementName,
                "Consultant_PowerCAD_Data_Report.csv",
                "Consultant PowerCAD Data"
            );
            
            generator.GenerateReport();
        }

        private void ExportModelGeneratedDataButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new PowerCADDataReportGenerator(
                _doc,
                _commandData,
                _pcExtensible,
                PC_ExtensibleClass.ModelGeneratedSchemaGuid,
                PC_ExtensibleClass.ModelGeneratedSchemaName,
                PC_ExtensibleClass.ModelGeneratedFieldName,
                PC_ExtensibleClass.ModelGeneratedDataStorageElementName,
                "Model_Generated_Data_Report.csv",
                "Model Generated Data"
            );
            
            generator.GenerateReport();
        }

        private void GenerateRSGxCableScheduleCsvReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new RSGxCableScheduleGenerator(
                _doc,
                _commandData,
                _pcExtensible
            );
            
            generator.GenerateReport();
        }

        private void GenerateRSGxCableSummaryXlsxReport_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new RSGxCableSummaryGenerator(
                _doc,
                _commandData,
                _pcExtensible
            );
            
            generator.GenerateReport();
        }

        private void GenerateRoutingSequenceReportButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            
            var generator = new RoutingSequenceReportGenerator(
                _doc,
                _commandData,
                _pcExtensible
            );
            
            generator.GenerateReport();
        }
    }
}
