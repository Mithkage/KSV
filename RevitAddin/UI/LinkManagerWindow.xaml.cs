// --- FILE: LinkManagerWindow.xaml.cs (UPDATED) ---
//
// File: LinkManagerWindow.xaml.cs
//
// Namespace: RTS.UI
//
// Class: LinkManagerWindow, LinkViewModel, BulkEditDialog
//
// Function: This file contains the code-behind logic for the Link Manager WPF window.
//           It is responsible for loading, displaying, and saving Revit link metadata
//           to a "Project Profile" within the project's Extensible Storage. It now
//           supports adding and deleting placeholder rows for links, and includes
//           various UI enhancements.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 23, 2025: Refined Revizto matching logic to ignore placeholders.
// - July 23, 2025: Corrected ContextMenu logic to reliably attach to the DataGridRow, fixing the right-click feature.
// - July 23, 2025: Implemented "Reload From..." context menu to relink existing links or resolve placeholders.
// - July 23, 2025: Added Import Profile CSV functionality and updated Save button behavior.
// - July 23, 2025: Removed obsolete Excel export method.
// - July 23, 2025: Updated export functionality to generate a CSV file instead of an XLSX file.
// - July 23, 2025: Implemented advanced data merging in LoadLinks to handle re-imports and placeholder resolution.
// - July 23, 2025: Corrected a typo in the Revizto extension matching logic to ensure correct status reporting.
// - July 23, 2025: Updated Revizto matching to ignore file extensions and report alternative matches. Stopped populating Link Description from Revizto data.
// - July 23, 2025: Corrected Save Profile logic to save all links in the grid, preserving metadata for all placeholder types.
// - July 23, 2025: Modified delete and clear profile logic to permanently remove associated Revizto records from Extensible Storage, preventing them from reappearing.
// - July 23, 2025: Added functionality to import Revizto link data from a CSV file.
// - July 23, 2025: Corrected icon loading to use an explicit assembly reference in the Pack URI, resolving Revit host warnings.
// - July 22, 2025: Refined Revizto import to handle duplicate entries by prioritizing the latest 'Last Modified' date.
// - July 22, 2025: Added specific exception handling for TypeInitializationException on Excel import/export.
// - July 22, 2025: Enhanced error reporting to guide users on resolving ClosedXML dependency conflicts.
// - July 19, 2025: Added specific exception handling for TypeLoadException on Excel export to diagnose version conflicts.
// - July 19, 2025: Implemented dynamic ContextMenu creation in code-behind to resolve runtime errors.
// - July 19, 2025: Added DispatcherUnhandledException handler for improved XAML runtime error diagnostics.
// - July 19, 2025: Added robust icon loading in constructor to prevent runtime errors from missing image resources.
// - July 19, 2025: Corrected Revit 2022 API calls for IsRelativePath and GetLoadStatus to resolve compiler errors.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO; // Required for TextFieldParser
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics; // Required for Process.Start
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection; // Required for Assembly.GetExecutingAssembly
using System.Text; // Required for StringBuilder
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media; // Required for VisualTreeHelper
using System.Windows.Media.Imaging; // Required for BitmapImage
using System.Windows.Threading; // Required for DispatcherUnhandledException
#endregion

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for LinkManagerWindow.xaml
    /// </summary>
    public partial class LinkManagerWindow : Window
    {
        private Document _doc;
        private ICollectionView _linksView;
        private Dictionary<string, string> _activeColumnFilters = new Dictionary<string, string>();

        // --- Extensible Storage Definitions for Link Manager Profile ---
        public static readonly Guid ProfileSchemaGuid = new Guid("D4C6E8B0-6A2E-4B1C-9D7E-8C4F2A6B9E1F");
        public const string ProfileSchemaName = "RTS_LinkManagerProfileSchema";
        public const string ProfileFieldName = "LinkManagerProfileJson";
        public const string ProfileDataStorageElementName = "RTS_LinkManager_Profile_Storage";
        public const string VendorId = "ReTick_Solutions"; // Must match your .addin file

        public ObservableCollection<LinkViewModel> Links { get; set; }

        public LinkManagerWindow(Document doc)
        {
            // Add an exception handler for more detailed XAML parsing errors
            this.Dispatcher.UnhandledException += OnDispatcherUnhandledException;

            InitializeComponent();
            LoadIcon(); // Load icon robustly
            _doc = doc;
            this.DataContext = this;

            Links = new ObservableCollection<LinkViewModel>();
            _linksView = CollectionViewSource.GetDefaultView(Links);
            _linksView.Filter = new Predicate<object>(FilterLinks);

            LoadLinks();
        }

        /// <summary>
        /// Provides detailed exception information for XAML parsing and other UI thread errors.
        /// </summary>
        void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            // Prevent the application from crashing
            e.Handled = true;

            string errorMessage = $"An unexpected error occurred: {e.Exception.Message}";

            // The InnerException often contains the most useful information for XAML errors
            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\nInner Exception: {e.Exception.InnerException.Message}";
            }

            // Show a detailed error message to the user
            TaskDialog.Show("Runtime Error", errorMessage);
        }


        /// <summary>
        /// Loads the window icon safely, using a robust Pack URI that specifies the assembly.
        /// This method prevents errors when the window is hosted in an external application like Revit.
        /// </summary>
        private void LoadIcon()
        {
            try
            {
                // Get the current assembly name programmatically to create a robust Pack URI.
                // This resolves issues when the add-in is loaded into a host application like Revit.
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.png", UriKind.Absolute);
                this.Icon = new BitmapImage(iconUri);
            }
            catch (Exception ex)
            {
                // If the icon resource is missing or there's an error loading it,
                // this will prevent the application from crashing.
                // The original warning from Revit is now handled here with more specific debug info.
                Debug.WriteLine($"Warning: Could not load window icon. {ex.Message}");
            }
        }

        /// <summary>
        /// Filters the links displayed in the DataGrid based on the active column filters.
        /// </summary>
        private bool FilterLinks(object item)
        {
            if (item is LinkViewModel link)
            {
                foreach (var filter in _activeColumnFilters)
                {
                    string propertyName = filter.Key;
                    string filterValue = filter.Value;

                    if (string.IsNullOrEmpty(filterValue)) continue;

                    var prop = typeof(LinkViewModel).GetProperty(propertyName);
                    if (prop != null)
                    {
                        string cellValue = prop.GetValue(link)?.ToString() ?? string.Empty;
                        if (!cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase))
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Handles the click event on a DataGridColumnHeader to display a filter context menu.
        /// </summary>
        private void ColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var header = FindParent<DataGridColumnHeader>(button);
            if (header == null) return;

            string sortMemberPath = header.Column.SortMemberPath;
            if (string.IsNullOrEmpty(sortMemberPath)) return;

            ContextMenu contextMenu = new ContextMenu();

            MenuItem clearFilterMenuItem = new MenuItem { Header = "Clear Filter" };
            clearFilterMenuItem.Click += (s, args) =>
            {
                _activeColumnFilters.Remove(sortMemberPath);
                _linksView.Refresh();
                // Fully qualify Control to resolve ambiguity
                header.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
            };
            contextMenu.Items.Add(clearFilterMenuItem);
            contextMenu.Items.Add(new Separator());

            var uniqueValues = Links.Select(l => typeof(LinkViewModel).GetProperty(sortMemberPath).GetValue(l)?.ToString())
                                        .Where(v => !string.IsNullOrEmpty(v))
                                        .Distinct(StringComparer.OrdinalIgnoreCase)
                                        .OrderBy(v => v)
                                        .ToList();

            foreach (var value in uniqueValues)
            {
                MenuItem menuItem = new MenuItem { Header = value };
                menuItem.IsChecked = _activeColumnFilters.ContainsKey(sortMemberPath) && _activeColumnFilters[sortMemberPath].Equals(value, StringComparison.OrdinalIgnoreCase);
                menuItem.Click += (s, args) =>
                {
                    _activeColumnFilters[sortMemberPath] = value;
                    _linksView.Refresh();
                    header.Background = System.Windows.Media.Brushes.LightBlue;
                };
                contextMenu.Items.Add(menuItem);
            }

            contextMenu.PlacementTarget = button;
            contextMenu.IsOpen = true;
        }

        /// <summary>
        /// Scans all data sources (Revit, saved profile, Revizto data) and intelligently merges them to build the link list.
        /// </summary>
        private void LoadLinks()
        {
            // 1. Load all data sources
            var savedProfiles = RecallDataFromExtensibleStorage<LinkViewModel>(_doc, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
            var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
            var allLinkTypes = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>().Where(lt => !lt.IsNestedLink).ToList();
            var linkInstancesGroupedByType = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().GroupBy(inst => inst.GetTypeId()).ToDictionary(g => g.Key, g => g.ToList());

            var finalLinkList = new List<LinkViewModel>();
            var processedPlaceholders = new HashSet<string>();

            // 2. Process all live Revit links
            foreach (var type in allLinkTypes)
            {
                var viewModel = new LinkViewModel { LinkName = type.Name, IsRevitLink = true };

                // Populate live Revit data
                PopulateRevitLinkData(viewModel, type, linkInstancesGroupedByType);

                // Check for a matching placeholder in the saved profiles to inherit its data
                string linkNameWithoutExt = Path.GetFileNameWithoutExtension(type.Name);
                var matchingPlaceholder = savedProfiles.FirstOrDefault(p => !p.IsRevitLink && Path.GetFileNameWithoutExtension(p.LinkName).Equals(linkNameWithoutExt, StringComparison.OrdinalIgnoreCase));

                if (matchingPlaceholder != null)
                {
                    // A live link now exists for a previous placeholder. Merge the data.
                    viewModel.LinkDescription = matchingPlaceholder.LinkDescription;
                    viewModel.SelectedDiscipline = matchingPlaceholder.SelectedDiscipline;
                    viewModel.CompanyName = matchingPlaceholder.CompanyName;
                    viewModel.ResponsiblePerson = matchingPlaceholder.ResponsiblePerson;
                    viewModel.ContactDetails = matchingPlaceholder.ContactDetails;
                    viewModel.Comments = matchingPlaceholder.Comments;
                    processedPlaceholders.Add(matchingPlaceholder.LinkName);
                }
                else
                {
                    // No placeholder match, check for a direct match in the saved profiles
                    var savedProfile = savedProfiles.FirstOrDefault(p => p.LinkName.Equals(type.Name, StringComparison.OrdinalIgnoreCase));
                    if (savedProfile != null)
                    {
                        viewModel.LinkDescription = savedProfile.LinkDescription;
                        viewModel.SelectedDiscipline = savedProfile.SelectedDiscipline;
                        viewModel.CompanyName = savedProfile.CompanyName;
                        viewModel.ResponsiblePerson = savedProfile.ResponsiblePerson;
                        viewModel.ContactDetails = savedProfile.ContactDetails;
                        viewModel.Comments = savedProfile.Comments;
                    }
                    else
                    {
                        // No saved data at all, set defaults
                        viewModel.SelectedDiscipline = "Unconfirmed";
                        viewModel.CompanyName = "Not Set";
                    }
                }
                finalLinkList.Add(viewModel);
            }

            // 3. Process remaining saved profiles (placeholders that were not matched to a live link)
            foreach (var savedProfile in savedProfiles)
            {
                if (processedPlaceholders.Contains(savedProfile.LinkName) || finalLinkList.Any(l => l.LinkName.Equals(savedProfile.LinkName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // Skip placeholders that were already merged or links that are already processed
                }
                SetPlaceholderProperties(savedProfile);
                finalLinkList.Add(savedProfile);
            }

            // 4. Update Revizto status for all links in the list and add any Revizto-only links
            foreach (var linkVm in finalLinkList)
            {
                // Only attempt to match if it's a real Revit link
                if (linkVm.IsRevitLink)
                {
                    string linkNameWithoutExt = Path.GetFileNameWithoutExtension(linkVm.LinkName);
                    var reviztoMatch = reviztoRecords.FirstOrDefault(r => Path.GetFileNameWithoutExtension(r.LinkName).Equals(linkNameWithoutExt, StringComparison.OrdinalIgnoreCase));

                    if (reviztoMatch != null)
                    {
                        // A match was found in the current Revizto import
                        linkVm.LastModified = reviztoMatch.LastModified;
                        string reviztoExt = Path.GetExtension(reviztoMatch.LinkName);
                        string linkExt = Path.GetExtension(linkVm.LinkName);
                        linkVm.ReviztoStatus = linkExt.Equals(reviztoExt, StringComparison.OrdinalIgnoreCase) ? "Matched" : $"Match on: {reviztoExt}";
                    }
                    else
                    {
                        // No match was found in the current Revizto import
                        if (linkVm.ReviztoStatus == "Matched" || (linkVm.ReviztoStatus != null && linkVm.ReviztoStatus.StartsWith("Match on:")))
                        {
                            // This item used to match a Revizto link, but doesn't anymore
                            linkVm.ReviztoStatus = "Removed";
                        }
                        else
                        {
                            linkVm.ReviztoStatus = "Not Matched";
                        }
                    }
                }
            }

            // Add Revizto records that don't match any existing link or placeholder
            foreach (var reviztoRecord in reviztoRecords)
            {
                string reviztoNameWithoutExt = Path.GetFileNameWithoutExtension(reviztoRecord.LinkName);
                if (!finalLinkList.Any(l => Path.GetFileNameWithoutExtension(l.LinkName).Equals(reviztoNameWithoutExt, StringComparison.OrdinalIgnoreCase)))
                {
                    var viewModel = new LinkViewModel
                    {
                        LinkName = reviztoRecord.LinkName,
                        LastModified = reviztoRecord.LastModified,
                        ReviztoStatus = "Missing"
                    };
                    SetPlaceholderProperties(viewModel);
                    finalLinkList.Add(viewModel);
                }
            }

            // 5. Populate the UI
            Links.Clear();
            foreach (var link in finalLinkList)
            {
                Links.Add(link);
            }
            _linksView.Refresh();
        }

        /// <summary>
        /// Populates a LinkViewModel with data from a live RevitLinkType element.
        /// </summary>
        private void PopulateRevitLinkData(LinkViewModel viewModel, RevitLinkType type, Dictionary<ElementId, List<RevitLinkInstance>> linkInstancesGroupedByType)
        {
            string lastModifiedDate = "Not Found";
            string linkPath = string.Empty;
            string fullPath = string.Empty;
            try
            {
                ModelPath modelPath = type.GetExternalFileReference()?.GetPath();
                if (modelPath != null)
                {
                    linkPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    if (!string.IsNullOrEmpty(linkPath) && File.Exists(linkPath))
                    {
                        fullPath = Path.GetFullPath(linkPath);
                        lastModifiedDate = File.GetLastWriteTime(fullPath).ToString("yyyy-MM-dd HH:mm");
                        viewModel.PathType = Path.IsPathRooted(linkPath) ? "Absolute" : "Relative";
                        viewModel.HasValidPath = true;
                    }
                    else
                    {
                        viewModel.PathType = "N/A (File Not Found)";
                    }
                }
            }
            catch { /* Ignore */ }
            viewModel.LastModified = lastModifiedDate;
            viewModel.FullPath = fullPath;

            viewModel.LinkStatus = RevitLinkType.IsLoaded(_doc, type.Id) ? "Loaded" : "Unloaded";

            if (viewModel.LinkStatus == "Loaded" && viewModel.HasValidPath)
            {
                DateTime fileLastWriteTime = File.GetLastWriteTime(fullPath);
                if (DateTime.TryParse(viewModel.LastModified, out DateTime revitLinkLastWriteTime))
                {
                    viewModel.VersionStatus = fileLastWriteTime > revitLinkLastWriteTime ? "Newer local file" : "Up-to-date";
                }
                else
                {
                    viewModel.VersionStatus = "N/A (Date Parse Error)";
                }
            }
            else if (viewModel.LinkStatus == "Unloaded")
            {
                viewModel.VersionStatus = "N/A (Unloaded)";
            }
            else
            {
                viewModel.VersionStatus = "N/A (File not found)";
            }

            if (linkInstancesGroupedByType.TryGetValue(type.Id, out var instances) && instances.Any())
            {
                viewModel.LinkInstanceIds = instances.Select(i => i.Id).ToList();
                viewModel.LinkTypeId = type.Id;
                viewModel.NumberOfInstances = instances.Count;

                if (_doc.IsWorkshared)
                {
                    var worksetNames = instances.Select(i => _doc.GetWorksetTable().GetWorkset(i.WorksetId)?.Name).Where(n => n != null).Distinct().ToList();
                    viewModel.LinkWorkset = worksetNames.Count > 1 ? "Multiple Worksets" : worksetNames.FirstOrDefault() ?? "Unknown";
                }
                else
                {
                    viewModel.LinkWorkset = "N/A";
                }

                Parameter sharedSiteParam = instances.First().LookupParameter("Shared Site");
                if (sharedSiteParam != null && sharedSiteParam.HasValue)
                {
                    var siteLocation = _doc.GetElement(sharedSiteParam.AsElementId()) as ProjectLocation;
                    viewModel.SelectedCoordinates = siteLocation != null ? siteLocation.Name : "Origin to Origin";
                }
                else
                {
                    viewModel.SelectedCoordinates = "Unconfirmed";
                }
            }
            else
            {
                viewModel.NumberOfInstances = 0;
                viewModel.LinkWorkset = "N/A";
                viewModel.SelectedCoordinates = "N/A";
            }
        }

        private void SetPlaceholderProperties(LinkViewModel vm)
        {
            vm.IsRevitLink = false;
            vm.HasValidPath = false;
            vm.LinkInstanceIds = new List<ElementId>();
            vm.LinkWorkset = "N/A";
            vm.NumberOfInstances = 0;
            vm.SelectedCoordinates = "N/A";
            vm.PathType = "N/A";
            vm.LinkStatus = "N/A";
            vm.VersionStatus = "N/A";
            vm.SelectedDiscipline ??= "Unconfirmed";
            vm.CompanyName ??= "Not Set";
        }

        /// <summary>
        /// Handles the click event for the "Save" button. Saves the profile without closing the window.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            using (Transaction tx = new Transaction(_doc, "Save Link Manager Profile"))
            {
                tx.Start();
                try
                {
                    // Save all links currently in the grid. The LoadLinks method is responsible
                    // for correctly merging this saved data with live Revit and Revizto data on next open.
                    var linksToSave = Links.ToList();
                    SaveDataToExtensibleStorage(_doc, linksToSave, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);

                    tx.Commit();
                    TaskDialog.Show("Success", "Project link profile saved successfully!", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to save profile: {ex.Message}", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);
                }
            }
        }

        /// <summary>
        /// Handles the click event for the "Add Placeholder" button.
        /// </summary>
        private void AddPlaceholderButton_Click(object sender, RoutedEventArgs e)
        {
            var placeholder = new LinkViewModel { LinkName = "Placeholder", ReviztoStatus = "Placeholder" };
            SetPlaceholderProperties(placeholder);
            placeholder.LinkDescription = "Awaiting Link Model";
            Links.Add(placeholder);
            _linksView.Refresh();
        }

        /// <summary>
        /// Handles deleting placeholders with the Delete key. This now also removes the corresponding
        /// record from the Revizto extensible storage to prevent it from reappearing.
        /// </summary>
        private void LinksDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                var grid = sender as DataGrid;
                if (grid?.SelectedItems.Count > 0)
                {
                    var placeholdersToDelete = grid.SelectedItems.Cast<LinkViewModel>().Where(vm => !vm.IsRevitLink).ToList();
                    if (placeholdersToDelete.Any())
                    {
                        var result = TaskDialog.Show("Confirm Delete", $"Are you sure you want to permanently delete {placeholdersToDelete.Count} placeholder(s)? This will also remove them from the saved Revizto import data.", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                        if (result == TaskDialogResult.Yes)
                        {
                            using (var tx = new Transaction(_doc, "Delete Placeholders and Revizto Records"))
                            {
                                tx.Start();
                                try
                                {
                                    // Get the current Revizto records from storage
                                    var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                                    bool reviztoDataModified = false;

                                    var placeholdersToRemoveFromUi = new List<LinkViewModel>();

                                    foreach (var placeholder in placeholdersToDelete)
                                    {
                                        // Find and remove the matching Revizto record if it exists
                                        var recordToRemove = reviztoRecords.FirstOrDefault(r => r.LinkName.Equals(placeholder.LinkName, StringComparison.OrdinalIgnoreCase));
                                        if (recordToRemove != null)
                                        {
                                            reviztoRecords.Remove(recordToRemove);
                                            reviztoDataModified = true;
                                        }
                                        placeholdersToRemoveFromUi.Add(placeholder);
                                    }

                                    // If we removed any Revizto records, save the updated list back to storage
                                    if (reviztoDataModified)
                                    {
                                        SaveDataToExtensibleStorage(_doc, reviztoRecords, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                                    }

                                    tx.Commit();

                                    // Remove from the UI collection after the transaction is successful
                                    foreach (var p in placeholdersToRemoveFromUi)
                                    {
                                        Links.Remove(p);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    tx.RollBack();
                                    TaskDialog.Show("Error", $"Failed to delete placeholders: {ex.Message}");
                                }
                            }
                        }
                    }
                    else
                    {
                        TaskDialog.Show("Deletion Not Allowed", "Only placeholder rows can be deleted. To remove a Revit link, use the right-click context menu.", TaskDialogCommonButtons.Ok);
                    }
                    e.Handled = true;
                }
            }
        }

        #region Context Menu Handlers

        private void LinksDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            try
            {
                // The source of the event is often a low-level element like a TextBlock or Border.
                // We need to find the DataGridRow that contains this element.
                var dependencyObject = e.OriginalSource as DependencyObject;
                var row = FindParent<DataGridRow>(dependencyObject);

                if (row == null) return;

                var vm = row.DataContext as LinkViewModel;
                if (vm == null) return;

                var menu = new ContextMenu();
                var reloadItem = new MenuItem { Header = "Reload From...", Tag = vm };
                reloadItem.Click += ReloadFrom_Click;
                menu.Items.Add(reloadItem);

                // Assign the context menu to the row itself.
                row.ContextMenu = menu;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating context menu: {ex.Message}");
                e.Handled = true;
            }
        }

        /// <summary>
        /// Handles the "Reload From..." context menu command for both live links and placeholders.
        /// </summary>
        private void ReloadFrom_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var vm = menuItem?.Tag as LinkViewModel;
            if (vm == null) return;

            var openFileDialog = new OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt|All Files (*.*)|*.*",
                Title = "Select a Revit file to link"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string newPath = openFileDialog.FileName;
                using (var tx = new Transaction(_doc, "Reload Link"))
                {
                    tx.Start();
                    try
                    {
                        if (vm.IsRevitLink)
                        {
                            // Case 1: Reload an existing Revit link
                            var linkType = _doc.GetElement(vm.LinkTypeId) as RevitLinkType;
                            if (linkType != null)
                            {
                                var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath);
                                linkType.LoadFrom(modelPath, new WorksetConfiguration());
                            }
                        }
                        else
                        {
                            // Case 2: Resolve a placeholder by creating a new link
                            var linkOptions = new RevitLinkOptions(false);
                            RevitLinkType.Create(_doc, ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath), linkOptions);
                        }
                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Error", $"Failed to reload link: {ex.Message}");
                        return;
                    }
                }
                // Refresh the entire grid to reflect the changes
                LoadLinks();
            }
        }


        #endregion

        /// <summary>
        /// Handles the click event for the "Bulk Edit" button.
        /// </summary>
        private void BulkEditButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedLinks = LinksDataGrid.SelectedItems.Cast<LinkViewModel>().ToList();
            if (!selectedLinks.Any())
            {
                TaskDialog.Show("Bulk Edit", "Please select at least one link to bulk edit.", TaskDialogCommonButtons.Ok);
                return;
            }

            var bulkEditDialog = new BulkEditDialog();
            if (bulkEditDialog.ShowDialog() == true)
            {
                foreach (var linkVm in selectedLinks)
                {
                    if (bulkEditDialog.ApplyDiscipline) linkVm.SelectedDiscipline = bulkEditDialog.SelectedDiscipline;
                    if (bulkEditDialog.ApplyCompanyName) linkVm.CompanyName = bulkEditDialog.CompanyName;
                    if (bulkEditDialog.ApplyComments) linkVm.Comments = bulkEditDialog.Comments;
                }
                TaskDialog.Show("Bulk Edit", "Selected links updated. Click 'Save Profile' to make changes permanent.", TaskDialogCommonButtons.Ok);
            }
        }

        /// <summary>
        /// Handles the click event for the "Close" button.
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Handles the click event for the "Clear Profile" button.
        /// Prompts the user for confirmation and clears ALL saved profiles (Link Manager and Revizto) from extensible storage.
        /// </summary>
        private void ClearProfileButton_Click(object sender, RoutedEventArgs e)
        {
            var result = Autodesk.Revit.UI.TaskDialog.Show(
                "Clear All Saved Data",
                "Are you sure you want to clear all saved link data? This will remove the Link Manager profile AND the imported Revizto data. This action cannot be undone.",
                Autodesk.Revit.UI.TaskDialogCommonButtons.Yes | Autodesk.Revit.UI.TaskDialogCommonButtons.No,
                Autodesk.Revit.UI.TaskDialogResult.No);

            if (result == Autodesk.Revit.UI.TaskDialogResult.Yes)
            {
                using (var tx = new Autodesk.Revit.DB.Transaction(_doc, "Clear All Link Manager Data"))
                {
                    tx.Start();
                    try
                    {
                        var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
                        var allDataStorage = collector.ToElements();

                        // Find and delete the Link Manager Profile DataStorage element
                        var profileDataStorage = allDataStorage.FirstOrDefault(ds => ds.Name == ProfileDataStorageElementName);
                        if (profileDataStorage != null)
                        {
                            _doc.Delete(profileDataStorage.Id);
                        }

                        // Find and delete the Revizto Link DataStorage element
                        var reviztoDataStorage = allDataStorage.FirstOrDefault(ds => ds.Name == ReviztoLinkRecord.DataStorageName);
                        if (reviztoDataStorage != null)
                        {
                            _doc.Delete(reviztoDataStorage.Id);
                        }

                        tx.Commit();
                        LoadLinks(); // Refresh the grid, which should now be empty of placeholders
                        TaskDialog.Show("All Data Cleared", "The saved link profile and Revizto data have been cleared.", TaskDialogCommonButtons.Ok);
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Error", $"Failed to clear data: {ex.Message}", TaskDialogCommonButtons.Ok);
                    }
                }
            }
        }


        // Helper to find parent of a certain type in the visual tree
        public static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            T parent = parentObject as T;
            if (parent != null)
                return parent;
            else
                return FindParent<T>(parentObject);
        }

        private void LinksDataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            string sortMemberPath = e.Column.SortMemberPath;
            if (string.IsNullOrEmpty(sortMemberPath)) return;
            ListSortDirection direction = (e.Column.SortDirection == null || e.Column.SortDirection == ListSortDirection.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
            foreach (var col in LinksDataGrid.Columns) { col.SortDirection = null; }
            e.Column.SortDirection = direction;
            var sortedLinks = (direction == ListSortDirection.Ascending) ? Links.OrderBy(l => l.GetType().GetProperty(sortMemberPath).GetValue(l, null) as IComparable).ToList() : Links.OrderByDescending(l => l.GetType().GetProperty(sortMemberPath).GetValue(l, null) as IComparable).ToList();
            Links.Clear();
            foreach (var sortedLink in sortedLinks) { Links.Add(sortedLink); }
        }

        /// <summary>
        /// Handles importing Revizto link data from a CSV file.
        /// </summary>
        private void ImportReviztoCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*",
                Title = "Select Revizto Link Export CSV File"
            };

            if (openFileDialog.ShowDialog() != true) return;

            string filePath = openFileDialog.FileName;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                TaskDialog.Show("Error", "Invalid file path or file does not exist.", TaskDialogCommonButtons.Ok);
                return;
            }

            try
            {
                var allReviztoLinks = new List<ReviztoLinkRecord>();

                using (TextFieldParser parser = new TextFieldParser(filePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    // Read header to find column indices
                    if (parser.EndOfData)
                    {
                        TaskDialog.Show("Import Error", "CSV file is empty or invalid.");
                        return;
                    }
                    string[] headers = parser.ReadFields();
                    int linkNameIndex = Array.IndexOf(headers, "Original name");
                    int lastModifiedIndex = Array.IndexOf(headers, "Last exported");
                    int descriptionIndex = Array.IndexOf(headers, "Model");

                    if (linkNameIndex == -1 || lastModifiedIndex == -1 || descriptionIndex == -1)
                    {
                        TaskDialog.Show("Import Error", "CSV file is missing required columns: 'Original name', 'Last exported', and 'Model'.");
                        return;
                    }

                    while (!parser.EndOfData)
                    {
                        try
                        {
                            string[] fields = parser.ReadFields();
                            var record = new ReviztoLinkRecord
                            {
                                LinkName = fields[linkNameIndex],
                                Description = fields[descriptionIndex],
                                LastModified = fields[lastModifiedIndex],
                                FilePath = "" // Not available in CSV export
                            };
                            if (!string.IsNullOrWhiteSpace(record.LinkName))
                            {
                                allReviztoLinks.Add(record);
                            }
                        }
                        catch (MalformedLineException)
                        {
                            // Optionally log or notify user about skipped lines
                        }
                    }
                }

                // Group by link name and take the one with the most recent 'LastModified' date
                var uniqueReviztoLinks = allReviztoLinks
                    .GroupBy(r => r.LinkName)
                    .Select(g => g.OrderByDescending(r => {
                        DateTime.TryParse(r.LastModified, out DateTime dt);
                        return dt;
                    }).First())
                    .ToList();

                using (var tx = new Autodesk.Revit.DB.Transaction(_doc, "Import Revizto Link Data from CSV"))
                {
                    tx.Start();
                    var schemaGuid = ReviztoLinkRecord.SchemaGuid;
                    var dataStorageName = ReviztoLinkRecord.DataStorageName;
                    var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
                    var existing = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageName);
                    if (existing != null) _doc.Delete(existing.Id);

                    SaveDataToExtensibleStorage(_doc, uniqueReviztoLinks, schemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, dataStorageName);
                    tx.Commit();
                }
                TaskDialog.Show("Import Complete", $"{uniqueReviztoLinks.Count} unique Revizto links imported from CSV and saved successfully.", TaskDialogCommonButtons.Ok);
                LoadLinks(); // Refresh the grid
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Import Error", $"An unexpected error occurred while importing the CSV file: {ex.Message}", TaskDialogCommonButtons.Ok);
            }
        }


        /// <summary>
        /// Handles exporting the current grid data to a CSV file.
        /// </summary>
        private void ExportToCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FileName = "RevitLinkManagerExport.csv",
                Title = "Save Link Data to CSV"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    var sb = new StringBuilder();

                    // Create Header Row
                    var headers = LinksDataGrid.Columns
                        .Where(c => c.Visibility == System.Windows.Visibility.Visible)
                        .Select(c => EscapeCsvField(c.Header.ToString().Replace("&#x0a;", " "))); // Replace newline with space for CSV header
                    sb.AppendLine(string.Join(",", headers));

                    // Create Data Rows
                    foreach (LinkViewModel link in LinksDataGrid.Items)
                    {
                        var fields = new List<string>();
                        foreach (DataGridColumn column in LinksDataGrid.Columns)
                        {
                            if (column.Visibility == System.Windows.Visibility.Visible)
                            {
                                string propertyName = (column as DataGridBoundColumn)?.Binding is System.Windows.Data.Binding binding ? binding.Path.Path : column.SortMemberPath;
                                if (!string.IsNullOrEmpty(propertyName))
                                {
                                    var prop = typeof(LinkViewModel).GetProperty(propertyName);
                                    if (prop != null)
                                    {
                                        string value = prop.GetValue(link)?.ToString() ?? "";
                                        fields.Add(EscapeCsvField(value));
                                    }
                                    else
                                    {
                                        fields.Add("");
                                    }
                                }
                                else
                                {
                                    // Handle template columns that might not have a direct binding path
                                    fields.Add("");
                                }
                            }
                        }
                        sb.AppendLine(string.Join(",", fields));
                    }

                    File.WriteAllText(saveFileDialog.FileName, sb.ToString());
                    TaskDialog.Show("Export Complete", "Link data exported to CSV successfully.", TaskDialogCommonButtons.Ok);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Export Error", $"Failed to export data to CSV: {ex.Message}", TaskDialogCommonButtons.Ok);
                }
            }
        }

        /// <summary>
        /// Handles importing user-editable data from a CSV file, updating existing links and creating placeholders for new ones.
        /// </summary>
        private void ImportProfileCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Title = "Select Profile CSV to Import"
            };

            if (openFileDialog.ShowDialog() != true) return;

            try
            {
                using (var parser = new TextFieldParser(openFileDialog.FileName))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    if (parser.EndOfData)
                    {
                        TaskDialog.Show("Import Error", "CSV file is empty.");
                        return;
                    }

                    // Build header map
                    var headers = parser.ReadFields().Select(h => h.Replace(" ", "").Replace("&#x0a;", "")).ToArray(); // Normalize headers
                    var headerMap = new Dictionary<string, int>();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        headerMap[headers[i]] = i;
                    }

                    int updatedCount = 0;
                    int createdCount = 0;

                    while (!parser.EndOfData)
                    {
                        var fields = parser.ReadFields();
                        string linkName = fields[headerMap["LinkFileName"]];

                        var existingLink = Links.FirstOrDefault(l => l.LinkName.Equals(linkName, StringComparison.OrdinalIgnoreCase));

                        if (existingLink != null)
                        {
                            // Update existing link
                            existingLink.LinkDescription = fields[headerMap["LinkDescription"]];
                            existingLink.SelectedDiscipline = fields[headerMap["Discipline"]];
                            existingLink.CompanyName = fields[headerMap["CompanyName"]];
                            existingLink.ResponsiblePerson = fields[headerMap["ResponsiblePerson"]];
                            existingLink.ContactDetails = fields[headerMap["ContactDetails"]];
                            existingLink.Comments = fields[headerMap["Comments"]];
                            updatedCount++;
                        }
                        else
                        {
                            // Create new placeholder
                            var newPlaceholder = new LinkViewModel
                            {
                                LinkName = linkName,
                                LinkDescription = fields[headerMap["LinkDescription"]],
                                SelectedDiscipline = fields[headerMap["Discipline"]],
                                CompanyName = fields[headerMap["CompanyName"]],
                                ResponsiblePerson = fields[headerMap["ResponsiblePerson"]],
                                ContactDetails = fields[headerMap["ContactDetails"]],
                                Comments = fields[headerMap["Comments"]],
                                ReviztoStatus = "Placeholder"
                            };
                            SetPlaceholderProperties(newPlaceholder);
                            Links.Add(newPlaceholder);
                            createdCount++;
                        }
                    }
                    _linksView.Refresh();
                    TaskDialog.Show("Import Complete", $"Profile data imported successfully.\n\n{updatedCount} links updated.\n{createdCount} new placeholders created.");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Import Error", $"An error occurred while importing the profile CSV: {ex.Message}");
            }
        }


        /// <summary>
        /// Escapes a string field for CSV format by adding quotes if necessary.
        /// </summary>
        private string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                // Enclose in double quotes and escape existing double quotes by doubling them
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }


        private static string GetCellValue(IList<string> cellValues, Dictionary<string, int> headerMap, string header) { if (headerMap.TryGetValue(header, out int idx) && idx < cellValues.Count) return cellValues[idx]; return string.Empty; }
        public class ReviztoLinkRecord { public static readonly Guid SchemaGuid = new Guid("A1B2C3D4-E5F6-47A8-9B0C-1234567890AB"); public const string SchemaName = "RTS_ReviztoLinkSchema"; public const string FieldName = "ReviztoLinkJson"; public const string DataStorageName = "RTS_Revizto_Link_Storage"; public string LinkName { get; set; } public string FilePath { get; set; } public string Description { get; set; } public string LastModified { get; set; } }
        private Schema GetOrCreateSchema(Guid schemaGuid, string schemaName, string fieldName) { Schema schema = Schema.Lookup(schemaGuid); if (schema == null) { SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid); schemaBuilder.SetSchemaName(schemaName); schemaBuilder.SetReadAccessLevel(AccessLevel.Public); schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor); schemaBuilder.SetVendorId(VendorId); schemaBuilder.AddSimpleField(fieldName, typeof(string)); schema = schemaBuilder.Finish(); } return schema; }
        private DataStorage GetOrCreateDataStorage(Document doc, string dataStorageElementName) { var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)); DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName); if (dataStorage == null) { dataStorage = DataStorage.Create(doc); dataStorage.Name = dataStorageElementName; } return dataStorage; }
        public void SaveDataToExtensibleStorage<T>(Document doc, List<T> dataList, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) { Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName); DataStorage dataStorage = GetOrCreateDataStorage(doc, dataStorageElementName); string jsonString = JsonSerializer.Serialize(dataList, new JsonSerializerOptions { WriteIndented = true }); Entity entity = new Entity(schema); entity.Set(schema.GetField(fieldName), jsonString); dataStorage.SetEntity(entity); }
        public List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) { Schema schema = Schema.Lookup(schemaGuid); if (schema == null) return new List<T>(); var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)); DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName); if (dataStorage == null) return new List<T>(); Entity entity = dataStorage.GetEntity(schema); if (!entity.IsValid()) return new List<T>(); string jsonString = entity.Get<string>(schema.GetField(fieldName)); if (string.IsNullOrEmpty(jsonString)) return new List<T>(); try { return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>(); } catch (Exception ex) { TaskDialog.Show("Profile Error", $"Failed to read saved profile: {ex.Message}", TaskDialogCommonButtons.Ok); return new List<T>(); } }
    }

    /// <summary>
    /// A view model representing a single row (a single link) in the DataGrid.
    /// </summary>
    public class LinkViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        // Revit-derived properties (not saved in profile)
        [JsonIgnore] public List<ElementId> LinkInstanceIds { get; set; } = new List<ElementId>();
        [JsonIgnore] public ElementId LinkTypeId { get; set; }
        [JsonIgnore] public bool IsRevitLink { get; set; }
        [JsonIgnore] public bool HasValidPath { get; set; }
        [JsonIgnore] public string FullPath { get; set; }
        private string _lastModified;
        [JsonIgnore] public string LastModified { get => _lastModified; set { _lastModified = value; OnPropertyChanged(nameof(LastModified)); } }
        private string _linkWorkset;
        [JsonIgnore] public string LinkWorkset { get => _linkWorkset; set { _linkWorkset = value; OnPropertyChanged(nameof(LinkWorkset)); } }
        private int _numberOfInstances;
        [JsonIgnore] public int NumberOfInstances { get => _numberOfInstances; set { _numberOfInstances = value; OnPropertyChanged(nameof(NumberOfInstances)); } }
        private string _reviztoStatus;
        [JsonIgnore] public string ReviztoStatus { get => _reviztoStatus; set { _reviztoStatus = value; OnPropertyChanged(nameof(ReviztoStatus)); } }
        private string _pathType;
        [JsonIgnore] public string PathType { get => _pathType; set { _pathType = value; OnPropertyChanged(nameof(PathType)); } }
        private string _linkStatus;
        [JsonIgnore] public string LinkStatus { get => _linkStatus; set { _linkStatus = value; OnPropertyChanged(nameof(LinkStatus)); OnPropertyChanged(nameof(UnloadReloadHeader)); } }
        private string _versionStatus;
        [JsonIgnore] public string VersionStatus { get => _versionStatus; set { _versionStatus = value; OnPropertyChanged(nameof(VersionStatus)); } }
        [JsonIgnore] public List<string> AvailableDisciplines { get; } = new List<string> { "Unconfirmed", "Architectural", "Structural", "Mechanical", "Electrical", "Hydraulic", "Fire", "Civil", "Landscape", "Other" };
        [JsonIgnore] public List<string> AvailableCoordinates { get; } = new List<string> { "Unconfirmed", "Origin to Origin", "Shared Coordinates", "N/A" };
        [JsonIgnore] public string UnloadReloadHeader => LinkStatus == "Loaded" ? "Unload" : "Reload";

        // User-editable properties (saved in profile)
        private string _linkName;
        public string LinkName { get => _linkName; set { _linkName = value; OnPropertyChanged(nameof(LinkName)); } }
        private string _linkDescription;
        public string LinkDescription { get => _linkDescription; set { _linkDescription = value; OnPropertyChanged(nameof(LinkDescription)); } }
        private string _selectedDiscipline;
        public string SelectedDiscipline { get => _selectedDiscipline; set { _selectedDiscipline = value; OnPropertyChanged(nameof(SelectedDiscipline)); } }
        private string _selectedCoordinates;
        public string SelectedCoordinates { get => _selectedCoordinates; set { _selectedCoordinates = value; OnPropertyChanged(nameof(SelectedCoordinates)); } }
        private string _companyName;
        public string CompanyName { get => _companyName; set { _companyName = value; OnPropertyChanged(nameof(CompanyName)); } }
        private string _responsiblePerson;
        public string ResponsiblePerson { get => _responsiblePerson; set { _responsiblePerson = value; OnPropertyChanged(nameof(ResponsiblePerson)); } }
        private string _contactDetails;
        public string ContactDetails { get => _contactDetails; set { _contactDetails = value; OnPropertyChanged(nameof(ContactDetails)); } }
        private string _comments;
        public string Comments { get => _comments; set { _comments = value; OnPropertyChanged(nameof(Comments)); } }
    }

    /// <summary>
    /// Dialog for bulk editing LinkViewModel properties.
    /// </summary>
    public partial class BulkEditDialog : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private bool _applyDiscipline;
        public bool ApplyDiscipline { get => _applyDiscipline; set { _applyDiscipline = value; OnPropertyChanged(nameof(ApplyDiscipline)); } }
        private string _selectedDiscipline;
        public string SelectedDiscipline { get => _selectedDiscipline; set { _selectedDiscipline = value; OnPropertyChanged(nameof(SelectedDiscipline)); } }
        private bool _applyCompanyName;
        public bool ApplyCompanyName { get => _applyCompanyName; set { _applyCompanyName = value; OnPropertyChanged(nameof(ApplyCompanyName)); } }
        private string _companyName;
        public string CompanyName { get => _companyName; set { _companyName = value; OnPropertyChanged(nameof(CompanyName)); } }
        private bool _applyComments;
        public bool ApplyComments { get => _applyComments; set { _applyComments = value; OnPropertyChanged(nameof(ApplyComments)); } }
        private string _comments;
        public string Comments { get => _comments; set { _comments = value; OnPropertyChanged(nameof(Comments)); } }
        public List<string> AvailableDisciplines { get; } = new List<string> { "Unconfirmed", "Architectural", "Structural", "Mechanical", "Electrical", "Hydraulic", "Fire", "Civil", "Landscape", "Other" };

        public BulkEditDialog()
        {
            // This requires a BulkEditDialog.xaml file with corresponding controls.
            // For this example, we assume the XAML is set up correctly.
            // InitializeComponent(); 
            this.DataContext = this;
            SelectedDiscipline = "Unconfirmed";
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e) => DialogResult = true;
        private void CancelButton_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
