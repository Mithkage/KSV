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
// - July 19, 2025: Added specific exception handling for TypeLoadException on Excel export to diagnose version conflicts.
// - July 19, 2025: Implemented dynamic ContextMenu creation in code-behind to resolve runtime errors.
// - July 19, 2025: Added DispatcherUnhandledException handler for improved XAML runtime error diagnostics.
// - July 19, 2025: Added robust icon loading in constructor to prevent runtime errors from missing image resources.
// - July 19, 2025: Corrected Revit 2022 API calls for IsRelativePath and GetLoadStatus to resolve compiler errors.
//

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics; // Required for Process.Start
using System.IO;
using System.Linq;
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
        /// Loads the window icon safely, ignoring errors if the resource is not found.
        /// </summary>
        private void LoadIcon()
        {
            try
            {
                // Use a Pack URI to reference the image resource compiled with the application.
                this.Icon = new BitmapImage(new Uri("pack://application:,,,/Resources/RTS_Icon.png"));
            }
            catch (Exception ex)
            {
                // If the icon resource is missing or there's an error loading it,
                // this will prevent the application from crashing.
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
        /// Scans the Revit document for all RevitLinkInstances and populates the DataGrid.
        /// </summary>
        private void LoadLinks()
        {
            var savedProfiles = RecallDataFromExtensibleStorage<LinkViewModel>(_doc, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
            var savedProfileDict = savedProfiles.ToDictionary(p => p.LinkName, p => p);
            var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
            var reviztoRecordDict = reviztoRecords.ToDictionary(r => r.LinkName, r => r, StringComparer.OrdinalIgnoreCase);

            Links.Clear();
            _activeColumnFilters.Clear();
            _linksView.Refresh();

            var processedLinkNames = new HashSet<string>();
            var collector = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance));
            var linkInstancesGroupedByType = collector.Cast<RevitLinkInstance>().GroupBy(inst => inst.GetTypeId()).ToDictionary(g => g.Key, g => g.ToList());
            var allLinkTypes = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkType)).ToElements();

            foreach (RevitLinkType type in allLinkTypes)
            {
                if (type == null || type.IsNestedLink) continue;

                var viewModel = new LinkViewModel { LinkName = type.Name, IsRevitLink = true };
                processedLinkNames.Add(type.Name);

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
                            // Corrected for Revit 2022 API using System.IO
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

                // Corrected for Revit 2022 API
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
                    viewModel.LinkTypeId = type.Id; // Store the LinkType ID
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

                    // Corrected for Revit 2022 API - using LookupParameter for robustness
                    Parameter sharedSiteParam = instances.First().LookupParameter("Shared Site");
                    if (sharedSiteParam != null && sharedSiteParam.HasValue)
                    {
                        var siteLocation = _doc.GetElement(sharedSiteParam.AsElementId()) as ProjectLocation;
                        if (siteLocation != null)
                        {
                            viewModel.SelectedCoordinates = siteLocation.Name;
                        }
                        else
                        {
                            viewModel.SelectedCoordinates = "Origin to Origin";
                        }
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

                viewModel.ReviztoStatus = reviztoRecordDict.ContainsKey(type.Name) ? "Matched" : "Not Matched";

                if (savedProfileDict.TryGetValue(type.Name, out var savedProfile))
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
                    if (reviztoRecordDict.TryGetValue(type.Name, out var reviztoMatch))
                    {
                        viewModel.LinkDescription = reviztoMatch.Description;
                    }
                    viewModel.SelectedDiscipline = "Unconfirmed";
                    viewModel.CompanyName = "Not Set";
                }
                Links.Add(viewModel);
            }

            foreach (var savedProfile in savedProfiles)
            {
                if (!processedLinkNames.Contains(savedProfile.LinkName))
                {
                    savedProfile.ReviztoStatus = reviztoRecordDict.ContainsKey(savedProfile.LinkName) ? "Missing" : "Placeholder";
                    savedProfile.LastModified = reviztoRecordDict.TryGetValue(savedProfile.LinkName, out var rec) ? rec.LastModified : "N/A";
                    SetPlaceholderProperties(savedProfile);
                    Links.Add(savedProfile);
                    processedLinkNames.Add(savedProfile.LinkName);
                }
            }

            foreach (var reviztoRecord in reviztoRecords)
            {
                if (!processedLinkNames.Contains(reviztoRecord.LinkName))
                {
                    var viewModel = new LinkViewModel
                    {
                        LinkName = reviztoRecord.LinkName,
                        LinkDescription = reviztoRecord.Description,
                        LastModified = reviztoRecord.LastModified,
                        ReviztoStatus = "Missing"
                    };
                    SetPlaceholderProperties(viewModel);
                    Links.Add(viewModel);
                }
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
        /// Handles the click event for the "Save Profile" button.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            using (Transaction tx = new Transaction(_doc, "Save Link Manager Profile"))
            {
                tx.Start();
                try
                {
                    var linksToSave = Links.Where(vm => vm.IsRevitLink || vm.ReviztoStatus == "Placeholder").ToList();
                    SaveDataToExtensibleStorage(_doc, linksToSave, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
                    tx.Commit();
                    TaskDialog.Show("Success", "Project link profile saved successfully!", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);
                    this.DialogResult = true;
                    this.Close();
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
        /// Handles deleting placeholders with the Delete key.
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
                        var result = TaskDialog.Show("Confirm Delete", $"Are you sure you want to delete {placeholdersToDelete.Count} placeholder(s)?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                        if (result == TaskDialogResult.Yes)
                        {
                            foreach (var p in placeholdersToDelete) Links.Remove(p);
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
                var fe = e.Source as FrameworkElement;
                if (fe == null) return;

                var vm = fe.DataContext as LinkViewModel;
                if (vm == null) return;

                var menu = new ContextMenu();
                menu.Tag = vm; // Store the ViewModel in the Tag for easy access in click handlers

                var openFileItem = new MenuItem { Header = "Open Linked File", IsEnabled = vm.HasValidPath };
                openFileItem.Click += OpenLinkedFile_Click;
                menu.Items.Add(openFileItem);

                var openFolderItem = new MenuItem { Header = "Open Containing Folder", IsEnabled = vm.HasValidPath };
                openFolderItem.Click += OpenContainingFolder_Click;
                menu.Items.Add(openFolderItem);

                menu.Items.Add(new Separator());

                var relinkItem = new MenuItem { Header = "Relink...", IsEnabled = vm.IsRevitLink };
                relinkItem.Click += Relink_Click;
                menu.Items.Add(relinkItem);

                var unloadReloadItem = new MenuItem { Header = vm.UnloadReloadHeader, IsEnabled = vm.IsRevitLink };
                unloadReloadItem.Click += UnloadReload_Click;
                menu.Items.Add(unloadReloadItem);

                var removeItem = new MenuItem { Header = "Remove Link", IsEnabled = vm.IsRevitLink };
                removeItem.Click += RemoveLink_Click;
                menu.Items.Add(removeItem);

                menu.Items.Add(new Separator());

                var copyNameItem = new MenuItem { Header = "Copy Link Name" };
                copyNameItem.Click += CopyLinkName_Click;
                menu.Items.Add(copyNameItem);

                var copyPathItem = new MenuItem { Header = "Copy Link Path", IsEnabled = vm.HasValidPath };
                copyPathItem.Click += CopyLinkPath_Click;
                menu.Items.Add(copyPathItem);

                fe.ContextMenu = menu;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating context menu: {ex.Message}");
                e.Handled = true; // Prevent the default (and possibly crashing) context menu from showing
            }
        }


        private LinkViewModel GetLinkViewModelFromSender(object sender)
        {
            if (sender is MenuItem menuItem && menuItem.Parent is ContextMenu contextMenu)
            {
                // Get the LinkViewModel from the Tag property
                return contextMenu.Tag as LinkViewModel;
            }
            return null;
        }

        private void OpenLinkedFile_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm != null && vm.HasValidPath)
            {
                try
                {
                    Process.Start(new ProcessStartInfo(vm.FullPath) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Could not open file: {ex.Message}");
                }
            }
        }

        private void OpenContainingFolder_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm != null && vm.HasValidPath)
            {
                try
                {
                    Process.Start("explorer.exe", $"/select,\"{vm.FullPath}\"");
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Could not open folder: {ex.Message}");
                }
            }
        }

        private void RemoveLink_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm != null && vm.IsRevitLink)
            {
                var result = TaskDialog.Show("Confirm Remove Link",
                    $"Are you sure you want to permanently remove the link '{vm.LinkName}' from the project? This action cannot be undone.",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (result == TaskDialogResult.Yes)
                {
                    using (var tx = new Transaction(_doc, "Remove Revit Link"))
                    {
                        try
                        {
                            tx.Start();
                            _doc.Delete(vm.LinkTypeId);
                            tx.Commit();
                            TaskDialog.Show("Success", $"Link '{vm.LinkName}' was removed successfully.");
                            LoadLinks(); // Refresh the grid
                        }
                        catch (Exception ex)
                        {
                            tx.RollBack();
                            TaskDialog.Show("Error", $"Failed to remove link: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void CopyLinkName_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm != null)
            {
                Clipboard.SetText(vm.LinkName);
            }
        }

        private void CopyLinkPath_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm != null && vm.HasValidPath)
            {
                Clipboard.SetText(vm.FullPath);
            }
        }

        private void Relink_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm == null || !vm.IsRevitLink) return;

            RevitLinkType linkType = _doc.GetElement(vm.LinkTypeId) as RevitLinkType;
            if (linkType == null)
            {
                TaskDialog.Show("Error", "Could not find the Revit Link Type.");
                return;
            }

            var openFileDialog = new OpenFileDialog { Filter = "Revit Files (*.rvt)|*.rvt", Title = $"Select New File for {vm.LinkName}" };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    ModelPath newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(openFileDialog.FileName);
                    using (var tx = new Transaction(_doc, "Relink"))
                    {
                        tx.Start();
                        // For Revit 2022, LoadFrom is a robust way to relink
                        linkType.LoadFrom(newModelPath, new WorksetConfiguration());
                        tx.Commit();
                    }
                    TaskDialog.Show("Success", $"{vm.LinkName} relinked successfully.");
                    LoadLinks();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Relink Error", $"Failed to relink: {ex.Message}");
                }
            }
        }

        private void UnloadReload_Click(object sender, RoutedEventArgs e)
        {
            var vm = GetLinkViewModelFromSender(sender);
            if (vm == null || !vm.IsRevitLink) return;

            RevitLinkType linkType = _doc.GetElement(vm.LinkTypeId) as RevitLinkType;
            if (linkType == null)
            {
                TaskDialog.Show("Error", "Could not find the Revit Link Type.");
                return;
            }

            // Corrected for Revit 2022 API
            string action = RevitLinkType.IsLoaded(_doc, linkType.Id) ? "Unload" : "Reload";
            try
            {
                using (var tx = new Transaction(_doc, $"{action} Link"))
                {
                    tx.Start();
                    if (action == "Unload")
                    {
                        linkType.Unload(null);
                    }
                    else
                    {
                        // Load() is sufficient to reload if the path is known
                        linkType.Load();
                    }
                    tx.Commit();
                }
                LoadLinks();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to {action.ToLower()}: {ex.Message}");
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

        private void ImportReviztoButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*", Title = "Select Revizto Link Export Excel File" };
            if (openFileDialog.ShowDialog() != true) return;
            string filePath = openFileDialog.FileName;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) { TaskDialog.Show("Error", "Invalid file path or file does not exist.", TaskDialogCommonButtons.Ok); return; }
            try
            {
                var reviztoLinks = new List<ReviztoLinkRecord>();
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheets.First();
                    var headerRow = worksheet.FirstRowUsed().RowUsed();
                    var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToList();
                    var headerMap = headers.Select((h, i) => new { h, i }).ToDictionary(x => x.h, x => x.i, StringComparer.OrdinalIgnoreCase);
                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var cellValues = row.Cells().Select(c => c.GetString().Trim()).ToList();
                        var record = new ReviztoLinkRecord
                        {
                            LinkName = GetCellValue(cellValues, headerMap, "Link Name"),
                            FilePath = GetCellValue(cellValues, headerMap, "File Path"),
                            Description = GetCellValue(cellValues, headerMap, "Description"),
                            LastModified = GetCellValue(cellValues, headerMap, "Last Modified"),
                        };
                        if (!string.IsNullOrWhiteSpace(record.LinkName)) reviztoLinks.Add(record);
                    }
                }
                using (var tx = new Autodesk.Revit.DB.Transaction(_doc, "Import Revizto Link Data"))
                {
                    tx.Start();
                    var schemaGuid = ReviztoLinkRecord.SchemaGuid;
                    var dataStorageName = ReviztoLinkRecord.DataStorageName;
                    var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
                    var existing = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageName);
                    if (existing != null) _doc.Delete(existing.Id);
                    SaveDataToExtensibleStorage(_doc, reviztoLinks, schemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, dataStorageName);
                    tx.Commit();
                }
                TaskDialog.Show("Import Complete", "Revizto link data imported and saved successfully.", TaskDialogCommonButtons.Ok);
                LoadLinks();
            }
            catch (Exception ex) { TaskDialog.Show("Import Error", $"Failed to import Revizto data: {ex.Message}", TaskDialogCommonButtons.Ok); }
        }

        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx", FileName = "RevitLinkManagerExport.xlsx", Title = "Save Link Data to Excel" };
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Link Data");
                        int colIdx = 1;
                        foreach (DataGridColumn column in LinksDataGrid.Columns)
                        {
                            if (column.Visibility == System.Windows.Visibility.Visible)
                            {
                                worksheet.Cell(1, colIdx).Value = column.Header.ToString().Replace("&#x0a;", Environment.NewLine);
                                colIdx++;
                            }
                        }
                        int rowIdx = 2;
                        foreach (LinkViewModel link in LinksDataGrid.Items)
                        {
                            colIdx = 1;
                            foreach (DataGridColumn column in LinksDataGrid.Columns)
                            {
                                if (column.Visibility == System.Windows.Visibility.Visible)
                                {
                                    // Fully qualify Binding to resolve ambiguity
                                    string propertyName = (column as DataGridBoundColumn)?.Binding is System.Windows.Data.Binding binding ? binding.Path.Path : column.SortMemberPath;
                                    if (!string.IsNullOrEmpty(propertyName))
                                    {
                                        var prop = typeof(LinkViewModel).GetProperty(propertyName);
                                        if (prop != null)
                                        {
                                            worksheet.Cell(rowIdx, colIdx).Value = prop.GetValue(link)?.ToString();
                                        }
                                    }
                                    colIdx++;
                                }
                            }
                            rowIdx++;
                        }
                        workbook.SaveAs(saveFileDialog.FileName);
                    }
                    TaskDialog.Show("Export Complete", "Link data exported to Excel successfully.", TaskDialogCommonButtons.Ok);
                }
                catch (TypeLoadException tlEx)
                {
                    string errorMessage = "Failed to export data due to a library version conflict.\n\n" +
                                          "This error often occurs if a dependency of the 'ClosedXML' library (like 'ExcelNumberFormat') is missing or the wrong version. Please ensure all required NuGet packages are installed and up to date.\n\n" +
                                          $"Error Details: {tlEx.Message}";
                    TaskDialog.Show("Export Error - Version Conflict", errorMessage);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Export Error", $"Failed to export data to Excel: {ex.Message}", TaskDialogCommonButtons.Ok);
                }
            }
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
