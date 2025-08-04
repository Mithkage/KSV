// --- FILE: LinkManagerWindow.xaml.cs (UPDATED) ---
//
// File: LinkManagerWindow.xaml.cs
//
// Namespace: RTS.UI
//
// Class: LinkManagerWindow, LinkViewModel
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
// - July 30, 2025: Implemented a more robust, multi-step link finding method with enhanced diagnostics.
// - July 30, 2025: Added ProfileSettingsButton_Click event handler to launch the new settings window.
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
        /// Handles launching the Profile Settings window.
        /// </summary>
        private void ProfileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new ProfileSettingsWindow(_doc, this.Links);
            settingsWindow.Owner = this; // Set owner to center it correctly
            settingsWindow.ShowDialog();
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

            var (allLinkTypes, problematicNames) = GetAllProjectLinkTypes(_doc);
            if (problematicNames.Any())
            {
                TaskDialog.Show("Link Warning", "Could not read the following links (they may be unloaded or corrupt):\n\n" + string.Join("\n", problematicNames));
            }

            var linkInstancesGroupedByType = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().GroupBy(inst => inst.GetTypeId()).ToDictionary(g => g.Key, g => g.ToList());

            // Get available worksets for dropdown
            List<string> availableWorksetNames = new List<string>();
            if (_doc.IsWorkshared)
            {
                availableWorksetNames = new FilteredWorksetCollector(_doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .Select(w => w.Name)
                    .ToList();
            }


            var finalLinkList = new List<LinkViewModel>();
            var processedPlaceholders = new HashSet<string>();

            // 2. Process all live Revit links
            foreach (var linkInfo in allLinkTypes)
            {
                string cleanLinkName = linkInfo.Key;
                RevitLinkType type = linkInfo.Value;

                var viewModel = new LinkViewModel { LinkName = cleanLinkName, IsRevitLink = true, AvailableWorksets = availableWorksetNames };

                // Populate live Revit data
                PopulateRevitLinkData(viewModel, type, linkInstancesGroupedByType);

                // Check for a matching placeholder in the saved profiles to inherit its data
                string linkNameWithoutExt = Path.GetFileNameWithoutExtension(cleanLinkName);
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
                    var savedProfile = savedProfiles.FirstOrDefault(p => p.LinkName.Equals(cleanLinkName, StringComparison.OrdinalIgnoreCase));
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
                savedProfile.AvailableWorksets = availableWorksetNames;
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
                        ReviztoStatus = "Missing",
                        AvailableWorksets = availableWorksetNames
                    };
                    SetPlaceholderProperties(viewModel);
                    finalLinkList.Add(viewModel);
                }
            }

            // 5. Calculate and format coordinates
            FormatCoordinates(finalLinkList);

            // 6. Populate the UI
            Links.Clear();
            foreach (var link in finalLinkList)
            {
                Links.Add(link);
            }
            _linksView.Refresh();
        }

        /// <summary>
        /// Robustly gets all valid Revit Link Types from the project.
        /// </summary>
        private (Dictionary<string, RevitLinkType> types, List<string> problematicNames) GetAllProjectLinkTypes(Document doc)
        {
            var types = new Dictionary<string, RevitLinkType>(StringComparer.OrdinalIgnoreCase);
            var problematicNames = new List<string>();

            var allLinkTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkType))
                .Cast<RevitLinkType>();

            foreach (var linkType in allLinkTypes)
            {
                try
                {
                    string fileName = null;
                    try
                    {
                        var externalRef = linkType.GetExternalFileReference();
                        if (externalRef != null && externalRef.GetPath() != null)
                        {
                            string visiblePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(externalRef.GetPath());
                            fileName = Path.GetFileName(visiblePath);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // Fallback to using the linkType.Name.
                    }

                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = linkType.Name.Split(':').FirstOrDefault()?.Trim();
                    }

                    if (!string.IsNullOrEmpty(fileName) && !types.ContainsKey(fileName))
                    {
                        types.Add(fileName, linkType);
                    }
                    else if (string.IsNullOrEmpty(fileName))
                    {
                        problematicNames.Add($"Link Type ID {linkType.Id} could not be resolved to a filename.");
                    }
                }
                catch (Exception ex)
                {
                    problematicNames.Add($"Link Type ID {linkType.Id} failed with error: {ex.Message}");
                }
            }
            return (types, problematicNames);
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
            catch
            {
                viewModel.PathType = "N/A (Cloud Path)";
            }
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

                // Determine Pin Status
                int pinnedCount = instances.Count(i => i.Pinned);
                if (pinnedCount == 0)
                {
                    viewModel.IsPinned = false;
                }
                else if (pinnedCount == instances.Count)
                {
                    viewModel.IsPinned = true;
                }
                else
                {
                    viewModel.IsPinned = null; // Indeterminate state
                }

                // Store transforms for later processing
                viewModel.InstanceTransforms = instances.Select(i => i.GetTotalTransform()).ToList();
            }
            else
            {
                viewModel.NumberOfInstances = 0;
                viewModel.LinkWorkset = "N/A";
                viewModel.IsPinned = false;
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
            vm.IsPinned = false;
            vm.LinkCoordinates = "N/A";
            vm.SelectedDiscipline ??= "Unconfirmed";
            vm.CompanyName ??= "Not Set";
        }

        /// <summary>
        /// Calculates and formats the coordinate strings for all links based on their instance transforms.
        /// </summary>
        private void FormatCoordinates(List<LinkViewModel> linkViewModels)
        {
            // Error Handling: Check for a valid active project location.
            var activeProjectLocation = _doc.ActiveProjectLocation;
            if (activeProjectLocation == null)
            {
                foreach (var vm in linkViewModels)
                {
                    vm.LinkCoordinates = "Error: Active Project Location not found.";
                }
                return; // Stop processing if the location is missing.
            }

            // This transform converts from the site (shared) coordinate system to the project internal coordinate system.
            // We need the inverse to go from project to shared.
            Autodesk.Revit.DB.Transform projectToSharedTransform = activeProjectLocation.GetTransform().Inverse;


            // Pre-calculate padding by finding the max integer length for each coordinate
            int maxXDigits = 0, maxYDigits = 0, maxZDigits = 0;
            foreach (var vm in linkViewModels.Where(l => l.IsRevitLink && l.InstanceTransforms.Any()))
            {
                foreach (var transform in vm.InstanceTransforms)
                {
                    // The link instance transform is in project coordinates.
                    // We multiply by the inverse of the site transform to get shared coordinates.
                    var finalTransform = projectToSharedTransform.Multiply(transform);
                    var origin = finalTransform.Origin;

                    // Convert from feet to meters for display
                    double xMeters = UnitUtils.ConvertFromInternalUnits(origin.X, UnitTypeId.Meters);
                    double yMeters = UnitUtils.ConvertFromInternalUnits(origin.Y, UnitTypeId.Meters);
                    double zMeters = UnitUtils.ConvertFromInternalUnits(origin.Z, UnitTypeId.Meters);

                    maxXDigits = Math.Max(maxXDigits, ((int)Math.Abs(xMeters)).ToString().Length);
                    maxYDigits = Math.Max(maxYDigits, ((int)Math.Abs(yMeters)).ToString().Length);
                    maxZDigits = Math.Max(maxZDigits, ((int)Math.Abs(zMeters)).ToString().Length);
                }
            }

            // Now, format each link's coordinate string
            foreach (var vm in linkViewModels)
            {
                if (!vm.IsRevitLink || !vm.InstanceTransforms.Any())
                {
                    vm.LinkCoordinates = "N/A";
                    continue;
                }

                try
                {
                    var firstTransform = vm.InstanceTransforms.First();
                    bool allSameLocation = vm.InstanceTransforms.All(t => AreTransformsClose(t, firstTransform));

                    if (allSameLocation)
                    {
                        var finalTransform = projectToSharedTransform.Multiply(firstTransform);
                        var origin = finalTransform.Origin;

                        // Convert from feet to meters for display
                        double xMeters = UnitUtils.ConvertFromInternalUnits(origin.X, UnitTypeId.Meters);
                        double yMeters = UnitUtils.ConvertFromInternalUnits(origin.Y, UnitTypeId.Meters);
                        double zMeters = UnitUtils.ConvertFromInternalUnits(origin.Z, UnitTypeId.Meters);

                        // Calculate rotation in degrees
                        double angle = Math.Atan2(finalTransform.BasisY.X, finalTransform.BasisX.X) * (180.0 / Math.PI);

                        // Create padded format strings
                        string xFormat = new string('0', maxXDigits) + ".00";
                        string yFormat = new string('0', maxYDigits) + ".00";
                        string zFormat = new string('0', maxZDigits) + ".00";

                        vm.LinkCoordinates = $"X: {xMeters.ToString(xFormat)} m | Y: {yMeters.ToString(yFormat)} m | Z: {zMeters.ToString(zFormat)} m | Rotation: {angle:F1}°";
                    }
                    else
                    {
                        vm.LinkCoordinates = "Multiple Locations";
                    }
                }
                catch (Exception)
                {
                    vm.LinkCoordinates = "Error calculating coordinates.";
                }
            }
        }

        /// <summary>
        /// Compares two Revit transforms to see if they are effectively identical within a small tolerance.
        /// </summary>
        private bool AreTransformsClose(Autodesk.Revit.DB.Transform t1, Autodesk.Revit.DB.Transform t2)
        {
            const double tolerance = 1e-6; // A small tolerance for floating-point comparison

            // Compare origins
            if (!t1.Origin.IsAlmostEqualTo(t2.Origin, tolerance)) return false;

            // Compare basis vectors (rotation and scale)
            if (!t1.BasisX.IsAlmostEqualTo(t2.BasisX, tolerance)) return false;
            if (!t1.BasisY.IsAlmostEqualTo(t2.BasisY, tolerance)) return false;
            if (!t1.BasisZ.IsAlmostEqualTo(t2.BasisZ, tolerance)) return false;

            return true;
        }

        /// <summary>
        /// Handles the click event for the "Save" button. Saves the profile and applies any workset and pin changes.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            using (Transaction tx = new Transaction(_doc, "Save Link Manager Changes"))
            {
                tx.Start();
                try
                {
                    // Apply workset and pin changes
                    if (_doc.IsWorkshared)
                    {
                        var worksetTable = _doc.GetWorksetTable();
                        var userWorksets = new FilteredWorksetCollector(_doc).OfKind(WorksetKind.UserWorkset).ToDictionary(ws => ws.Name, ws => ws.Id);

                        foreach (var vm in Links.Where(l => l.IsRevitLink))
                        {
                            WorksetId targetWorksetId = WorksetId.InvalidWorksetId;
                            if (userWorksets.TryGetValue(vm.LinkWorkset, out WorksetId foundId))
                            {
                                targetWorksetId = foundId;
                            }

                            foreach (var instanceId in vm.LinkInstanceIds)
                            {
                                var instance = _doc.GetElement(instanceId);
                                if (instance == null) continue;

                                // Update Workset
                                if (targetWorksetId.IntegerValue != WorksetId.InvalidWorksetId.IntegerValue && instance.WorksetId.IntegerValue != targetWorksetId.IntegerValue)
                                {
                                    var worksetParam = instance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                    if (worksetParam != null && !worksetParam.IsReadOnly)
                                    {
                                        worksetParam.Set(targetWorksetId.IntegerValue);
                                    }
                                }

                                // Update Pin State
                                if (vm.IsPinned.HasValue && instance.Pinned != vm.IsPinned.Value)
                                {
                                    instance.Pinned = vm.IsPinned.Value;
                                }
                            }
                        }
                    }

                    // Save the profile metadata
                    var linksToSave = Links.ToList();
                    SaveDataToExtensibleStorage(_doc, linksToSave, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);

                    tx.Commit();
                    TaskDialog.Show("Success", "Changes saved successfully!", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);

                    // Refresh the grid to show the new state after save
                    LoadLinks();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to save changes: {ex.Message}", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);
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
                var dependencyObject = e.OriginalSource as DependencyObject;
                var row = FindParent<DataGridRow>(dependencyObject);
                if (row == null) return;

                var vm = row.DataContext as LinkViewModel;
                if (vm == null) return;

                // Find the column that was right-clicked
                DataGrid dataGrid = sender as DataGrid;
                DataGridCell cell = FindParent<DataGridCell>(dependencyObject);
                if (cell == null) return;

                DataGridColumn column = cell.Column;
                if (column == null) return;

                // Only show the custom menu for the "File Name" column
                // Adjust the header or SortMemberPath as needed to match your column
                if (column.SortMemberPath == "LinkName" || (column.Header?.ToString() ?? "") == "File Name")
                {
                    var menu = new ContextMenu();

                    var openFileMenuItem = new MenuItem { Header = "Open File Location", Tag = vm };
                    openFileMenuItem.Click += (s, args) =>
                    {
                        if (!string.IsNullOrEmpty(vm.FullPath) && File.Exists(vm.FullPath))
                        {
                            Process.Start("explorer.exe", $"/select,\"{vm.FullPath}\"");
                        }
                        else
                        {
                            MessageBox.Show("File path not found or file does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    };
                    menu.Items.Add(openFileMenuItem);

                    var copyNameMenuItem = new MenuItem { Header = "Copy File Name", Tag = vm };
                    copyNameMenuItem.Click += (s, args) =>
                    {
                        Clipboard.SetText(vm.LinkName ?? "");
                    };
                    menu.Items.Add(copyNameMenuItem);

                    // Assign the context menu to the cell
                    cell.ContextMenu = menu;
                }
                else
                {
                    // Optionally, clear context menu for other columns
                    cell.ContextMenu = null;
                }
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

        private void LinksDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var row = FindParent<DataGridRow>(e.OriginalSource as DependencyObject);
            if (row != null)
            {
                if (!row.IsSelected)
                {
                    row.IsSelected = true;
                }
                // If you want to clear other selections and only select the right-clicked row:
                // var grid = sender as DataGrid;
                // grid.SelectedItem = row.DataContext;
            }
        }


        #endregion

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

        private void PinCheckBox_Click(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as CheckBox;
            var vm = checkBox?.DataContext as LinkViewModel;
            if (vm == null || !vm.IsRevitLink) return;

            // This logic ensures that a click on an indeterminate state makes it checked (true).
            // A click on a checked state makes it unchecked (false).
            // A click on an unchecked state makes it checked (true).
            vm.IsPinned = vm.IsPinned != true;
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
        public string LinkWorkset { get => _linkWorkset; set { _linkWorkset = value; OnPropertyChanged(nameof(LinkWorkset)); } }
        private int _numberOfInstances;
        [JsonIgnore] public int NumberOfInstances { get => _numberOfInstances; set { _numberOfInstances = value; OnPropertyChanged(nameof(NumberOfInstances)); } }
        private string _reviztoStatus;
        [JsonIgnore]
        public string ReviztoStatus
        {
            get => _reviztoStatus;
            set
            {
                _reviztoStatus = value;
                OnPropertyChanged(nameof(ReviztoStatus));
                OnPropertyChanged(nameof(IsAlternativeMatch)); // Notify that the dependent property has also changed
            }
        }
        [JsonIgnore]
        public bool IsAlternativeMatch => !string.IsNullOrEmpty(ReviztoStatus) && ReviztoStatus.StartsWith("Match on:");
        private string _pathType;
        [JsonIgnore] public string PathType { get => _pathType; set { _pathType = value; OnPropertyChanged(nameof(PathType)); } }
        private string _linkStatus;
        [JsonIgnore] public string LinkStatus { get => _linkStatus; set { _linkStatus = value; OnPropertyChanged(nameof(LinkStatus)); OnPropertyChanged(nameof(UnloadReloadHeader)); } }
        private string _versionStatus;
        [JsonIgnore] public string VersionStatus { get => _versionStatus; set { _versionStatus = value; OnPropertyChanged(nameof(VersionStatus)); } }
        [JsonIgnore] public List<string> AvailableDisciplines { get; } = new List<string> { "Unconfirmed", "Architectural", "Structural", "Mechanical", "Electrical", "Hydraulic", "Fire", "Civil", "Landscape", "Other" };
        [JsonIgnore] public List<string> AvailableCoordinates { get; } = new List<string> { "Unconfirmed", "Origin to Origin", "Shared Coordinates", "N/A" };
        [JsonIgnore] public List<string> AvailableWorksets { get; set; }
        [JsonIgnore] public string UnloadReloadHeader => LinkStatus == "Loaded" ? "Unload" : "Reload";

        private bool? _isPinned;
        [JsonIgnore]
        public bool? IsPinned
        {
            get => _isPinned;
            set { _isPinned = value; OnPropertyChanged(nameof(IsPinned)); }
        }

        private string _linkCoordinates;
        [JsonIgnore]
        public string LinkCoordinates
        {
            get => _linkCoordinates;
            set { _linkCoordinates = value; OnPropertyChanged(nameof(LinkCoordinates)); }
        }

        [JsonIgnore] public List<Autodesk.Revit.DB.Transform> InstanceTransforms { get; set; } = new List<Autodesk.Revit.DB.Transform>();


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
}
