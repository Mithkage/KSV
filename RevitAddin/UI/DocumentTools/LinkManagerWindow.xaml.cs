// --- FILE: LinkManagerWindow.xaml.cs (UPDATED) ---
//
// File: LinkManagerWindow.xaml.cs
// Namespace: RTS.UI
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// This file contains the code-behind for the Link Manager window. It handles loading,
// displaying, and saving Revit link metadata to Extensible Storage.
//
// Log:
// - 2025-08-11: Corrected WorksetId API usage to be compatible with both Revit 2022 and 2024.
// - 2025-08-11: Updated methods relying on ProfileSettings to use the new data structure.
// - 2025-08-11: Corrected duplicate key exception in LoadLinksInBackground by grouping items before creating dictionaries.
// - 2025-08-06: Implemented asynchronous data loading for better UI responsiveness.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.FileIO; // Required for TextFieldParser
using Microsoft.Win32;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics; // Required for Process.Start
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text; // Required for StringBuilder
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions; // Required for Regex
using System.Threading.Tasks;
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

        // Extensible Storage definitions for the Link Manager profile.
        public static readonly Guid ProfileSchemaGuid = new Guid("D4C6E8B0-6A2E-4B1C-9D7E-8C4F2A6B9E1F");
        public const string ProfileSchemaName = "RTS_LinkManagerProfileSchema";
        public const string ProfileFieldName = "LinkManagerProfileJson";
        public const string ProfileDataStorageElementName = "RTS_LinkManager_Profile_Storage";
        public const string VendorId = "ReTick_Solutions";

        public ObservableCollection<LinkViewModel> Links { get; set; }
        public ObservableCollection<ProjectRoleContact> ProjectRoles { get; set; }

        public LinkManagerWindow(Document doc)
        {
            this.Dispatcher.UnhandledException += OnDispatcherUnhandledException;

            InitializeComponent();
            LoadIcon();
            _doc = doc;
            this.DataContext = this;

            Links = new ObservableCollection<LinkViewModel>();
            _linksView = CollectionViewSource.GetDefaultView(Links);
            _linksView.Filter = new Predicate<object>(FilterLinks);

            ProjectRoles = new ObservableCollection<ProjectRoleContact>
            {
                new ProjectRoleContact { Role = "Project Manager" },
                new ProjectRoleContact { Role = "BIM Manager" },
                new ProjectRoleContact { Role = "BIM Coordinator" }
            };

            // Asynchronously load initial data to keep the UI responsive.
            _ = LoadDataAsync();
        }

        /// <summary>
        /// Launches the Profile Settings window.
        /// </summary>
        private void ProfileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new ProfileSettingsWindow(_doc, this.Links);
            settingsWindow.Owner = this;
            settingsWindow.ShowDialog();
        }

        /// <summary>
        /// Provides detailed exception info for UI thread errors.
        /// </summary>
        void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            string errorMessage = $"An unexpected error occurred: {e.Exception.Message}";
            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\nInner Exception: {e.Exception.InnerException.Message}";
            }
            TaskDialog.Show("Runtime Error", errorMessage);
        }

        /// <summary>
        /// Loads the window icon using a robust Pack URI.
        /// </summary>
        private void LoadIcon()
        {
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.png", UriKind.Absolute);
                this.Icon = new BitmapImage(iconUri);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning: Could not load window icon. {ex.Message}");
            }
        }

        /// <summary>
        /// Filters the DataGrid based on active column filters.
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
        /// Displays a filter context menu when a column header is clicked.
        /// </summary>
        private void ColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var header = FindParent<DataGridColumnHeader>(button);
            if (header == null) return;

            string sortMemberPath = header.Column.SortMemberPath;
            if (string.IsNullOrEmpty(sortMemberPath)) return;

            var uniqueValues = Links.Select(l => typeof(LinkViewModel).GetProperty(sortMemberPath).GetValue(l)?.ToString())
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToList();

            var popup = new Popup
            {
                PlacementTarget = button,
                Placement = PlacementMode.Bottom,
                StaysOpen = false,
                AllowsTransparency = true,
                Width = 250,
                Height = 300
            };

            var stackPanel = new StackPanel { Background = Brushes.White, Margin = new Thickness(5) };
            var searchBox = new System.Windows.Controls.TextBox { Margin = new Thickness(0, 0, 0, 5) };
            var listBox = new ListBox { Height = 220 };

            void UpdateList()
            {
                string pattern = searchBox.Text.Trim();
                if (string.IsNullOrEmpty(pattern))
                {
                    listBox.ItemsSource = uniqueValues;
                }
                else
                {
                    string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
                    var regex = new Regex(regexPattern, RegexOptions.IgnoreCase);
                    listBox.ItemsSource = uniqueValues.Where(v => regex.IsMatch(v));
                }
            }

            searchBox.TextChanged += (s, args) => UpdateList();
            UpdateList();

            listBox.SelectionChanged += (s, args) =>
            {
                if (listBox.SelectedItem is string selected)
                {
                    _activeColumnFilters[sortMemberPath] = selected;
                    _linksView.Refresh();
                    header.Background = Brushes.LightBlue;
                    popup.IsOpen = false;
                }
            };

            var clearButton = new Button { Content = "Clear Filter", Margin = new Thickness(0, 5, 0, 0), Background = Brushes.LightGray };
            clearButton.Click += (s, args) =>
            {
                _activeColumnFilters.Remove(sortMemberPath);
                _linksView.Refresh();
                header.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                popup.IsOpen = false;
            };

            stackPanel.Children.Add(searchBox);
            stackPanel.Children.Add(listBox);
            stackPanel.Children.Add(clearButton);

            popup.Child = stackPanel;
            popup.IsOpen = true;
        }

        /// <summary>
        /// Asynchronously loads link data on a background thread.
        /// </summary>
        private async Task LoadDataAsync()
        {
            try
            {
                var newLinks = await Task.Run(() => LoadLinksInBackground());

                Links.Clear();
                foreach (var link in newLinks)
                {
                    Links.Add(link);
                }
                _linksView.Refresh();

                // If no links are found, inform the user.
                if (!Links.Any())
                {
                    TaskDialog.Show("Link Manager", "No Revit links were found in the current project, and no saved link profile exists.\n\nTo get started, you can:\n  - Load a Revit link into the project and reopen this tool.\n  - Use the 'Add Placeholder' button to manually create a profile entry.");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Loading Links", $"An unexpected error occurred while loading link data: {ex.Message}");
            }
        }

        /// <summary>
        /// Performs the data-intensive work of loading and merging link info.
        /// </summary>
        private List<LinkViewModel> LoadLinksInBackground()
        {
            var savedProfiles = RecallDataFromExtensibleStorage<LinkViewModel>(_doc, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
            var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
            var (allLinkTypes, problematicNames) = GetAllProjectLinkTypes(_doc);

            if (problematicNames.Any())
            {
                Dispatcher.Invoke(() => TaskDialog.Show("Link Warning", "Could not read the following links (they may be unloaded or corrupt):\n\n" + string.Join("\n", problematicNames)));
            }

            var linkInstancesGroupedByType = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().GroupBy(inst => inst.GetTypeId()).ToDictionary(g => g.Key, g => g.ToList());
            var availableWorksetNames = _doc.IsWorkshared ? new FilteredWorksetCollector(_doc).OfKind(WorksetKind.UserWorkset).Select(w => w.Name).ToList() : new List<string>();

            var savedProfilesDict = savedProfiles
                .GroupBy(p => p.LinkName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var reviztoRecordsDict = reviztoRecords
                .GroupBy(r => Path.GetFileNameWithoutExtension(r.LinkName), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(r => {
                        DateTime.TryParse(r.LastModified, out DateTime dt);
                        return dt;
                    }).First(),
                    StringComparer.OrdinalIgnoreCase);

            var finalLinkList = new List<LinkViewModel>();
            var processedProfileKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            ProcessRevitLinks(allLinkTypes, savedProfilesDict, availableWorksetNames, linkInstancesGroupedByType, finalLinkList, processedProfileKeys);
            ProcessPlaceholders(savedProfilesDict, availableWorksetNames, finalLinkList, processedProfileKeys);
            IntegrateReviztoData(finalLinkList, reviztoRecordsDict, availableWorksetNames);
            FormatCoordinates(finalLinkList);

            return finalLinkList;
        }

        /// <summary>
        /// Processes live Revit links and merges them with saved profile data.
        /// </summary>
        private void ProcessRevitLinks(Dictionary<string, RevitLinkType> allLinkTypes, Dictionary<string, LinkViewModel> savedProfilesDict, List<string> availableWorksetNames, Dictionary<ElementId, List<RevitLinkInstance>> linkInstancesGroupedByType, List<LinkViewModel> finalLinkList, HashSet<string> processedProfileKeys)
        {
            foreach (var linkInfo in allLinkTypes)
            {
                string cleanLinkName = linkInfo.Key;
                RevitLinkType type = linkInfo.Value;

                var viewModel = new LinkViewModel { LinkName = cleanLinkName, IsRevitLink = true, AvailableWorksets = availableWorksetNames };
                PopulateRevitLinkData(viewModel, type, linkInstancesGroupedByType);

                string linkNameWithoutExt = Path.GetFileNameWithoutExtension(cleanLinkName);
                var matchingPlaceholder = savedProfilesDict.Values.FirstOrDefault(p => !p.IsRevitLink && Path.GetFileNameWithoutExtension(p.LinkName).Equals(linkNameWithoutExt, StringComparison.OrdinalIgnoreCase));

                if (matchingPlaceholder != null)
                {
                    viewModel.ApplyProfileData(matchingPlaceholder);
                    processedProfileKeys.Add(matchingPlaceholder.LinkName);
                }
                else if (savedProfilesDict.TryGetValue(cleanLinkName, out var savedProfile))
                {
                    viewModel.ApplyProfileData(savedProfile);
                    processedProfileKeys.Add(savedProfile.LinkName);
                }
                else
                {
                    viewModel.SelectedDiscipline = "Unconfirmed";
                    viewModel.CompanyName = "Not Set";
                }
                finalLinkList.Add(viewModel);
            }
        }

        /// <summary>
        /// Processes saved profiles that were not matched to a live Revit link.
        /// </summary>
        private void ProcessPlaceholders(Dictionary<string, LinkViewModel> savedProfilesDict, List<string> availableWorksetNames, List<LinkViewModel> finalLinkList, HashSet<string> processedProfileKeys)
        {
            foreach (var savedProfile in savedProfilesDict.Values)
            {
                if (!processedProfileKeys.Contains(savedProfile.LinkName))
                {
                    SetPlaceholderProperties(savedProfile);
                    savedProfile.AvailableWorksets = availableWorksetNames;
                    finalLinkList.Add(savedProfile);
                }
            }
        }

        /// <summary>
        /// Integrates Revizto link data into the final list.
        /// </summary>
        private void IntegrateReviztoData(List<LinkViewModel> finalLinkList, Dictionary<string, ReviztoLinkRecord> reviztoRecordsDict, List<string> availableWorksetNames)
        {
            var processedReviztoKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var linkVm in finalLinkList)
            {
                if (linkVm.IsRevitLink)
                {
                    string linkNameWithoutExt = Path.GetFileNameWithoutExtension(linkVm.LinkName);
                    if (reviztoRecordsDict.TryGetValue(linkNameWithoutExt, out var reviztoMatch))
                    {
                        linkVm.LastModified = reviztoMatch.LastModified;
                        string reviztoExt = Path.GetExtension(reviztoMatch.LinkName);
                        string linkExt = Path.GetExtension(linkVm.LinkName);
                        linkVm.ReviztoStatus = linkExt.Equals(reviztoExt, StringComparison.OrdinalIgnoreCase) ? "Matched" : $"Match on: {reviztoExt}";
                        processedReviztoKeys.Add(linkNameWithoutExt);
                    }
                    else
                    {
                        linkVm.ReviztoStatus = "Not Matched";
                    }
                }
            }

            foreach (var reviztoRecord in reviztoRecordsDict.Values)
            {
                string reviztoNameWithoutExt = Path.GetFileNameWithoutExtension(reviztoRecord.LinkName);
                if (!processedReviztoKeys.Contains(reviztoNameWithoutExt))
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
        }

        /// <summary>
        /// Gets all valid Revit Link Types from the project.
        /// </summary>
        private (Dictionary<string, RevitLinkType> types, List<string> problematicNames) GetAllProjectLinkTypes(Document doc)
        {
            var types = new Dictionary<string, RevitLinkType>(StringComparer.OrdinalIgnoreCase);
            var problematicNames = new List<string>();
            var allLinkTypes = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>();

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
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { /* Fallback to name */ }

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
                        problematicNames.Add($"Link Type ID {linkType.Id} could not be resolved.");
                    }
                }
                catch (Exception ex)
                {
                    problematicNames.Add($"Link Type ID {linkType.Id} failed: {ex.Message}");
                }
            }
            return (types, problematicNames);
        }

        /// <summary>
        /// Populates a LinkViewModel with data from a live RevitLinkType.
        /// </summary>
        private void PopulateRevitLinkData(LinkViewModel viewModel, RevitLinkType type, Dictionary<ElementId, List<RevitLinkInstance>> linkInstancesGroupedByType)
        {
            try
            {
                ModelPath modelPath = type.GetExternalFileReference()?.GetPath();
                if (modelPath != null)
                {
                    string linkPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    if (!string.IsNullOrEmpty(linkPath) && File.Exists(linkPath))
                    {
                        viewModel.FullPath = Path.GetFullPath(linkPath);
                        viewModel.LastModified = File.GetLastWriteTime(viewModel.FullPath).ToString("yyyy-MM-dd HH:mm");
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

            viewModel.LinkStatus = RevitLinkType.IsLoaded(_doc, type.Id) ? "Loaded" : "Unloaded";

            if (viewModel.LinkStatus == "Loaded" && viewModel.HasValidPath)
            {
                DateTime fileLastWriteTime = File.GetLastWriteTime(viewModel.FullPath);
                if (DateTime.TryParse(viewModel.LastModified, out DateTime revitLinkLastWriteTime))
                {
                    viewModel.VersionStatus = fileLastWriteTime > revitLinkLastWriteTime ? "Newer local file" : "Up-to-date";
                }
            }
            else
            {
                viewModel.VersionStatus = "N/A";
            }

            if (linkInstancesGroupedByType.TryGetValue(type.Id, out var instances) && instances.Any())
            {
                viewModel.LinkInstanceIds = instances.Select(i => i.Id).ToList();
                viewModel.LinkTypeId = type.Id;
                viewModel.NumberOfInstances = instances.Count;
                viewModel.InstanceTransforms = instances.Select(i => i.GetTotalTransform()).ToList();

                if (_doc.IsWorkshared)
                {
                    var worksetNames = instances.Select(i => _doc.GetWorksetTable().GetWorkset(i.WorksetId)?.Name).Where(n => n != null).Distinct().ToList();
                    viewModel.LinkWorkset = worksetNames.Count > 1 ? "Multiple Worksets" : worksetNames.FirstOrDefault() ?? "Unknown";
                }
                else
                {
                    viewModel.LinkWorkset = "N/A";
                }

                int pinnedCount = instances.Count(i => i.Pinned);
                viewModel.IsPinned = (pinnedCount == instances.Count) ? true : (pinnedCount == 0) ? false : (bool?)null;
            }
            else
            {
                viewModel.NumberOfInstances = 0;
                viewModel.LinkWorkset = "N/A";
                viewModel.IsPinned = false;
            }
        }

        /// <summary>
        /// Sets default properties for a placeholder link.
        /// </summary>
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
        /// Calculates and formats coordinate strings for all links.
        /// </summary>
        private void FormatCoordinates(List<LinkViewModel> linkViewModels)
        {
            var activeProjectLocation = _doc.ActiveProjectLocation;
            if (activeProjectLocation == null)
            {
                foreach (var vm in linkViewModels) vm.LinkCoordinates = "Error: Active Project Location not found.";
                return;
            }

            Autodesk.Revit.DB.Transform projectToSharedTransform = activeProjectLocation.GetTransform().Inverse;

            int maxXDigits = 0, maxYDigits = 0, maxZDigits = 0;
            foreach (var vm in linkViewModels.Where(l => l.IsRevitLink && l.InstanceTransforms.Any()))
            {
                foreach (var transform in vm.InstanceTransforms)
                {
                    var origin = projectToSharedTransform.Multiply(transform).Origin;
                    double xMeters = UnitUtils.ConvertFromInternalUnits(origin.X, UnitTypeId.Meters);
                    double yMeters = UnitUtils.ConvertFromInternalUnits(origin.Y, UnitTypeId.Meters);
                    double zMeters = UnitUtils.ConvertFromInternalUnits(origin.Z, UnitTypeId.Meters);
                    maxXDigits = Math.Max(maxXDigits, ((int)Math.Abs(xMeters)).ToString().Length);
                    maxYDigits = Math.Max(maxYDigits, ((int)Math.Abs(yMeters)).ToString().Length);
                    maxZDigits = Math.Max(maxZDigits, ((int)Math.Abs(zMeters)).ToString().Length);
                }
            }

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
                    if (vm.InstanceTransforms.All(t => AreTransformsClose(t, firstTransform)))
                    {
                        var finalTransform = projectToSharedTransform.Multiply(firstTransform);
                        var origin = finalTransform.Origin;
                        double xMeters = UnitUtils.ConvertFromInternalUnits(origin.X, UnitTypeId.Meters);
                        double yMeters = UnitUtils.ConvertFromInternalUnits(origin.Y, UnitTypeId.Meters);
                        double zMeters = UnitUtils.ConvertFromInternalUnits(origin.Z, UnitTypeId.Meters);
                        double angle = Math.Atan2(finalTransform.BasisY.X, finalTransform.BasisX.X) * (180.0 / Math.PI);
                        vm.LinkCoordinates = $"X: {xMeters.ToString(new string('0', maxXDigits) + ".00")} m | Y: {yMeters.ToString(new string('0', maxYDigits) + ".00")} m | Z: {zMeters.ToString(new string('0', maxZDigits) + ".00")} m | Rotation: {angle:F1}°";
                    }
                    else
                    {
                        vm.LinkCoordinates = "Multiple Locations";
                    }
                }
                catch { vm.LinkCoordinates = "Error calculating coordinates."; }
            }
        }

        /// <summary>
        /// Compares two Revit transforms for effective equality.
        /// </summary>
        private bool AreTransformsClose(Autodesk.Revit.DB.Transform t1, Autodesk.Revit.DB.Transform t2)
        {
            const double tolerance = 1e-6;
            if (!t1.Origin.IsAlmostEqualTo(t2.Origin, tolerance)) return false;
            if (!t1.BasisX.IsAlmostEqualTo(t2.BasisX, tolerance)) return false;
            if (!t1.BasisY.IsAlmostEqualTo(t2.BasisY, tolerance)) return false;
            if (!t1.BasisZ.IsAlmostEqualTo(t2.BasisZ, tolerance)) return false;
            return true;
        }

        /// <summary>
        /// Saves the profile and applies any pending model changes.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            int moveSuccess = 0, moveFail = 0;
            var moveErrors = new List<string>();

            foreach (var vm in Links.Where(l => l.PendingMoveToSharedCoordinates))
            {
                using (Transaction tx = new Transaction(_doc, $"Move Link '{vm.LinkName}' to Shared Coordinates"))
                {
                    tx.Start();
                    try
                    {
                        Autodesk.Revit.DB.Transform sharedCoordTransform = GetSharedCoordinatesTransform();
                        if (sharedCoordTransform == null) throw new Exception("Shared Coordinates transform not found.");

                        foreach (var id in vm.LinkInstanceIds)
                        {
                            var inst = _doc.GetElement(id) as RevitLinkInstance;
                            if (inst?.Location is LocationPoint location)
                            {
                                bool wasPinned = inst.Pinned;
                                if (wasPinned) inst.Pinned = false;

                                XYZ targetPosition = sharedCoordTransform.Origin;
                                location.Move(targetPosition - location.Point);

                                double targetAngle = Math.Atan2(sharedCoordTransform.BasisX.Y, sharedCoordTransform.BasisX.X);
                                double angleDelta = targetAngle - location.Rotation;
                                if (Math.Abs(angleDelta) > 1e-6)
                                {
                                    location.Rotate(Line.CreateBound(targetPosition, targetPosition + XYZ.BasisZ), angleDelta);
                                }
                                inst.Pinned = vm.IsPinned ?? wasPinned;
                            }
                        }
                        tx.Commit();
                        moveSuccess++;
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        moveFail++;
                        moveErrors.Add($"{vm.LinkName}: {ex.Message}");
                    }
                    vm.PendingMoveToSharedCoordinates = false;
                }
            }

            _ = LoadDataAsync();

            using (Transaction tx = new Transaction(_doc, "Save Link Manager Changes"))
            {
                tx.Start();
                try
                {
                    if (_doc.IsWorkshared)
                    {
                        var userWorksets = new FilteredWorksetCollector(_doc).OfKind(WorksetKind.UserWorkset).ToDictionary(ws => ws.Name, ws => ws.Id);
                        foreach (var vm in Links.Where(l => l.IsRevitLink))
                        {
                            if (userWorksets.TryGetValue(vm.LinkWorkset, out WorksetId targetWorksetId))
                            {
                                foreach (var instanceId in vm.LinkInstanceIds)
                                {
                                    var instance = _doc.GetElement(instanceId);
                                    if (instance == null) continue;

                                    // CORRECTED: WorksetId uses .IntegerValue for both Revit 2022 and 2024
                                    if (instance.WorksetId.IntegerValue != targetWorksetId.IntegerValue)
                                    {
                                        instance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)?.Set(targetWorksetId.IntegerValue);
                                    }

                                    if (vm.IsPinned.HasValue && instance.Pinned != vm.IsPinned.Value)
                                    {
                                        instance.Pinned = vm.IsPinned.Value;
                                    }
                                }
                            }
                        }
                    }

                    SaveDataToExtensibleStorage(_doc, Links.ToList(), ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
                    tx.Commit();
                    TaskDialog.Show("Success", "Changes saved successfully!");
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to save changes: {ex.Message}");
                }
            }

            if (moveSuccess > 0 || moveFail > 0)
            {
                var msg = $"Move to Shared Coordinates complete.\n\nSuccess: {moveSuccess}\nFailed: {moveFail}";
                if (moveErrors.Any()) msg += "\n\nErrors:\n" + string.Join("\n", moveErrors);
                TaskDialog.Show("Move Summary", msg);
            }
        }

        /// <summary>
        /// Adds a new placeholder row to the DataGrid.
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
        /// Deletes selected placeholder rows on Delete key press.
        /// </summary>
        private void LinksDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && LinksDataGrid.SelectedItems.Count > 0)
            {
                var placeholdersToDelete = LinksDataGrid.SelectedItems.Cast<LinkViewModel>().Where(vm => !vm.IsRevitLink).ToList();
                if (placeholdersToDelete.Any())
                {
                    var result = TaskDialog.Show("Confirm Delete", $"Permanently delete {placeholdersToDelete.Count} placeholder(s)?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                    if (result == TaskDialogResult.Yes)
                    {
                        using (var tx = new Transaction(_doc, "Delete Placeholders"))
                        {
                            tx.Start();
                            try
                            {
                                var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                                var namesToDelete = new HashSet<string>(placeholdersToDelete.Select(p => p.LinkName), StringComparer.OrdinalIgnoreCase);
                                int removedCount = reviztoRecords.RemoveAll(r => namesToDelete.Contains(r.LinkName));

                                if (removedCount > 0)
                                {
                                    SaveDataToExtensibleStorage(_doc, reviztoRecords, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                                }
                                tx.Commit();

                                foreach (var p in placeholdersToDelete) Links.Remove(p);
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
                    TaskDialog.Show("Deletion Not Allowed", "Only placeholder rows can be deleted.");
                }
                e.Handled = true;
            }
        }

        #region Context Menu Handlers

        private void LinksDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
            if (cell?.DataContext is LinkViewModel vm && (cell.Column.SortMemberPath == "LinkName" || (cell.Column.Header?.ToString() ?? "") == "File Name"))
            {
                var menu = new ContextMenu();
                menu.Items.Add(new MenuItem { Header = "Unload", Command = new RelayCommand(() => FileName_Unload_Click(vm)), IsEnabled = vm.IsUnloadAvailable });
                menu.Items.Add(new MenuItem { Header = "Reload", Command = new RelayCommand(() => FileName_Reload_Click(vm)), IsEnabled = vm.IsReloadAvailable });
                menu.Items.Add(new MenuItem { Header = "Reload From...", Command = new RelayCommand(() => FileName_ReloadFrom_Click(vm)) });
                menu.Items.Add(new MenuItem { Header = "Open Location", Command = new RelayCommand(() => FileName_OpenLocation_Click(vm)) });
                menu.Items.Add(new MenuItem { Header = "Remove Placeholder", Command = new RelayCommand(() => FileName_RemovePlaceholder_Click(vm)), IsEnabled = vm.IsPlaceholder });
                cell.ContextMenu = menu;
            }
        }

        private void ReloadFrom_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as MenuItem)?.Tag is LinkViewModel vm)
            {
                FileName_ReloadFrom_Click(vm);
            }
        }

        private void LinksDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (FindParent<DataGridRow>(e.OriginalSource as DependencyObject) is DataGridRow row && !row.IsSelected)
            {
                row.IsSelected = true;
            }
        }

        private void CoordinatesCell_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGridCell cell)
            {
                var selectedLinks = LinksDataGrid.SelectedItems.Cast<LinkViewModel>().ToList();
                if (!selectedLinks.Any()) selectedLinks.Add(cell.DataContext as LinkViewModel);

                var menu = new ContextMenu();
                var moveMenuItem = new MenuItem { Header = "Move to Shared Coordinates", IsEnabled = selectedLinks.All(l => l.IsRevitLink && l.LinkStatus == "Loaded") };
                moveMenuItem.Click += (s, args) =>
                {
                    foreach (var link in selectedLinks.Where(l => l.IsRevitLink && l.LinkStatus == "Loaded"))
                    {
                        link.PendingMoveToSharedCoordinates = true;
                    }
                    MessageBox.Show($"{selectedLinks.Count(l => l.PendingMoveToSharedCoordinates)} link(s) marked for move. Click Save to apply.", "Move to Shared Coordinates", MessageBoxButton.OK, MessageBoxImage.Information);
                };
                menu.Items.Add(moveMenuItem);
                cell.ContextMenu = menu;
                menu.IsOpen = true;
                e.Handled = true;
            }
        }

        private void FileName_Unload_Click(LinkViewModel link)
        {
            if (link == null || !link.IsRevitLink || link.LinkStatus != "Loaded") return;
            using (var tx = new Transaction(_doc, $"Unload Link {link.LinkName}"))
            {
                tx.Start();
                try
                {
                    if (_doc.GetElement(link.LinkTypeId) is RevitLinkType linkType)
                    {
                        linkType.Unload(null);
                        tx.Commit();
                        link.LinkStatus = "Unloaded";
                        link.VersionStatus = "N/A (Unloaded)";
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to unload link: {ex.Message}");
                }
            }
        }

        private void FileName_Reload_Click(LinkViewModel link)
        {
            if (link == null || !link.IsRevitLink) return;
            using (var tx = new Transaction(_doc, $"Reload Link {link.LinkName}"))
            {
                tx.Start();
                try
                {
                    if (_doc.GetElement(link.LinkTypeId) is RevitLinkType linkType)
                    {
                        linkType.Reload();
                        tx.Commit();
                        _ = LoadDataAsync();
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to reload link: {ex.Message}");
                }
            }
        }

        private void FileName_ReloadFrom_Click(LinkViewModel link)
        {
            if (link == null) return;
            var dlg = new OpenFileDialog { Title = "Select New Link Source", Filter = "Revit Files (*.rvt)|*.rvt" };
            if (dlg.ShowDialog() == true)
            {
                using (var tx = new Transaction(_doc, $"Reload Link From {link.LinkName}"))
                {
                    tx.Start();
                    try
                    {
                        if (link.IsRevitLink && _doc.GetElement(link.LinkTypeId) is RevitLinkType linkType)
                        {
                            linkType.LoadFrom(ModelPathUtils.ConvertUserVisiblePathToModelPath(dlg.FileName), new WorksetConfiguration());
                        }
                        else
                        {
                            RevitLinkType.Create(_doc, ModelPathUtils.ConvertUserVisiblePathToModelPath(dlg.FileName), new RevitLinkOptions(false));
                        }
                        tx.Commit();
                        _ = LoadDataAsync();
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Error", $"Failed to reload link: {ex.Message}");
                    }
                }
            }
        }

        private void FileName_OpenLocation_Click(LinkViewModel link)
        {
            if (link != null && !string.IsNullOrEmpty(link.FullPath))
            {
                try
                {
                    if (File.Exists(link.FullPath))
                    {
                        Process.Start("explorer.exe", $"/select,\"{link.FullPath}\"");
                    }
                    else
                    {
                        TaskDialog.Show("File Not Found", $"The file path does not exist:\n{link.FullPath}");
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Failed to open file location: {ex.Message}");
                }
            }
        }

        private void FileName_RemovePlaceholder_Click(LinkViewModel link)
        {
            if (link == null || link.IsRevitLink) return;
            Links.Remove(link);
            using (var tx = new Transaction(_doc, "Remove Placeholder"))
            {
                tx.Start();
                try
                {
                    var reviztoRecords = RecallDataFromExtensibleStorage<ReviztoLinkRecord>(_doc, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                    if (reviztoRecords.RemoveAll(r => r.LinkName.Equals(link.LinkName, StringComparison.OrdinalIgnoreCase)) > 0)
                    {
                        SaveDataToExtensibleStorage(_doc, reviztoRecords, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                    }
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to remove placeholder from storage: {ex.Message}");
                }
            }
        }

        #endregion

        private void CloseButton_Click(object sender, RoutedEventArgs e) => this.Close();

        private void ClearProfileButton_Click(object sender, RoutedEventArgs e)
        {
            var result = TaskDialog.Show("Clear All Saved Data", "Are you sure you want to clear all saved link data?", TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
            if (result == TaskDialogResult.Yes)
            {
                using (var tx = new Transaction(_doc, "Clear All Link Manager Data"))
                {
                    tx.Start();
                    try
                    {
                        var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
                        foreach (var ds in collector)
                        {
                            if (ds.Name == ProfileDataStorageElementName || ds.Name == ReviztoLinkRecord.DataStorageName)
                            {
                                _doc.Delete(ds.Id);
                            }
                        }
                        tx.Commit();
                        _ = LoadDataAsync();
                        TaskDialog.Show("All Data Cleared", "The saved link profile and Revizto data have been cleared.");
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        TaskDialog.Show("Error", $"Failed to clear data: {ex.Message}");
                    }
                }
            }
        }

        public static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            return parentObject is T parent ? parent : FindParent<T>(parentObject);
        }

        private void LinksDataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            string sortMemberPath = e.Column.SortMemberPath;
            if (string.IsNullOrEmpty(sortMemberPath)) return;

            ListSortDirection direction = (e.Column.SortDirection == ListSortDirection.Ascending) ? ListSortDirection.Descending : ListSortDirection.Ascending;
            e.Column.SortDirection = direction;

            _linksView.SortDescriptions.Clear();
            _linksView.SortDescriptions.Add(new SortDescription(sortMemberPath, direction));
        }

        private void ImportReviztoCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv", Title = "Select Revizto Link Export CSV" };
            if (openFileDialog.ShowDialog() != true) return;

            try
            {
                var allReviztoLinks = new List<ReviztoLinkRecord>();
                using (TextFieldParser parser = new TextFieldParser(openFileDialog.FileName))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    if (parser.EndOfData) { TaskDialog.Show("Import Error", "CSV file is empty."); return; }

                    string[] headers = parser.ReadFields();
                    int linkNameIndex = Array.IndexOf(headers, "Original name");
                    int lastModifiedIndex = Array.IndexOf(headers, "Last exported");
                    int descriptionIndex = Array.IndexOf(headers, "Model");

                    if (linkNameIndex == -1 || lastModifiedIndex == -1 || descriptionIndex == -1)
                    {
                        TaskDialog.Show("Import Error", "CSV is missing required columns: 'Original name', 'Last exported', 'Model'.");
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
                                LastModified = fields[lastModifiedIndex]
                            };
                            if (!string.IsNullOrWhiteSpace(record.LinkName)) allReviztoLinks.Add(record);
                        }
                        catch (MalformedLineException) { /* Skip malformed lines */ }
                    }
                }

                var uniqueReviztoLinks = allReviztoLinks
                    .GroupBy(r => r.LinkName)
                    .Select(g => g.OrderByDescending(r => { DateTime.TryParse(r.LastModified, out DateTime dt); return dt; }).First())
                    .ToList();

                using (var tx = new Transaction(_doc, "Import Revizto Link Data"))
                {
                    tx.Start();
                    var existing = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage)).FirstOrDefault(ds => ds.Name == ReviztoLinkRecord.DataStorageName);
                    if (existing != null) _doc.Delete(existing.Id);
                    SaveDataToExtensibleStorage(_doc, uniqueReviztoLinks, ReviztoLinkRecord.SchemaGuid, ReviztoLinkRecord.SchemaName, ReviztoLinkRecord.FieldName, ReviztoLinkRecord.DataStorageName);
                    tx.Commit();
                }
                TaskDialog.Show("Import Complete", $"{uniqueReviztoLinks.Count} unique Revizto links imported.");
                _ = LoadDataAsync();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Import Error", $"An unexpected error occurred: {ex.Message}");
            }
        }

        private void ExportToCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "CSV Files (*.csv)|*.csv", FileName = "RevitLinkManagerExport.csv", Title = "Save Link Data to CSV" };
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    var sb = new StringBuilder();
                    var headers = LinksDataGrid.Columns.Where(c => c.Visibility == System.Windows.Visibility.Visible).Select(c => EscapeCsvField(c.Header.ToString().Replace("&#x0a;", " ")));
                    sb.AppendLine(string.Join(",", headers));

                    foreach (LinkViewModel link in LinksDataGrid.Items)
                    {
                        var fields = LinksDataGrid.Columns.Where(c => c.Visibility == System.Windows.Visibility.Visible).Select(column =>
                        {
                            string propertyName = string.Empty;
                            if (column is DataGridBoundColumn boundColumn && boundColumn.Binding is System.Windows.Data.Binding wpfBinding)
                            {
                                propertyName = wpfBinding.Path.Path;
                            }
                            else
                            {
                                propertyName = column.SortMemberPath;
                            }

                            if (!string.IsNullOrEmpty(propertyName))
                            {
                                var prop = typeof(LinkViewModel).GetProperty(propertyName);
                                return EscapeCsvField(prop?.GetValue(link)?.ToString() ?? "");
                            }
                            return "";
                        });
                        sb.AppendLine(string.Join(",", fields));
                    }
                    File.WriteAllText(saveFileDialog.FileName, sb.ToString());
                    TaskDialog.Show("Export Complete", "Link data exported to CSV successfully.");
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Export Error", $"Failed to export data to CSV: {ex.Message}");
                }
            }
        }

        private void ImportProfileCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv", Title = "Select Profile CSV to Import" };
            if (openFileDialog.ShowDialog() != true) return;

            try
            {
                using (var parser = new TextFieldParser(openFileDialog.FileName))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    if (parser.EndOfData) { TaskDialog.Show("Import Error", "CSV file is empty."); return; }

                    var headers = parser.ReadFields().Select(h => h.Replace(" ", "").Replace("&#x0a;", "")).ToArray();
                    var headerMap = new Dictionary<string, int>();
                    for (int i = 0; i < headers.Length; i++) headerMap[headers[i]] = i;

                    int updatedCount = 0, createdCount = 0;
                    while (!parser.EndOfData)
                    {
                        var fields = parser.ReadFields();
                        string linkName = fields[headerMap["LinkFileName"]];
                        var existingLink = Links.FirstOrDefault(l => l.LinkName.Equals(linkName, StringComparison.OrdinalIgnoreCase));

                        if (existingLink != null)
                        {
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
                    TaskDialog.Show("Import Complete", $"Profile data imported.\n{updatedCount} links updated.\n{createdCount} new placeholders created.");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Import Error", $"An error occurred while importing: {ex.Message}");
            }
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private void PinCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as CheckBox)?.DataContext is LinkViewModel vm && vm.IsRevitLink)
            {
                vm.IsPinned = vm.IsPinned != true;
            }
        }

        private static string GetCellValue(IList<string> cellValues, Dictionary<string, int> headerMap, string header) { if (headerMap.TryGetValue(header, out int idx) && idx < cellValues.Count) return cellValues[idx]; return string.Empty; }
        public class ReviztoLinkRecord { public static readonly Guid SchemaGuid = new Guid("A1B2C3D4-E5F6-47A8-9B0C-1234567890AB"); public const string SchemaName = "RTS_ReviztoLinkSchema"; public const string FieldName = "ReviztoLinkJson"; public const string DataStorageName = "RTS_Revizto_Link_Storage"; public string LinkName { get; set; } public string FilePath { get; set; } public string Description { get; set; } public string LastModified { get; set; } }
        private Schema GetOrCreateSchema(Guid schemaGuid, string schemaName, string fieldName) { Schema schema = Schema.Lookup(schemaGuid); if (schema == null) { SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid); schemaBuilder.SetSchemaName(schemaName); schemaBuilder.SetReadAccessLevel(AccessLevel.Public); schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor); schemaBuilder.SetVendorId(VendorId); schemaBuilder.AddSimpleField(fieldName, typeof(string)); schema = schemaBuilder.Finish(); } return schema; }
        private DataStorage GetOrCreateDataStorage(Document doc, string dataStorageElementName) { var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)); DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName); if (dataStorage == null) { dataStorage = DataStorage.Create(doc); dataStorage.Name = dataStorageElementName; } return dataStorage; }
        public void SaveDataToExtensibleStorage<T>(Document doc, List<T> dataList, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) { Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName); DataStorage dataStorage = GetOrCreateDataStorage(doc, dataStorageElementName); string jsonString = JsonSerializer.Serialize(dataList, new JsonSerializerOptions { WriteIndented = true }); Entity entity = new Entity(schema); entity.Set(schema.GetField(fieldName), jsonString); dataStorage.SetEntity(entity); }
        public List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) { Schema schema = Schema.Lookup(schemaGuid); if (schema == null) return new List<T>(); var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)); DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName); if (dataStorage == null) return new List<T>(); Entity entity = dataStorage.GetEntity(schema); if (!entity.IsValid()) return new List<T>(); string jsonString = entity.Get<string>(schema.GetField(fieldName)); if (string.IsNullOrEmpty(jsonString)) return new List<T>(); try { return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>(); } catch (Exception ex) { TaskDialog.Show("Profile Error", $"Failed to read saved profile: {ex.Message}"); return new List<T>(); } }

        private Autodesk.Revit.DB.Transform GetSharedCoordinatesTransform()
        {
            var settings = RTS_RevitUtils.GetProfileSettings(_doc);
            if (settings == null) return null;

            var sharedCoordsMapping = settings.CoordinateSystemMappings.FirstOrDefault(m => m.SystemName == "Shared Coordinates Source");
            if (sharedCoordsMapping == null || string.IsNullOrWhiteSpace(sharedCoordsMapping.SelectedLink) || sharedCoordsMapping.SelectedLink == "<None>") return null;

            string linkName = RTS_RevitUtils.ParseLinkName(sharedCoordsMapping.SelectedLink);

            var linkVm = Links.FirstOrDefault(l => l.LinkName.Equals(linkName, StringComparison.OrdinalIgnoreCase) && l.IsRevitLink && l.LinkStatus == "Loaded" && l.InstanceTransforms.Any());
            return linkVm?.InstanceTransforms.First();
        }
    }

    /// <summary>
    /// ViewModel for a single row in the DataGrid.
    /// </summary>
    public class LinkViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        // Revit-derived properties (not saved)
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
            set { _reviztoStatus = value; OnPropertyChanged(nameof(ReviztoStatus)); OnPropertyChanged(nameof(IsAlternativeMatch)); }
        }
        [JsonIgnore] public bool IsAlternativeMatch => !string.IsNullOrEmpty(ReviztoStatus) && ReviztoStatus.StartsWith("Match on:");
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
        [JsonIgnore] public bool? IsPinned { get => _isPinned; set { _isPinned = value; OnPropertyChanged(nameof(IsPinned)); } }
        private string _linkCoordinates;
        [JsonIgnore] public string LinkCoordinates { get => _linkCoordinates; set { _linkCoordinates = value; OnPropertyChanged(nameof(LinkCoordinates)); } }
        [JsonIgnore] public List<Autodesk.Revit.DB.Transform> InstanceTransforms { get; set; } = new List<Autodesk.Revit.DB.Transform>();

        // User-editable properties (saved)
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

        [JsonIgnore] public bool PendingMoveToSharedCoordinates { get; set; }
        [JsonIgnore] public bool IsUnloadAvailable => IsRevitLink && LinkStatus == "Loaded";
        [JsonIgnore] public bool IsReloadAvailable => IsRevitLink;
        [JsonIgnore] public bool IsPlaceholder => !IsRevitLink;

        /// <summary>
        /// Applies saved data from another LinkViewModel.
        /// </summary>
        public void ApplyProfileData(LinkViewModel source)
        {
            this.LinkDescription = source.LinkDescription;
            this.SelectedDiscipline = source.SelectedDiscipline;
            this.CompanyName = source.CompanyName;
            this.ResponsiblePerson = source.ResponsiblePerson;
            this.ContactDetails = source.ContactDetails;
            this.Comments = source.Comments;
        }
    }

    public class ProjectRoleContact : INotifyPropertyChanged
    {
        public string Role { get; set; }
        private string _name;
        public string Name { get => _name; set { _name = value; OnPropertyChanged(nameof(Name)); } }
        private string _email;
        public string Email { get => _email; set { _email = value; OnPropertyChanged(nameof(Email)); } }
        private string _phone;
        public string Phone { get => _phone; set { _phone = value; OnPropertyChanged(nameof(Phone)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class ReviztoStatusToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var status = value as string;
            if (status == "Matched" || (status != null && status.StartsWith("Match on:"))) return Brushes.Green;
            return Brushes.Red;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class ActiveModelStatusToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is LinkViewModel vm)
            {
                if (vm.LinkStatus == "Loaded") return Brushes.Green;
                if (vm.IsRevitLink) return Brushes.Orange;
            }
            return Brushes.Red;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    /// <summary>
    /// A simple ICommand implementation for use in XAML.
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        public event EventHandler CanExecuteChanged { add { } remove { } }
        public RelayCommand(Action execute) => _execute = execute;
        public bool CanExecute(object parameter) => true;
        public void Execute(object parameter) => _execute();
    }
}