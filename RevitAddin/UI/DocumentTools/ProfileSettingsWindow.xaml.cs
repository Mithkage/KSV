using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Text.RegularExpressions;
using System.Windows.Media;

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for ProfileSettingsWindow.xaml.
    /// This window allows users to map specific model categories (Grids, Walls, etc.)
    /// to a source Revit link from the main Link Hub.
    /// </summary>
    public partial class ProfileSettingsWindow : Window
    {
        private Document _doc;

        // Extensible Storage definitions for the settings profile.
        public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
        public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
        public const string SettingsFieldName = "ProfileSettingsJson";
        public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";
        public const string VendorId = "ReTick_Solutions";

        public ProfileSettings Settings { get; set; }
        public ObservableCollection<string> AvailableLinks { get; set; }
        public ObservableCollection<ModelCategoryMapping> ModelCategoryMappings { get; set; }
        public ObservableCollection<ModelCategoryMapping> AnnotationCategoryMappings { get; set; }
        public ObservableCollection<CoordinateSystemMapping> CoordinateSystemMappings { get; set; }
        public ICollectionView FilteredModelCategoryMappings { get; private set; }

        private Popup _modelCategoryFilterPopup;
        private ICollectionView _filteredModelCategoryMappings;

        private Popup _columnFilterPopup;
        private Dictionary<string, Predicate<object>> _activeFilters = new Dictionary<string, Predicate<object>>();

        /// <summary>
        /// Constructor for the Profile Settings window.
        /// </summary>
        public ProfileSettingsWindow(Document doc, IEnumerable<LinkViewModel> links)
        {
            InitializeComponent();
            _doc = doc;
            this.DataContext = this;

            // Populate the dropdown source list from the links in the main window, sorted alphabetically.
            AvailableLinks = new ObservableCollection<string> { "<None>" };
            foreach (var link in links
                .Where(l => l.IsRevitLink)
                .OrderBy(l => l.SelectedDiscipline)
                .ThenBy(l => l.LinkName))
            {
                AvailableLinks.Add($"[{link.SelectedDiscipline}] - {link.LinkName}");
            }

            // Helper to format category names
            string FormatCategoryName(string raw)
            {
                if (raw.StartsWith("OST_"))
                    raw = raw.Substring(4);
                return Regex.Replace(raw, "(?<!^)([A-Z])", " $1");
            }

            // Dynamically populate all model categories from BuiltInCategory
            ModelCategoryMappings = new ObservableCollection<ModelCategoryMapping>(
                Enum.GetValues(typeof(BuiltInCategory))
                    .Cast<BuiltInCategory>()
                    .Where(cat => IsModelCategory(cat))
                    .Select(cat => new ModelCategoryMapping
                    {
                        CategoryName = FormatCategoryName(cat.ToString()),
                        SelectedLink = "<None>"
                    })
            );

            // Dynamically populate all annotation categories from BuiltInCategory
            AnnotationCategoryMappings = new ObservableCollection<ModelCategoryMapping>(
                Enum.GetValues(typeof(BuiltInCategory))
                    .Cast<BuiltInCategory>()
                    .Where(cat => IsAnnotationCategory(cat))
                    .Select(cat => new ModelCategoryMapping
                    {
                        CategoryName = FormatCategoryName(cat.ToString()),
                        SelectedLink = "<None>"
                    })
            );

            // Restore Coordinate System Mappings
            CoordinateSystemMappings = new ObservableCollection<CoordinateSystemMapping>
            {
                new CoordinateSystemMapping { SystemName = "Project Base Point", SelectedLink = "<None>" },
                new CoordinateSystemMapping { SystemName = "Survey Point", SelectedLink = "<None>" },
                new CoordinateSystemMapping { SystemName = "Internal Origin", SelectedLink = "<None>" }
            };

            // Setup filtered view for Model Categories
            _filteredModelCategoryMappings = CollectionViewSource.GetDefaultView(ModelCategoryMappings);
            _filteredModelCategoryMappings.Filter = null; // No filter by default

            FilteredModelCategoryMappings = CollectionViewSource.GetDefaultView(ModelCategoryMappings);
            FilteredModelCategoryMappings.Filter = null;

            LoadSettings();
        }

        /// <summary>
        /// Loads saved settings from Extensible Storage.
        /// </summary>
        private void LoadSettings()
        {
            var recalledSettings = RecallDataFromExtensibleStorage<ProfileSettings>(_doc, SettingsSchemaGuid, SettingsSchemaName, SettingsFieldName, SettingsDataStorageElementName);
            Settings = recalledSettings.FirstOrDefault() ?? new ProfileSettings();
        }

        /// <summary>
        /// Handles the click event for the "Save" button.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var tx = new Transaction(_doc, "Save Profile Settings"))
                {
                    tx.Start();
                    SaveDataToExtensibleStorage(_doc, new List<ProfileSettings> { Settings }, SettingsSchemaGuid, SettingsSchemaName, SettingsFieldName, SettingsDataStorageElementName);
                    tx.Commit();
                }
                TaskDialog.Show("Success", "Profile settings saved successfully.");
                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to save profile settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Handles the click event for the "Clear Profile" button.
        /// </summary>
        private void ClearProfileButton_Click(object sender, RoutedEventArgs e)
        {
            var result = Autodesk.Revit.UI.TaskDialog.Show(
                "Clear Profile",
                "Are you sure you want to clear all saved profile data for this project? This action cannot be undone.",
                Autodesk.Revit.UI.TaskDialogCommonButtons.Yes | Autodesk.Revit.UI.TaskDialogCommonButtons.No,
                Autodesk.Revit.UI.TaskDialogResult.No);

            if (result == Autodesk.Revit.UI.TaskDialogResult.Yes)
            {
                using (var tx = new Autodesk.Revit.DB.Transaction(_doc, "Clear Profile Data"))
                {
                    tx.Start();
                    try
                    {
                        var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
                        var allDataStorage = collector.ToElements();

                        // Find and delete the Link Manager Profile DataStorage element
                        var profileDataStorage = allDataStorage.FirstOrDefault(ds => ds.Name == LinkManagerWindow.ProfileDataStorageElementName);
                        if (profileDataStorage != null)
                        {
                            _doc.Delete(profileDataStorage.Id);
                        }

                        tx.Commit();
                        Autodesk.Revit.UI.TaskDialog.Show("Profile Cleared", "The saved profile data has been cleared.", Autodesk.Revit.UI.TaskDialogCommonButtons.Ok);
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        Autodesk.Revit.UI.TaskDialog.Show("Error", $"Failed to clear profile data: {ex.Message}", Autodesk.Revit.UI.TaskDialogCommonButtons.Ok);
                    }
                }
            }
        }

        #region Extensible Storage Helpers
        public void SaveDataToExtensibleStorage<T>(Document doc, List<T> dataList, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName)
        {
            Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName);
            DataStorage dataStorage = GetOrCreateDataStorage(doc, dataStorageElementName);
            string jsonString = JsonSerializer.Serialize(dataList);
            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(fieldName), jsonString);
            dataStorage.SetEntity(entity);
        }

        public List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName) where T : new()
        {
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema == null) return new List<T>();
            var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            if (dataStorage == null) return new List<T>();
            Entity entity = dataStorage.GetEntity(schema);
            if (!entity.IsValid()) return new List<T>();
            string jsonString = entity.Get<string>(schema.GetField(fieldName));
            if (string.IsNullOrEmpty(jsonString)) return new List<T>();
            try
            {
                return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>();
            }
            catch
            {
                return new List<T>();
            }
        }

        private Schema GetOrCreateSchema(Guid schemaGuid, string schemaName, string fieldName)
        {
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid);
                schemaBuilder.SetSchemaName(schemaName);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor);
                schemaBuilder.SetVendorId(VendorId);
                schemaBuilder.AddSimpleField(fieldName, typeof(string));
                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        private DataStorage GetOrCreateDataStorage(Document doc, string dataStorageElementName)
        {
            var collector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            DataStorage dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            if (dataStorage == null)
            {
                dataStorage = DataStorage.Create(doc);
                dataStorage.Name = dataStorageElementName;
            }
            return dataStorage;
        }
        #endregion

        /// <summary>
        /// Gets the shared coordinates transform for the selected link.
        /// </summary>
        public Autodesk.Revit.DB.Transform GetSharedCoordinatesTransform(IEnumerable<LinkViewModel> links)
        {
            string selected = Settings.SharedCoordinatesLink;
            if (string.IsNullOrWhiteSpace(selected) || selected == "<None>")
                return null;

            string linkName = selected;
            int idx = linkName.IndexOf("] - ");
            if (idx > 0)
                linkName = linkName.Substring(idx + 4);

            var linkVm = links.FirstOrDefault(l => l.LinkName.Equals(linkName, StringComparison.OrdinalIgnoreCase) && l.IsRevitLink && l.LinkStatus == "Loaded");
            if (linkVm != null && linkVm.InstanceTransforms != null && linkVm.InstanceTransforms.Any())
            {
                return linkVm.InstanceTransforms.First();
            }
            return null;
        }

        private bool IsModelCategory(BuiltInCategory cat)
        {
            if (cat == BuiltInCategory.INVALID) return false;
            string name = cat.ToString();
            string[] excludeKeywords = { "ANNOTATION", "TAG", "TEXT", "REVISION", "KEYNOTE", "LINE", "VIEW", "SHEET", "TITLEBLOCK", "SCHEDULE", "LEGEND", "GRID", "LEVEL", "REFERENCE", "INTERNAL", "CURTAIN", "ROOM", "SPACE", "AREA", "ZONE", "WARNING" };
            return !excludeKeywords.Any(kw => name.ToUpperInvariant().Contains(kw));
        }

        private bool IsAnnotationCategory(BuiltInCategory cat)
        {
            if (cat == BuiltInCategory.INVALID) return false;
            string name = cat.ToString();
            string[] includeKeywords = {
                "ANNOTATION", "TAG", "TEXT", "REVISION", "KEYNOTE", "DIMENSION", "LINE",
                "VIEW", "SHEET", "TITLEBLOCK", "SCHEDULE", "LEGEND", "GRID", "LEVEL",
                "REFERENCE", "SYMBOL", "MARKER", "BUBBLE", "CLOUD", "TITLE", "CALLOUT"
            };
            return includeKeywords.Any(kw => name.ToUpperInvariant().Contains(kw));
        }

        // Dropdown filter for Model Category column header
        private void ModelCategoryHeaderFilter_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            if (_modelCategoryFilterPopup == null)
            {
                _modelCategoryFilterPopup = new Popup
                {
                    PlacementTarget = button,
                    Placement = PlacementMode.Bottom,
                    StaysOpen = false,
                    AllowsTransparency = true,
                    Width = 220,
                    Height = 80
                };

                var panel = new StackPanel { Background = System.Windows.Media.Brushes.White, Margin = new Thickness(5) };
                var textBox = new System.Windows.Controls.TextBox { Width = 200, Margin = new Thickness(0, 0, 0, 5) };
                var applyButton = new Button { Content = "Apply Filter", Width = 90, Margin = new Thickness(0, 0, 0, 0) };

                applyButton.Click += (s, args) =>
                {
                    string pattern = textBox.Text.Trim();
                    if (string.IsNullOrWhiteSpace(pattern))
                    {
                        _filteredModelCategoryMappings.Filter = null;
                    }
                    else
                    {
                        _filteredModelCategoryMappings.Filter = item =>
                        {
                            if (item is ModelCategoryMapping mapping)
                            {
                                string regexPattern = "^" + Regex.Escape(pattern)
                                    .Replace("\\*", ".*")
                                    .Replace("\\?", ".") + "$";
                                return Regex.IsMatch(mapping.CategoryName, regexPattern, RegexOptions.IgnoreCase);
                            }
                            return true;
                        };
                    }
                    _filteredModelCategoryMappings.Refresh();
                    _modelCategoryFilterPopup.IsOpen = false;
                };

                panel.Children.Add(new TextBlock { Text = "Filter (wildcards * ?)", Foreground = System.Windows.Media.Brushes.Gray });
                panel.Children.Add(textBox);
                panel.Children.Add(applyButton);

                _modelCategoryFilterPopup.Child = panel;
            }

            _modelCategoryFilterPopup.PlacementTarget = button;
            _modelCategoryFilterPopup.IsOpen = true;
        }

        private void ColumnHeaderFilter_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            var header = FindParent<DataGridColumnHeader>(button);
            if (header == null || header.Column == null) return;

            string sortMemberPath = header.Column.SortMemberPath;
            if (string.IsNullOrEmpty(sortMemberPath))
            {
                // For template columns, use header text as key
                sortMemberPath = header.Column.Header?.ToString();
            }

            // Get the DataGrid and its items
            var dataGrid = FindParent<DataGrid>(header);
            if (dataGrid == null) return;

            // Create the popup
            if (_columnFilterPopup != null)
                _columnFilterPopup.IsOpen = false;

            _columnFilterPopup = new Popup
            {
                PlacementTarget = button,
                Placement = PlacementMode.Bottom,
                StaysOpen = false,
                AllowsTransparency = true,
                Width = 250,
                Height = 90
            };

            var stackPanel = new StackPanel { Background = System.Windows.Media.Brushes.White, Margin = new Thickness(5) };
            var searchBox = new System.Windows.Controls.TextBox { Margin = new Thickness(0, 0, 0, 5), Width = 220 };

            // Filter on Enter key
            searchBox.KeyDown += (s, args) =>
            {
                if (args.Key == System.Windows.Input.Key.Enter)
                {
                    string searchText = searchBox.Text.Trim();
                    if (string.IsNullOrEmpty(searchText))
                    {
                        _activeFilters.Remove(sortMemberPath);
                    }
                    else
                    {
                        _activeFilters[sortMemberPath] = item =>
                        {
                            var prop = item.GetType().GetProperty(sortMemberPath);
                            var value = prop?.GetValue(item)?.ToString() ?? "";
                            return value.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
                        };
                    }
                    ApplyFilters(dataGrid, sortMemberPath);
                    header.Background = System.Windows.Media.Brushes.LightBlue;
                    _columnFilterPopup.IsOpen = false;
                }
            };

            var clearButton = new Button
            {
                Content = "Clear Filter",
                Margin = new Thickness(0, 5, 0, 0),
                Background = System.Windows.Media.Brushes.LightGray,
                Width = 100
            };
            clearButton.Click += (s, args) =>
            {
                _activeFilters.Remove(sortMemberPath);
                ApplyFilters(dataGrid, sortMemberPath);
                header.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                _columnFilterPopup.IsOpen = false;
            };

            stackPanel.Children.Add(new TextBlock { Text = "Search (press Enter)", Foreground = System.Windows.Media.Brushes.Gray });
            stackPanel.Children.Add(searchBox);
            stackPanel.Children.Add(clearButton);

            _columnFilterPopup.Child = stackPanel;
            _columnFilterPopup.IsOpen = true;
        }

        // Helper to apply all active filters to a DataGrid's ICollectionView
        private void ApplyFilters(DataGrid dataGrid, string sortMemberPath)
        {
            var view = CollectionViewSource.GetDefaultView(dataGrid.ItemsSource);
            if (view == null) return;

            view.Filter = item =>
            {
                foreach (var filter in _activeFilters)
                {
                    if (!filter.Value(item))
                        return false;
                }
                return true;
            };
            view.Refresh();
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
    }

    public class ProfileSettings : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private string _gridsLink;
        public string GridsLink { get => _gridsLink; set { _gridsLink = value; OnPropertyChanged(nameof(GridsLink)); } }
        private string _levelsLink;
        public string LevelsLink { get => _levelsLink; set { _levelsLink = value; OnPropertyChanged(nameof(LevelsLink)); } }
        private string _wallsLink;
        public string WallsLink { get => _wallsLink; set { _wallsLink = value; OnPropertyChanged(nameof(WallsLink)); } }
        private string _floorsLink;
        public string FloorsLink { get => _floorsLink; set { _floorsLink = value; OnPropertyChanged(nameof(FloorsLink)); } }
        private string _ceilingsLink;
        public string CeilingsLink { get => _ceilingsLink; set { _ceilingsLink = value; OnPropertyChanged(nameof(CeilingsLink)); } }
        private string _slabsLink;
        public string SlabsLink { get => _slabsLink; set { _slabsLink = value; OnPropertyChanged(nameof(SlabsLink)); } }
        private string _sharedCoordinatesLink;
        public string SharedCoordinatesLink
        {
            get => _sharedCoordinatesLink;
            set { _sharedCoordinatesLink = value; OnPropertyChanged(nameof(SharedCoordinatesLink)); }
        }
    }

    public class ModelCategoryMapping : INotifyPropertyChanged
    {
        public string CategoryName { get; set; }
        private string _selectedLink;
        public string SelectedLink
        {
            get => _selectedLink;
            set { _selectedLink = value; OnPropertyChanged(nameof(SelectedLink)); }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class CoordinateSystemMapping
    {
        public string SystemName { get; set; }
        public string SelectedLink { get; set; }
    }
}