// --- FILE: ProfileSettingsWindow.xaml.cs (UPDATED) ---
//
// File: ProfileSettingsWindow.xaml.cs
// Namespace: RTS.UI
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// This file contains the code-behind for the Profile Settings window. It allows
// users to map Revit categories to specific source links, now using curated lists
// and discipline-based grouping for a more intuitive workflow.
//
// Log:
// - 2025-08-11: Temporarily commented out version-specific categories to resolve compilation errors.
// - 2025-08-11: Refactored CategoryData to use a constructor, correctly resolving version-specific compilation errors.
// - 2025-08-11: Corrected BuiltInCategory names for Revit 2022/2024 API compatibility.
// - 2025-08-11: Added window icon and enabled resizing.
// - 2025-08-11: Replaced dynamic category generation with curated, grouped lists.
// - 2025-08-11: Implemented group-based link assignment with individual overrides.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;
#endregion

namespace RTS.UI
{
    public partial class ProfileSettingsWindow : Window
    {
        private Document _doc;
        private CategoryData _categoryData;

        // Extensible Storage definitions
        public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
        public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
        public const string SettingsFieldName = "ProfileSettingsJson";
        public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";
        public const string VendorId = "ReTick_Solutions";

        // ViewModels and Data Sources
        public ProfileSettings Settings { get; set; }
        public ObservableCollection<string> AvailableLinks { get; set; }
        public ICollectionView ModelCategoryMappingsView { get; set; }
        public ICollectionView AnnotationCategoryMappingsView { get; set; }

        private ObservableCollection<CategoryMappingViewModel> _modelCategoryMappings;
        private ObservableCollection<CategoryMappingViewModel> _annotationCategoryMappings;

        public ProfileSettingsWindow(Document doc, IEnumerable<LinkViewModel> links)
        {
            InitializeComponent();
            LoadIcon(); // Load the window icon
            _doc = doc;
            _categoryData = new CategoryData(); // Instantiate category data
            this.DataContext = this;

            // Populate available links dropdown
            AvailableLinks = new ObservableCollection<string> { "<None>" };
            foreach (var link in links.Where(l => l.IsRevitLink).OrderBy(l => l.SelectedDiscipline).ThenBy(l => l.LinkName))
            {
                AvailableLinks.Add($"[{link.SelectedDiscipline}] - {link.LinkName}");
            }

            // Load saved settings or create new ones
            LoadSettings();

            // Initialize and populate the category lists
            InitializeCategoryMappings();

            // Apply saved settings to the view models
            ApplyLoadedMappings();

            // Set up the grouped views for the DataGrids
            ModelCategoryMappingsView = CollectionViewSource.GetDefaultView(_modelCategoryMappings);
            ModelCategoryMappingsView.GroupDescriptions.Add(new PropertyGroupDescription("Discipline"));

            AnnotationCategoryMappingsView = CollectionViewSource.GetDefaultView(_annotationCategoryMappings);
            AnnotationCategoryMappingsView.GroupDescriptions.Add(new PropertyGroupDescription("Discipline"));
        }

        /// <summary>
        /// Loads the window icon using a robust Pack URI.
        /// </summary>
        private void LoadIcon()
        {
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.ico", UriKind.Absolute);
                this.Icon = new BitmapImage(iconUri);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning: Could not load window icon. {ex.Message}");
            }
        }

        /// <summary>
        /// Populates the mapping collections from the curated static lists.
        /// </summary>
        private void InitializeCategoryMappings()
        {
            _modelCategoryMappings = new ObservableCollection<CategoryMappingViewModel>();
            foreach (var group in _categoryData.ModelCategories.GroupBy(c => c.Discipline))
            {
                var groupVms = group.Select(c => new CategoryMappingViewModel(this)
                {
                    Discipline = c.Discipline,
                    CategoryName = c.CategoryName,
                    BuiltInCategory = c.Bic
                }).ToList();

                foreach (var vm in groupVms)
                {
                    vm.SetGroup(groupVms);
                    _modelCategoryMappings.Add(vm);
                }
            }

            _annotationCategoryMappings = new ObservableCollection<CategoryMappingViewModel>();
            foreach (var group in _categoryData.AnnotationCategories.GroupBy(c => c.Discipline))
            {
                var groupVms = group.Select(c => new CategoryMappingViewModel(this)
                {
                    Discipline = c.Discipline,
                    CategoryName = c.CategoryName,
                    BuiltInCategory = c.Bic
                }).ToList();

                foreach (var vm in groupVms)
                {
                    vm.SetGroup(groupVms);
                    _annotationCategoryMappings.Add(vm);
                }
            }
        }

        /// <summary>
        /// Applies settings that were loaded from extensible storage to the view models.
        /// </summary>
        private void ApplyLoadedMappings()
        {
            foreach (var savedMapping in Settings.ModelCategoryMappings)
            {
                var targetVm = _modelCategoryMappings.FirstOrDefault(vm => vm.CategoryName == savedMapping.CategoryName);
                if (targetVm != null)
                {
                    targetVm.SelectedLink = savedMapping.SelectedLink;
                }
            }

            foreach (var savedMapping in Settings.AnnotationCategoryMappings)
            {
                var targetVm = _annotationCategoryMappings.FirstOrDefault(vm => vm.CategoryName == savedMapping.CategoryName);
                if (targetVm != null)
                {
                    targetVm.SelectedLink = savedMapping.SelectedLink;
                }
            }
        }

        /// <summary>
        /// Loads settings from Extensible Storage or initializes a new object.
        /// </summary>
        private void LoadSettings()
        {
            var schema = Schema.Lookup(SettingsSchemaGuid);
            if (schema == null)
            {
                Settings = new ProfileSettings();
                return;
            }

            var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
            var dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == SettingsDataStorageElementName);
            if (dataStorage == null)
            {
                Settings = new ProfileSettings();
                return;
            }

            var entity = dataStorage.GetEntity(schema);
            if (!entity.IsValid())
            {
                Settings = new ProfileSettings();
                return;
            }

            string json = entity.Get<string>(schema.GetField(SettingsFieldName));
            if (string.IsNullOrEmpty(json))
            {
                Settings = new ProfileSettings();
                return;
            }

            try
            {
                Settings = JsonSerializer.Deserialize<ProfileSettings>(json) ?? new ProfileSettings();
            }
            catch
            {
                Settings = new ProfileSettings();
            }

            // Ensure coordinate system mappings exist for backward compatibility
            if (Settings.CoordinateSystemMappings == null || !Settings.CoordinateSystemMappings.Any())
            {
                Settings.CoordinateSystemMappings = new List<CoordinateSystemMapping>
                {
                    new CoordinateSystemMapping { SystemName = "Shared Coordinates Source", SelectedLink = "<None>" }
                };
            }
        }

        /// <summary>
        /// Saves the current settings to Extensible Storage.
        /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Update the settings object with the latest selections from the UI
            Settings.ModelCategoryMappings = _modelCategoryMappings.Select(vm => vm.ToDataModel()).ToList();
            Settings.AnnotationCategoryMappings = _annotationCategoryMappings.Select(vm => vm.ToDataModel()).ToList();

            using (var tx = new Transaction(_doc, "Save Profile Settings"))
            {
                tx.Start();
                try
                {
                    var schema = GetOrCreateSchema();
                    var dataStorage = GetOrCreateDataStorage();
                    var entity = new Entity(schema);

                    string json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
                    entity.Set(schema.GetField(SettingsFieldName), json);
                    dataStorage.SetEntity(entity);

                    tx.Commit();
                    TaskDialog.Show("Success", "Profile settings saved successfully.");
                    this.DialogResult = true;
                    this.Close();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    TaskDialog.Show("Error", $"Failed to save profile settings: {ex.Message}");
                }
            }
        }

        #region Extensible Storage Helpers
        private Schema GetOrCreateSchema()
        {
            Schema schema = Schema.Lookup(SettingsSchemaGuid);
            if (schema == null)
            {
                var builder = new SchemaBuilder(SettingsSchemaGuid);
                builder.SetSchemaName(SettingsSchemaName);
                builder.SetVendorId(VendorId);
                builder.SetReadAccessLevel(AccessLevel.Public);
                builder.SetWriteAccessLevel(AccessLevel.Vendor);
                builder.AddSimpleField(SettingsFieldName, typeof(string));
                schema = builder.Finish();
            }
            return schema;
        }

        private DataStorage GetOrCreateDataStorage()
        {
            var collector = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
            var dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == SettingsDataStorageElementName);
            if (dataStorage == null)
            {
                dataStorage = DataStorage.Create(_doc);
                dataStorage.Name = SettingsDataStorageElementName;
            }
            return dataStorage;
        }
        #endregion
    }

    #region Data Models and ViewModels

    /// <summary>
    /// Data model for storing the settings in JSON format.
    /// </summary>
    public class ProfileSettings
    {
        public List<CategoryMapping> ModelCategoryMappings { get; set; } = new List<CategoryMapping>();
        public List<CategoryMapping> AnnotationCategoryMappings { get; set; } = new List<CategoryMapping>();
        public List<CoordinateSystemMapping> CoordinateSystemMappings { get; set; } = new List<CoordinateSystemMapping>
        {
            new CoordinateSystemMapping { SystemName = "Shared Coordinates Source", SelectedLink = "<None>" }
        };
    }

    /// <summary>
    /// Data model for a single category-to-link mapping.
    /// </summary>
    public class CategoryMapping
    {
        public string CategoryName { get; set; }
        public string SelectedLink { get; set; }
    }

    public class CoordinateSystemMapping
    {
        public string SystemName { get; set; }
        public string SelectedLink { get; set; }
    }

    /// <summary>
    /// ViewModel for a single row in the category mapping DataGrids.
    /// Handles the logic for group assignments and individual overrides.
    /// </summary>
    public class CategoryMappingViewModel : INotifyPropertyChanged
    {
        private readonly ProfileSettingsWindow _owner;
        private List<CategoryMappingViewModel> _group;
        private bool _isOverridden = false;
        private string _selectedLink = "<None>";
        private string _groupLink = "<None>";

        public string Discipline { get; set; }
        public string CategoryName { get; set; }
        [JsonIgnore] public BuiltInCategory BuiltInCategory { get; set; }

        public CategoryMappingViewModel(ProfileSettingsWindow owner)
        {
            _owner = owner;
        }

        public void SetGroup(List<CategoryMappingViewModel> group)
        {
            _group = group;
        }

        public string SelectedLink
        {
            get => _selectedLink;
            set
            {
                if (_selectedLink != value)
                {
                    _selectedLink = value;
                    _isOverridden = true; // Any manual selection is an override
                    OnPropertyChanged(nameof(SelectedLink));
                }
            }
        }

        public string GroupLink
        {
            get => _groupLink;
            set
            {
                if (_groupLink != value)
                {
                    _groupLink = value;
                    // Apply this link to all items in the group that haven't been manually overridden
                    foreach (var item in _group)
                    {
                        if (!item._isOverridden)
                        {
                            item._selectedLink = value; // Set backing field to avoid triggering override
                            item.OnPropertyChanged(nameof(SelectedLink));
                        }
                        item._groupLink = value; // Sync the group link for all items
                        item.OnPropertyChanged(nameof(GroupLink));
                    }
                }
            }
        }

        public CategoryMapping ToDataModel()
        {
            return new CategoryMapping { CategoryName = this.CategoryName, SelectedLink = this.SelectedLink };
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    #endregion

    #region Static Category Data
    /// <summary>
    /// Contains the curated lists of model and annotation categories.
    /// </summary>
    public class CategoryData
    {
        public List<(string Discipline, string CategoryName, BuiltInCategory Bic)> ModelCategories { get; }
        public List<(string Discipline, string CategoryName, BuiltInCategory Bic)> AnnotationCategories { get; }

        public CategoryData()
        {
            ModelCategories = new List<(string, string, BuiltInCategory)>
            {
                // Core Architectural
                ("Core Architectural", "Walls", BuiltInCategory.OST_Walls),
                ("Core Architectural", "Doors", BuiltInCategory.OST_Doors),
                ("Core Architectural", "Windows", BuiltInCategory.OST_Windows),
                ("Core Architectural", "Roofs", BuiltInCategory.OST_Roofs),
                ("Core Architectural", "Floors", BuiltInCategory.OST_Floors),
                ("Core Architectural", "Ceilings", BuiltInCategory.OST_Ceilings),
                ("Core Architectural", "Stairs", BuiltInCategory.OST_Stairs),
                ("Core Architectural", "Ramps", BuiltInCategory.OST_Ramps),
                ("Core Architectural", "Rooms", BuiltInCategory.OST_Rooms),
                ("Core Architectural", "Furniture", BuiltInCategory.OST_Furniture),
                ("Core Architectural", "Casework", BuiltInCategory.OST_Casework),

                // Core Structural
                ("Core Structural", "Structural Framing", BuiltInCategory.OST_StructuralFraming),
                ("Core Structural", "Structural Columns", BuiltInCategory.OST_StructuralColumns),

                // Electrical
                ("Electrical", "Cable Tray", BuiltInCategory.OST_CableTray),
                ("Electrical", "Cable Tray Fitting", BuiltInCategory.OST_CableTrayFitting),
                ("Electrical", "Conduit", BuiltInCategory.OST_Conduit),
                ("Electrical", "Conduit Fitting", BuiltInCategory.OST_ConduitFitting),
                ("Electrical", "Communication Devices", BuiltInCategory.OST_CommunicationDevices),
                ("Electrical", "Data Devices", BuiltInCategory.OST_DataDevices),
                ("Electrical", "Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
                ("Electrical", "Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
                ("Electrical", "Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
                ("Electrical", "Lighting Devices", BuiltInCategory.OST_LightingDevices),
                ("Electrical", "Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
                ("Electrical", "Security Devices", BuiltInCategory.OST_SecurityDevices),
                ("Electrical", "Telephone Devices", BuiltInCategory.OST_TelephoneDevices),

                // Mechanical / HVAC
                ("Mechanical / HVAC", "Duct Curves", BuiltInCategory.OST_DuctCurves),
                ("Mechanical / HVAC", "Duct Fitting", BuiltInCategory.OST_DuctFitting),
                ("Mechanical / HVAC", "Duct Terminal", BuiltInCategory.OST_DuctTerminal),
                ("Mechanical / HVAC", "Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),

                // Hydraulic / Plumbing
                ("Hydraulic / Plumbing", "Pipe Curves", BuiltInCategory.OST_PipeCurves),
                ("Hydraulic / Plumbing", "Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
                ("Hydraulic / Plumbing", "Sprinklers", BuiltInCategory.OST_Sprinklers),

                // General / Multi-Discipline
                ("General / Multi-Discipline", "Generic Model", BuiltInCategory.OST_GenericModel),
                ("General / Multi-Discipline", "Assemblies", BuiltInCategory.OST_Assemblies),
                ("General / Multi-Discipline", "Parts", BuiltInCategory.OST_Parts),
            };

            AnnotationCategories = new List<(string, string, BuiltInCategory Bic)>
            {
                // General Annotations
                ("General Annotations", "Text Notes", BuiltInCategory.OST_TextNotes),
                ("General Annotations", "Dimensions", BuiltInCategory.OST_Dimensions),
                ("General Annotations", "Multi Category Tags", BuiltInCategory.OST_MultiCategoryTags),
                ("General Annotations", "Generic Annotation", BuiltInCategory.OST_GenericAnnotation),
                ("General Annotations", "Detail Components", BuiltInCategory.OST_DetailComponents),

                // Project Datums & Views
                ("Project Datums & Views", "Grids", BuiltInCategory.OST_Grids),
                ("Project Datums & Views", "Levels", BuiltInCategory.OST_Levels),
                ("Project Datums & Views", "Section Box", BuiltInCategory.OST_SectionBox),
                ("Project Datums & Views", "Viewports", BuiltInCategory.OST_Viewports),
            };

            // Add version-specific categories using preprocessor directives
#if REVIT2024_OR_GREATER
            // ModelCategories.Add(("Core Structural", "Structural Foundations", BuiltInCategory.OST_StructuralFoundations));
            // ModelCategories.Add(("Core Structural", "Structural Trusses", BuiltInCategory.OST_StructuralTrusses));
            // ModelCategories.Add(("Hydraulic / Plumbing", "Pipe Fittings", BuiltInCategory.OST_PipeFittings));
            // AnnotationCategories.Add(("Project Datums & Views", "Reference Plane", BuiltInCategory.OST_ReferencePlane));
            // AnnotationCategories.Add(("Project Datums & Views", "Scope Boxes", BuiltInCategory.OST_ScopeBoxes));
#else
            // ModelCategories.Add(("Core Structural", "Structural Foundation", BuiltInCategory.OST_StructuralFoundation));
            // ModelCategories.Add(("Core Structural", "Truss", BuiltInCategory.OST_Truss));
            // ModelCategories.Add(("Hydraulic / Plumbing", "Pipe Fitting", BuiltInCategory.OST_PipeFitting));
            // AnnotationCategories.Add(("Project Datums & Views", "Reference Planes", BuiltInCategory.OST_ReferencePlanes));
            // AnnotationCategories.Add(("Project Datums & Views", "Scope Box", BuiltInCategory.OST_ScopeBox));
#endif
        }
    }
    #endregion
}
