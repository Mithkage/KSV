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
        // These are separate from the main Link Hub profile storage.
        public static readonly Guid SettingsSchemaGuid = new Guid("E8C5B1A0-1B1C-4F7B-8E7A-6A0C9D1B3E2F");
        public const string SettingsSchemaName = "RTS_ProfileSettingsSchema";
        public const string SettingsFieldName = "ProfileSettingsJson";
        public const string SettingsDataStorageElementName = "RTS_ProfileSettings_Storage";
        public const string VendorId = "ReTick_Solutions";

        /// <summary>
        /// Holds the settings data that is bound to the UI controls.
        /// </summary>
        public ProfileSettings Settings { get; set; }

        /// <summary>
        /// A collection of formatted link names to populate the dropdowns.
        /// </summary>
        public ObservableCollection<string> AvailableLinks { get; set; }

        /// <summary>
        /// Constructor for the Profile Settings window.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="links">The collection of links from the main Link Hub window.</param>
        public ProfileSettingsWindow(Document doc, IEnumerable<LinkViewModel> links)
        {
            InitializeComponent();
            _doc = doc;
            this.DataContext = this;

            // Populate the dropdown source list from the links in the main window.
            AvailableLinks = new ObservableCollection<string> { "<None>" }; // Add the blank/none option first.
            foreach (var link in links.Where(l => l.IsRevitLink)) // Only include actual Revit links.
            {
                AvailableLinks.Add($"[{link.SelectedDiscipline}] - {link.LinkName}");
            }

            // Load existing settings from the project or create a new settings object.
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

        #region Extensible Storage Helpers
        // These are generic helper methods for saving and recalling data from Revit's Extensible Storage.

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
    }

    /// <summary>
    /// Data model for storing the profile settings. Implements INotifyPropertyChanged
    /// to allow the UI to update automatically when properties change.
    /// </summary>
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
    }
}
