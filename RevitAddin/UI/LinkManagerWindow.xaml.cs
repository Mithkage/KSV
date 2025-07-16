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
//           supports adding and deleting placeholder rows for links.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 16, 2025: Added Company Name field and enabled column sorting.
// - July 16, 2025: Added functionality to delete placeholder rows with the Delete key.
// - July 16, 2025: Added "Add Placeholder" functionality and default coordinate settings.
// - July 16, 2025: Added Link Description, Last Modified Date, and Comments fields.
// - July 16, 2025: Corrected JsonIgnore attribute usage to resolve compiler errors.
//

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Required for KeyEventArgs

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for LinkManagerWindow.xaml
    /// </summary>
    public partial class LinkManagerWindow : Window
    {
        private Document _doc;

        // --- Extensible Storage Definitions for Link Manager Profile ---
        public static readonly Guid ProfileSchemaGuid = new Guid("D4C6E8B0-6A2E-4B1C-9D7E-8C4F2A6B9E1F");
        public const string ProfileSchemaName = "RTS_LinkManagerProfileSchema";
        public const string ProfileFieldName = "LinkManagerProfileJson";
        public const string ProfileDataStorageElementName = "RTS_LinkManager_Profile_Storage";
        public const string VendorId = "ReTick_Solutions"; // Must match your .addin file

        public ObservableCollection<LinkViewModel> Links { get; set; }

        public LinkManagerWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            this.DataContext = this;

            Links = new ObservableCollection<LinkViewModel>();
            LoadLinks();
        }

        /// <summary>
        /// Scans the Revit document for all RevitLinkInstances and populates the DataGrid,
        /// loading any previously saved profile data from Extensible Storage.
        /// </summary>
        private void LoadLinks()
        {
            // 1. Recall existing profile data from storage
            var savedProfiles = RecallDataFromExtensibleStorage<LinkViewModel>(_doc, ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);
            var savedProfileDict = savedProfiles.ToDictionary(p => p.LinkName, p => p);

            // 2. Get all current link instances in the project
            var collector = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance));
            var loadedLinkNames = new HashSet<string>();

            foreach (RevitLinkInstance instance in collector)
            {
                RevitLinkType type = _doc.GetElement(instance.GetTypeId()) as RevitLinkType;
                if (type == null || type.IsNestedLink) continue;

                loadedLinkNames.Add(type.Name);

                // 3. Create the ViewModel and populate with saved data or defaults
                var viewModel = new LinkViewModel
                {
                    LinkInstanceId = instance.Id,
                    LinkName = type.Name
                };

                // Get the file path to retrieve the last modified date
                string lastModifiedDate = "Not Found";
                try
                {
                    ModelPath modelPath = type.GetExternalFileReference()?.GetPath();
                    if (modelPath != null)
                    {
                        string userVisiblePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                        if (File.Exists(userVisiblePath))
                        {
                            lastModifiedDate = File.GetLastWriteTime(userVisiblePath).ToString("yyyy-MM-dd HH:mm");
                        }
                    }
                }
                catch { /* Ignore errors in finding path */ }
                viewModel.LastModified = lastModifiedDate;


                if (savedProfileDict.TryGetValue(type.Name, out var savedProfile))
                {
                    // A profile was found, load the saved data
                    viewModel.LinkDescription = savedProfile.LinkDescription;
                    viewModel.SelectedDiscipline = savedProfile.SelectedDiscipline;
                    viewModel.CompanyName = savedProfile.CompanyName;
                    viewModel.ResponsiblePerson = savedProfile.ResponsiblePerson;
                    viewModel.ContactDetails = savedProfile.ContactDetails;
                    viewModel.Comments = savedProfile.Comments;
                }
                else
                {
                    // No profile found, use defaults
                    viewModel.LinkDescription = "";
                    viewModel.SelectedDiscipline = "Architectural";
                    viewModel.CompanyName = "Not Set";
                    viewModel.ResponsiblePerson = "Not Set";
                    viewModel.ContactDetails = "Not Set";
                    viewModel.Comments = "";
                }

                // 4. Read the live coordinate setting from the instance
                Parameter sharedParam = null;
                foreach (Parameter p in instance.Parameters)
                {
                    if (p.Definition.Name == "Shared Site")
                    {
                        sharedParam = p;
                        break;
                    }
                }

                if (sharedParam != null && sharedParam.AsInteger() == 1)
                {
                    viewModel.SelectedCoordinates = "Shared Coordinates";
                }
                else
                {
                    // Default to Shared Coordinates if it can't be determined or is set to Origin
                    viewModel.SelectedCoordinates = "Shared Coordinates";
                }

                Links.Add(viewModel);
            }

            // Add any placeholders from the saved profile that are not currently loaded in the project
            foreach (var savedProfile in savedProfiles)
            {
                if (savedProfile.LinkInstanceId == ElementId.InvalidElementId && !Links.Any(l => l.LinkName == savedProfile.LinkName))
                {
                    savedProfile.LastModified = "N/A"; // Ensure placeholder state is correct
                    Links.Add(savedProfile);
                }
            }
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
                    // 1. Save the metadata profile to Extensible Storage
                    SaveDataToExtensibleStorage(_doc, Links.ToList(), ProfileSchemaGuid, ProfileSchemaName, ProfileFieldName, ProfileDataStorageElementName);

                    // 2. Update the coordinate system for each link instance if it was changed
                    foreach (var vm in Links)
                    {
                        // Only process real links, not placeholders
                        if (vm.LinkInstanceId == null || vm.LinkInstanceId == ElementId.InvalidElementId) continue;

                        RevitLinkInstance instance = _doc.GetElement(vm.LinkInstanceId) as RevitLinkInstance;
                        if (instance == null) continue;

                        Parameter sharedParam = null;
                        foreach (Parameter p in instance.Parameters)
                        {
                            if (p.Definition.Name == "Shared Site")
                            {
                                sharedParam = p;
                                break;
                            }
                        }

                        if (sharedParam != null && !sharedParam.IsReadOnly)
                        {
                            int currentValue = sharedParam.AsInteger();
                            int newValue = (vm.SelectedCoordinates == "Shared Coordinates") ? 1 : 0;
                            if (currentValue != newValue)
                            {
                                sharedParam.Set(newValue);
                            }
                        }
                    }

                    tx.Commit();
                    MessageBox.Show("Project link profile saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.DialogResult = true;
                    this.Close();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    MessageBox.Show($"Failed to save profile: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Handles the click event for the "Add Placeholder" button.
        /// </summary>
        private void AddPlaceholderButton_Click(object sender, RoutedEventArgs e)
        {
            var placeholder = new LinkViewModel
            {
                LinkName = "Placeholder",
                LinkDescription = "Awaiting Link Model",
                LastModified = "N/A",
                SelectedDiscipline = "Architectural",
                SelectedCoordinates = "Shared Coordinates",
                CompanyName = "Not Set",
                ResponsiblePerson = "Not Set",
                ContactDetails = "Not Set",
                Comments = "",
                LinkInstanceId = ElementId.InvalidElementId // Important to mark as not a real element
            };
            Links.Add(placeholder);
        }

        /// <summary>
        /// Handles the PreviewKeyDown event on the DataGrid to catch the Delete key.
        /// </summary>
        private void LinksDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                var grid = sender as DataGrid;
                if (grid == null || grid.SelectedItems.Count == 0) return;

                // Get a copy of the selected items to avoid modifying the collection while iterating
                var selectedItems = grid.SelectedItems.Cast<LinkViewModel>().ToList();

                // Filter for items that are actually placeholders (they have an invalid ElementId)
                var placeholdersToDelete = selectedItems.Where(vm => vm.LinkInstanceId == ElementId.InvalidElementId).ToList();

                if (placeholdersToDelete.Any())
                {
                    var result = MessageBox.Show($"Are you sure you want to delete {placeholdersToDelete.Count} placeholder(s)?",
                                                 "Confirm Delete",
                                                 MessageBoxButton.YesNo,
                                                 MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        foreach (var placeholder in placeholdersToDelete)
                        {
                            Links.Remove(placeholder);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Only placeholder rows can be deleted from this manager.", "Deletion Not Allowed", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                // Mark the event as handled to prevent further processing
                e.Handled = true;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        #region Extensible Storage Methods

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

        public void SaveDataToExtensibleStorage<T>(Document doc, List<T> dataList, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName)
        {
            Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName);
            DataStorage dataStorage = GetOrCreateDataStorage(doc, dataStorageElementName);
            string jsonString = JsonSerializer.Serialize(dataList, new JsonSerializerOptions { WriteIndented = true });

            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(fieldName), jsonString);
            dataStorage.SetEntity(entity);
        }

        public List<T> RecallDataFromExtensibleStorage<T>(Document doc, Guid schemaGuid, string schemaName, string fieldName, string dataStorageElementName)
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
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to read saved profile: {ex.Message}", "Profile Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return new List<T>();
            }
        }

        #endregion
    }

    /// <summary>
    /// A view model representing a single row (a single link) in the DataGrid.
    /// Implements INotifyPropertyChanged to allow the UI to update automatically when data changes.
    /// </summary>
    public class LinkViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        [JsonIgnore]
        public ElementId LinkInstanceId { get; set; }

        public string LinkName { get; set; }

        private string _linkDescription;
        private string _lastModified;
        private string _selectedDiscipline;
        private string _selectedCoordinates;
        private string _companyName;
        private string _responsiblePerson;
        private string _contactDetails;
        private string _comments;

        public string LinkDescription
        {
            get => _linkDescription;
            set { _linkDescription = value; OnPropertyChanged(nameof(LinkDescription)); }
        }

        [JsonIgnore]
        public string LastModified
        {
            get => _lastModified;
            set { _lastModified = value; OnPropertyChanged(nameof(LastModified)); }
        }

        public string SelectedDiscipline
        {
            get => _selectedDiscipline;
            set { _selectedDiscipline = value; OnPropertyChanged(nameof(SelectedDiscipline)); }
        }

        [JsonIgnore]
        public string SelectedCoordinates
        {
            get => _selectedCoordinates;
            set { _selectedCoordinates = value; OnPropertyChanged(nameof(SelectedCoordinates)); }
        }

        public string CompanyName
        {
            get => _companyName;
            set { _companyName = value; OnPropertyChanged(nameof(CompanyName)); }
        }

        public string ResponsiblePerson
        {
            get => _responsiblePerson;
            set { _responsiblePerson = value; OnPropertyChanged(nameof(ResponsiblePerson)); }
        }

        public string ContactDetails
        {
            get => _contactDetails;
            set { _contactDetails = value; OnPropertyChanged(nameof(ContactDetails)); }
        }

        public string Comments
        {
            get => _comments;
            set { _comments = value; OnPropertyChanged(nameof(Comments)); }
        }

        [JsonIgnore]
        public List<string> AvailableDisciplines { get; } = new List<string>
        {
            "Architectural", "Structural", "Mechanical", "Electrical", "Hydraulic", "Fire", "Civil", "Landscape", "Other"
        };

        [JsonIgnore]
        public List<string> AvailableCoordinates { get; } = new List<string>
        {
            "Origin to Origin", "Shared Coordinates"
        };
    }
}
