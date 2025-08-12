using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media.Imaging;

namespace RTS.UI
{
    /// <summary>
    /// Represents a schedule item in the UI for the MapCables command.
    /// </summary>
    public class ScheduleMappingItem
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }
    }

    /// <summary>
    /// Interaction logic for MapCablesWindow.xaml
    /// </summary>
    public partial class MapCablesWindow : Window
    {
        public List<ScheduleMappingItem> Schedules { get; set; }
        public bool GenerateTemplates { get; private set; }
        public bool MapFromFiles { get; private set; }

        public MapCablesWindow(List<ScheduleMappingItem> schedules)
        {
            InitializeComponent();
            LoadIcon(); // Dynamically load the window icon
            Schedules = schedules;
            SchedulesList.ItemsSource = Schedules;
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

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            // Store the user's action choice
            GenerateTemplates = GenerateTemplatesRadio.IsChecked == true;
            MapFromFiles = MapFromFilesRadio.IsChecked == true;

            // The 'Schedules' list is already bound to the UI, so any checkbox
            // changes are automatically reflected in the IsSelected property.

            this.DialogResult = true;
        }
    }
}
