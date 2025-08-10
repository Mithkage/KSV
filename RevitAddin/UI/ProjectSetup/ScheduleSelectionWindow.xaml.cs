//
// Copyright (c) 2025. All rights reserved.
//
// Author: ReTick Solutions
//
// This script is a Revit External Command designed to automate the creation
// of a standard set of project schedules. When executed, the script will:
// 1. Define a list of standard schedules with specific names and categories.
// 2. Present a WPF UI to the user to select which schedules to generate or remove.
// 3. For selected schedules, it checks if any of the target schedules are currently
//    open in the UI and closes them.
// 4. For each selected schedule, it checks if a schedule with the same name
//    already exists in the active Revit project.
// 5. If an existing schedule is found, it is deleted before recreation.
// 6. A new schedule is then created with the specified name and type.
// 7. Each newly created schedule is initialized with default fields and column widths.
//    - "Family" column width is set to 60mm.
//    - "Type" column width is set to 80mm.
//    - "RTS_ID" column width is set to 50mm.
// 8. For Cable Tray, Cable Tray Fitting, and Conduit schedules, a specific list of
//    shared parameters is also added. These schedules are filtered to only show items
//    where 'RTS_ID' has a value, and are sorted by 'RTS_ID' then by 'Type'.
// 9. Now, 'RTS_ID' is added to *all* generated schedules, if applicable to their category.
// 10. If "Remove All Schedules" is chosen, all predefined schedules are deleted.
// 11. A report is displayed to the user summarizing the creation/update/deletion status of each schedule.
//
// This ensures a consistent and standardized set of schedules is present in the project,
// ready for further customization, with user control over which schedules are affected.
//

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows; // For WPF
using System.Windows.Controls; // For WPF Controls
using System.Windows.Media; // For WPF Brushes
using System.Windows.Input; // For Mouse.OverrideCursor
#endregion

namespace RTS.UI // UPDATED: Namespace changed from RTS_Schedules to RTS.UI
{
    /// <summary>
    /// Represents a schedule item for the WPF UI.
    /// </summary>
    public class ScheduleItem
    {
        public string Name { get; set; }
        public BuiltInCategory Category { get; set; }
        public bool IsSelected { get; set; }
    }

    /// <summary>
    /// Interaction logic for ScheduleSelectionWindow.xaml
    /// This defines the WPF window for schedule selection.
    /// </summary>
    public partial class ScheduleSelectionWindow : Window
    {
        // Public property to expose the list of schedules selected by the user
        public List<ScheduleItem> Schedules { get; set; }
        public bool CreateAllSelected { get; private set; }
        public bool RemoveAllSelected { get; private set; }

        /// <summary>
        /// Constructor for the ScheduleSelectionWindow.
        /// </summary>
        /// <param name="schedules">The list of schedules to display and manage.</param>
        public ScheduleSelectionWindow(List<ScheduleItem> schedules)
        {
            InitializeComponent();

            Schedules = schedules;
            this.DataContext = this; // Set DataContext for binding

            // Create checkboxes dynamically
            // This assumes 'SchedulesPanel' is defined in your XAML with x:Name="SchedulesPanel"
            // and is a StackPanel or similar container.
            foreach (var schedule in Schedules)
            {
                CheckBox cb = new CheckBox
                {
                    Content = schedule.Name,
                    IsChecked = schedule.IsSelected,
                    Margin = new Thickness(5)
                };
                SchedulesPanel.Children.Add(cb); // Accessing SchedulesPanel (from XAML)
                cb.Checked += (sender, e) => schedule.IsSelected = true;
                cb.Unchecked += (sender, e) => schedule.IsSelected = false;
            }
        }

        /// <summary>
        /// Handles the click event for the "Create All" button.
        /// Selects all schedules in the list.
        /// </summary>
        private void CreateAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var schedule in Schedules)
            {
                schedule.IsSelected = true;
            }
            // Update UI checkboxes
            foreach (var child in SchedulesPanel.Children)
            {
                if (child is CheckBox cb)
                {
                    cb.IsChecked = true;
                }
            }
            CreateAllSelected = true;
            RemoveAllSelected = false; // Ensure only one mode is active
        }

        /// <summary>
        /// Handles the click event for the "Remove All" button.
        /// Selects all schedules for removal.
        /// </summary>
        private void RemoveAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var schedule in Schedules)
            {
                schedule.IsSelected = true; // Select all for removal
            }
            // Update UI checkboxes
            foreach (var child in SchedulesPanel.Children)
            {
                if (child is CheckBox cb)
                {
                    cb.IsChecked = true;
                }
            }
            RemoveAllSelected = true;
            CreateAllSelected = false; // Ensure only one mode is active
        }

        /// <summary>
        /// Handles the click event for the "Generate Selected" button.
        /// Sets DialogResult to true and closes the window.
        /// </summary>
        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handles the click event for the "Cancel" button.
        /// Sets DialogResult to false and closes the window.
        /// </summary>
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
