using Autodesk.Revit.DB;
using System.Windows;
using System.Windows.Controls;

namespace RTS.UI
{
    public partial class InitiateParametersWindow : Window
    {
        public bool InitiateAllProjectParameters { get; private set; }
        public bool InitiateMapCablesParameters { get; private set; }
        public bool InitiateScheduleParameters { get; private set; }
        public bool AddRtsTypeToFamily { get; private set; }
        public bool InitiateCopyRelativeParameters { get; private set; }

        // Properties for schedule sub-options
        public bool InitiateScheduleDetailItems { get; private set; }
        public bool InitiateScheduleElecEquip { get; private set; }
        public bool InitiateScheduleLighting { get; private set; }
        public bool InitiateScheduleElecFixtures { get; private set; }
        public bool InitiateScheduleCableTrays { get; private set; }

        public InitiateParametersWindow(Document doc)
        {
            InitializeComponent();

            if (doc.IsFamilyDocument)
            {
                AddRtsTypeFamilyCheckBox.IsEnabled = true;
                AddRtsTypeFamilyCheckBox.IsChecked = true;
                CopyRelativeCheckBox.IsEnabled = true;
                CopyRelativeCheckBox.IsChecked = true;

                InitiateAllParamsCheckBox.IsEnabled = false;
                InitiateAllParamsCheckBox.IsChecked = false;
                MapCablesCheckBox.IsEnabled = false;
                MapCablesCheckBox.IsChecked = false;
                SchedulesCheckBox.IsEnabled = false;
                SchedulesCheckBox.IsChecked = false;
            }
        }

        private void InitiateButton_Click(object sender, RoutedEventArgs e)
        {
            InitiateAllProjectParameters = InitiateAllParamsCheckBox.IsChecked == true;
            InitiateMapCablesParameters = MapCablesCheckBox.IsChecked == true;
            InitiateScheduleParameters = SchedulesCheckBox.IsChecked == true;
            AddRtsTypeToFamily = AddRtsTypeFamilyCheckBox.IsChecked == true;
            InitiateCopyRelativeParameters = CopyRelativeCheckBox.IsChecked == true;

            // Capture sub-option states
            InitiateScheduleDetailItems = ScheduleDetailItemsCheckBox.IsChecked == true;
            InitiateScheduleElecEquip = ScheduleElecEquipCheckBox.IsChecked == true;
            InitiateScheduleLighting = ScheduleLightingCheckBox.IsChecked == true;
            InitiateScheduleElecFixtures = ScheduleElecFixturesCheckBox.IsChecked == true;
            InitiateScheduleCableTrays = ScheduleCableTraysCheckBox.IsChecked == true;

            this.DialogResult = true;
        }

        private void SchedulesCheckBox_Click(object sender, RoutedEventArgs e)
        {
            bool isChecked = SchedulesCheckBox.IsChecked == true;
            ScheduleDetailItemsCheckBox.IsChecked = isChecked;
            ScheduleElecEquipCheckBox.IsChecked = isChecked;
            ScheduleLightingCheckBox.IsChecked = isChecked;
            ScheduleElecFixturesCheckBox.IsChecked = isChecked;
            ScheduleCableTraysCheckBox.IsChecked = isChecked;
        }
    }
}
