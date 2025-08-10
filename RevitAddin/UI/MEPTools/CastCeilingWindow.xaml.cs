using System.Windows;

namespace RTS.UI
{
    public partial class CastCeilingWindow : Window
    {
        public bool UseCeilings => CeilingsCheckBox.IsChecked == true;
        public bool UseSlabs => SlabsCheckBox.IsChecked == true;
        public bool UseFastProcessing => FastProcessingCheckBox.IsChecked == true;
        public bool UseAccurateProcessing => AccurateProcessingCheckBox.IsChecked == true;
        public bool SearchActiveViewOnly => ActiveViewOnlyCheckBox.IsChecked == true;

        public CastCeilingWindow()
        {
            InitializeComponent();
        }

        private void SelectElementsButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate at least one surface type and one processing method
            if (!UseCeilings && !UseSlabs)
            {
                MessageBox.Show("Please select at least one surface type (Ceilings or Slabs).", "Input Required");
                return;
            }
            if (!UseFastProcessing && !UseAccurateProcessing)
            {
                MessageBox.Show("Please select a processing method.", "Input Required");
                return;
            }
            if (UseFastProcessing && UseAccurateProcessing)
            {
                MessageBox.Show("Please select only one processing method.", "Input Required");
                return;
            }

            // Store options as needed, then close and trigger selection in command
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}