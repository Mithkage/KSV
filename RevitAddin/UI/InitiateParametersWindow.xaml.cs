using System.Windows;

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for InitiateParametersWindow.xaml
    /// </summary>
    public partial class InitiateParametersWindow : Window
    {
        public bool InitiateDetailItems { get; private set; }
        public bool InitiateCableConduit { get; private set; }
        public bool InitiateRtsId { get; private set; }
        public bool InitiateWires { get; private set; }

        public InitiateParametersWindow()
        {
            InitializeComponent();
        }

        private void InitiateButton_Click(object sender, RoutedEventArgs e)
        {
            // Store the user's selections in properties
            InitiateDetailItems = DetailItemsCheckBox.IsChecked == true;
            InitiateCableConduit = CableConduitCheckBox.IsChecked == true;
            InitiateRtsId = RtsIdCheckBox.IsChecked == true;
            InitiateWires = WiresCheckBox.IsChecked == true;

            // Set the DialogResult to true to indicate the user clicked "Initiate"
            this.DialogResult = true;
        }
    }
}
