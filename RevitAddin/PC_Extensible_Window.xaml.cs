using System.Windows;

namespace PC_Extensible
{
    /// <summary>
    /// Enum to represent the user's choice from the WPF window.
    /// </summary>
    public enum UserAction
    {
        None,
        ImportMyData,
        ImportConsultantData,
        ClearData,
        Cancel
    }

    /// <summary>
    /// Interaction logic for PC_Extensible_Window.xaml
    /// </summary>
    public partial class PC_Extensible_Window : Window
    {
        /// <summary>
        /// Stores the action selected by the user.
        /// </summary>
        public UserAction SelectedAction { get; private set; } = UserAction.None;

        /// <summary>
        /// Constructor for the window.
        /// </summary>
        public PC_Extensible_Window()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Handles the click event for the "Import My Data" button.
        /// Sets the selected action and closes the dialog with a 'true' result.
        /// </summary>
        private void ImportMyDataButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = UserAction.ImportMyData;
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handles the click event for the "Import Consultant Data" button.
        /// Sets the selected action and closes the dialog with a 'true' result.
        /// </summary>
        private void ImportConsultantDataButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = UserAction.ImportConsultantData;
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handles the click event for the "Clear Data" button.
        /// Sets the selected action and closes the dialog with a 'true' result.
        /// The main command will then handle the sub-dialog for clearing options.
        /// </summary>
        private void ClearDataButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = UserAction.ClearData;
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handles the click event for the "Cancel" button.
        /// Sets the selected action and closes the dialog with a 'false' result.
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = UserAction.Cancel;
            this.DialogResult = false;
            this.Close();
        }
    }
}
