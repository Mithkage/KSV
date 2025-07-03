using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Required for Keyboard.Focus

namespace RT_Isolate
{
    /// <summary>
    /// Defines the possible actions chosen by the user in the input window.
    /// </summary>
    public enum WindowAction
    {
        Ok,
        Cancel,
        ClearGraphics
    }

    /// <summary>
    /// Interaction logic for InputWindow.xaml
    /// </summary>
    public partial class InputWindow : Window
    {
        /// <summary>
        /// Public property to get and set the text in the InputTextBox.
        /// </summary>
        public string UserInput
        {
            get { return InputTextBox.Text; }
            set { InputTextBox.Text = value; }
        }

        /// <summary>
        /// Gets the action chosen by the user (OK, Cancel, or ClearGraphics).
        /// </summary>
        public WindowAction ActionChosen { get; private set; } = WindowAction.Cancel; // Default to Cancel

        /// <summary>
        /// Constructor for the InputWindow.
        /// </summary>
        public InputWindow()
        {
            InitializeComponent();

            // Set focus to the textbox when the window loads
            Loaded += (sender, e) =>
            {
                InputTextBox.Focus();
                InputTextBox.SelectAll(); // Select all text for easy replacement
            };
        }

        /// <summary>
        /// Handles the click event for the OK button.
        /// </summary>
        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            ActionChosen = WindowAction.Ok;
            this.DialogResult = true; // Signals that the dialog was accepted
        }

        /// <summary>
        /// Handles the click event for the Clear Overrides button.
        /// </summary>
        private void ClearOverrides_Click(object sender, RoutedEventArgs e)
        {
            ActionChosen = WindowAction.ClearGraphics;
            this.DialogResult = true; // Signals acceptance, but with a different action
        }

        /// <summary>
        /// Handles the click event for the Cancel button.
        /// </summary>
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            ActionChosen = WindowAction.Cancel;
            this.DialogResult = false; // Signals that the dialog was cancelled
        }
    }
}
