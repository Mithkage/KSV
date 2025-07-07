// InputWindow.xaml.cs
//
// Description:
// This file contains the C# code-behind for the InputWindow.xaml, which serves
// as a user interface for receiving text input and defining actions for a Revit add-in.
// It allows users to input text, confirm an action (OK), or cancel the operation.
//
// Script Function:
// The InputWindow class manages user interaction for a simple input dialog. It provides
// a text box for user input, and buttons to trigger specific actions (OK, Cancel).
// It exposes properties to retrieve the user's input and their chosen action,
// facilitating communication with the main application logic.
//
// Change Log:
// 2025-07-04: Added file description, script function description, and change log.
// 2025-07-04: Removed ClearGraphics WindowAction and associated button click handler
//             as the functionality is now handled by submitting a blank input.
//
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
        Cancel
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
        /// Gets the action chosen by the user (OK or Cancel).
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
        /// Handles the click event for the Cancel button.
        /// </summary>
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            ActionChosen = WindowAction.Cancel;
            this.DialogResult = false; // Signals that the dialog was cancelled
        }
    }
}