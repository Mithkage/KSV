using System.Windows;

namespace RT_Isolate
{
    /// <summary>
    /// Interaction logic for InputWindow.xaml
    /// </summary>
    public partial class InputWindow : Window
    {
        public string UserInput { get; private set; }

        public InputWindow()
        {
            InitializeComponent();
            // Optional: Set focus to the textbox when the window opens
            this.Loaded += (sender, e) => InputTextBox.Focus();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            UserInput = InputTextBox.Text;
            DialogResult = true; // Indicates success
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            UserInput = null; // Or string.Empty, depending on desired behavior for cancel
            DialogResult = false; // Indicates cancellation
            this.Close();
        }
    }
}