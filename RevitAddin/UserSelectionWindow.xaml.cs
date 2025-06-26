using System.Windows;

namespace RTS_MapCables
{
    public partial class UserSelectionWindow : Window
    {
        public bool GenerateTemplate { get; private set; }

        public UserSelectionWindow()
        {
            InitializeComponent();
        }

        private void YesButton_Click(object sender, RoutedEventArgs e)
        {
            GenerateTemplate = true;
            DialogResult = true; // Closes the window with a true result
        }

        private void NoButton_Click(object sender, RoutedEventArgs e)
        {
            GenerateTemplate = false;
            DialogResult = false; // Closes the window with a false result
        }
    }
}