//RoutingReportProgressBarWindow.xaml.cs

using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for RoutingReportProgressBarWindow.xaml
    /// </summary>
    public partial class RoutingReportProgressBarWindow : Window
    {
        public bool IsCancelled { get; private set; }
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private string _completedFilePath;

        public RoutingReportProgressBarWindow()
        {
            InitializeComponent();
            LoadIcon(); // Load the icon on initialization
            IsCancelled = false;
            _stopwatch.Start();
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

        /// <summary>
        /// Updates the UI specifically for the initial graph-building phase.
        /// </summary>
        public void UpdateGraphProgress(int currentElement, int totalElements, double percentage, string taskDescription)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateGraphProgress(currentElement, totalElements, percentage, taskDescription));
                return;
            }

            ProcessStepText.Text = "Step 1 of 2: Building Network Graph";
            CableDetailsGroup.IsEnabled = false;
            CableReferenceText.Text = "-";
            FromText.Text = "-";
            ToText.Text = "-";

            ProgressBar.Value = percentage;
            ProgressPercentageText.Text = $"{percentage:F0}%";
            TaskDescriptionText.Text = taskDescription;
        }


        /// <summary>
        /// Updates all UI elements of the progress window for the cable routing phase.
        /// </summary>
        public void UpdateProgress(int currentStep, int totalSteps, string currentItemName, string from, string to, double percentage, string taskDescription)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateProgress(currentStep, totalSteps, currentItemName, from, to, percentage, taskDescription));
                return;
            }

            ProcessStepText.Text = $"Step 2 of 2: Processing Cable Routes ({currentStep} of {totalSteps})";
            CableDetailsGroup.IsEnabled = true;

            ProgressBar.Value = percentage;
            ProgressPercentageText.Text = $"{percentage:F0}%";

            CableReferenceText.Text = currentItemName;
            FromText.Text = from;
            ToText.Text = to;

            string timeEstimate = "";
            if (currentStep > 0 && _stopwatch.Elapsed.TotalSeconds > 1)
            {
                var elapsed = _stopwatch.Elapsed;
                var timePerItem = elapsed.TotalSeconds / currentStep;
                var remainingItems = totalSteps - currentStep;
                var remainingTime = TimeSpan.FromSeconds(timePerItem * remainingItems);
                timeEstimate = $" (est. {FormatTimeSpan(remainingTime)} remaining)";
            }

            TaskDescriptionText.Text = taskDescription + timeEstimate;
            ErrorText.Text = "";
        }

        /// <summary>
        /// Transforms the UI to show the completion status and final action buttons.
        /// </summary>
        public void ShowCompletion(string filePath)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => ShowCompletion(filePath));
                return;
            }

            _completedFilePath = filePath;
            ProcessStepText.Text = "Export Complete!";
            TaskDescriptionText.Text = $"Report saved to: {filePath}";

            ProgressBar.Value = 100;
            ProgressPercentageText.Text = "100%";
            ProgressBar.Foreground = (SolidColorBrush)FindResource("SuccessBrush");

            CancelButton.Visibility = Visibility.Collapsed;
            OpenFolderButton.Visibility = Visibility.Visible;
            CloseButton.Visibility = Visibility.Visible;
        }

        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalMinutes >= 1)
            {
                return $"{(int)ts.TotalMinutes} min {ts.Seconds} sec";
            }
            return $"{(int)ts.TotalSeconds} sec";
        }

        public void ShowError(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => ShowError(message));
                return;
            }
            ErrorText.Text = $"Error: {message}";
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            CancelButton.IsEnabled = false;
            TaskDescriptionText.Text = "Cancellation requested, finishing current task...";
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(_completedFilePath))
                {
                    string folderPath = Path.GetDirectoryName(_completedFilePath);
                    if (Directory.Exists(folderPath))
                    {
                        Process.Start("explorer.exe", folderPath);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Could not open folder: {ex.Message}");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
