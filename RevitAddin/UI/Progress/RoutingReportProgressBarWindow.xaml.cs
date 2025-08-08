using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
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
        private readonly NotifyIcon _notifyIcon;

        public RoutingReportProgressBarWindow()
        {
            InitializeComponent();
            LoadIcon(); // Load the icon on initialization
            IsCancelled = false;
            _stopwatch.Start();

            // Initialize the NotifyIcon
            _notifyIcon = new NotifyIcon
            {
                Text = "Routing Report Progress",
                Visible = false
            };

            // Set the icon for the NotifyIcon (must be .ico)
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.ico", UriKind.Absolute);
                var iconStream = System.Windows.Application.GetResourceStream(iconUri)?.Stream;
                if (iconStream != null)
                {
                    _notifyIcon.Icon = new System.Drawing.Icon(iconStream);
                }
            }
            catch
            {
                _notifyIcon.Icon = System.Drawing.SystemIcons.Application;
            }

            // Handle double-click to restore the window
            _notifyIcon.DoubleClick += (s, args) => RestoreWindow();

            // Create a context menu for the tray icon
            _notifyIcon.ContextMenu = new ContextMenu();
            _notifyIcon.ContextMenu.MenuItems.Add("Restore", (s, args) => RestoreWindow());
        }

        /// <summary>
        /// Loads the window icon using a robust Pack URI.
        /// </summary>
        private void LoadIcon()
        {
            try
            {
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                // Note: Window icon can be .png, but NotifyIcon requires .ico. 
                // We use the .ico for both for consistency.
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.ico", UriKind.Absolute);
                this.Icon = new BitmapImage(iconUri);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning: Could not load window icon. {ex.Message}");
            }
        }

        /// <summary>
        /// Handles the window's state changed event to show/hide the tray icon.
        /// </summary>
        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                this.Hide();
                _notifyIcon.Visible = true;
            }
        }

        /// <summary>
        /// Restores the window from the system tray.
        /// </summary>
        private void RestoreWindow()
        {
            this.Show();
            this.WindowState = WindowState.Normal;
            this.Activate();
            _notifyIcon.Visible = false;
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

            // Set the UI state for graph building
            ProcessStepText.Text = "Step 1 of 2: Building Network Graph";
            CableDetailsGroup.IsEnabled = false; // Disable cable-specific details
            CableReferenceText.Text = "-";
            FromText.Text = "-";
            ToText.Text = "-";

            // Update progress and task description
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

            // Set the UI state for cable processing
            ProcessStepText.Text = $"Step 2 of 2: Processing Cable Routes ({currentStep} of {totalSteps})";
            CableDetailsGroup.IsEnabled = true; // Enable and populate cable-specific details

            // Update the progress bar and the percentage text overlay
            ProgressBar.Value = percentage;
            ProgressPercentageText.Text = $"{percentage:F0}%";

            // Update the specific cable information
            CableReferenceText.Text = currentItemName;
            FromText.Text = from;
            ToText.Text = to;

            // Calculate and display estimated time remaining
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

            // Clear any previous error messages
            ErrorText.Text = "";
        }

        /// <summary>
        /// Formats a TimeSpan into a user-friendly string like "1 min 25 sec".
        /// </summary>
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

        /// <summary>
        /// Clean up the NotifyIcon when the window is closed.
        /// </summary>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            _notifyIcon.Dispose();
        }
    }
}
