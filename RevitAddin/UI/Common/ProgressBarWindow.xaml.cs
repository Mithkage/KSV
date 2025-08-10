//
// --- FILE: ProgressBarWindow.xaml.cs ---
//
// Description:
// Code-behind for the ProgressBarWindow.xaml. It provides methods to update
// the progress and status messages, and handles the cancellation logic.
//
// Author: Kyle Vorster
// Company: ReTick Solutions Pty Ltd
//
// Change Log:
// - August 9, 2025: Updated icon reference to use .png for consistency with other non-tray windows.
// - August 9, 2025: Re-aligned methods to be compatible with older command classes.
// - August 9, 2025: Updated to support the new UI and provide a flexible UpdateStatus method.
// - July 31, 2025: Added UpdateRoomStatus method for more detailed feedback.
// - July 30, 2025: Initial creation of the code-behind.
//

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media.Imaging;

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for ProgressBarWindow.xaml
    /// </summary>
    public partial class ProgressBarWindow : Window
    {
        private readonly Stopwatch _stopwatch = new Stopwatch();
        public bool IsCancellationPending { get; private set; }

        public ProgressBarWindow()
        {
            InitializeComponent();
            LoadIcon();
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
                Uri iconUri = new Uri($"pack://application:,,,/{assemblyName};component/Resources/RTS_Icon.png", UriKind.Absolute);
                this.Icon = new BitmapImage(iconUri);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning: Could not load window icon. {ex.Message}");
            }
        }

        /// <summary>
        /// Updates the main progress bar and the overall progress text (e.g., "Processing 5 of 100...").
        /// This also calculates the estimated time remaining based on the overall progress.
        /// </summary>
        /// <param name="current">The current item number being processed.</param>
        /// <param name="total">The total number of items to process.</param>
        public void UpdateProgress(int current, int total)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateProgress(current, total));
                return;
            }

            if (total <= 0) return;

            // Update progress bar and percentage
            double percentage = (double)current / total * 100;
            ProgressBar.Value = percentage;
            ProgressPercentageText.Text = $"{percentage:F0}%";
            OverallProgressText.Text = $"Processing element {current} of {total}...";

            // Calculate and display estimated time remaining
            string timeEstimate = "";
            if (current > 5 && _stopwatch.Elapsed.TotalSeconds > 1) // Start estimating after a few items
            {
                var elapsed = _stopwatch.Elapsed;
                var timePerItem = elapsed.TotalSeconds / current;
                var remainingItems = total - current;
                var remainingTime = TimeSpan.FromSeconds(timePerItem * remainingItems);
                timeEstimate = $" (est. {FormatTimeSpan(remainingTime)} remaining)";
            }

            // Append time estimate to the existing task description if it exists
            string currentTask = TaskDescriptionText.Text;
            int estIndex = currentTask.IndexOf(" (est.");
            if (estIndex >= 0)
            {
                currentTask = currentTask.Substring(0, estIndex);
            }
            TaskDescriptionText.Text = currentTask + timeEstimate;
        }

        /// <summary>
        /// Updates the detailed task description text (e.g., for processing a specific room).
        /// This is the equivalent of the old UpdateRoomStatus method.
        /// </summary>
        public void UpdateRoomStatus(string roomName, int currentRoom, int totalRooms)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateRoomStatus(roomName, currentRoom, totalRooms));
                return;
            }

            if (totalRooms > 0)
            {
                TaskDescriptionText.Text = $"Processing Room {currentRoom} of {totalRooms}: {roomName}";
            }
            else
            {
                TaskDescriptionText.Text = roomName;
            }
        }


        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalMinutes >= 1)
            {
                return $"{(int)ts.TotalMinutes} min {ts.Seconds} sec";
            }
            return $"{(int)ts.TotalSeconds} sec";
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancellationPending = true;
            CancelButton.IsEnabled = false;
            OverallProgressText.Text = "Cancelling operation...";
            TaskDescriptionText.Text = "Waiting for the current task to finish...";
        }
    }
}
