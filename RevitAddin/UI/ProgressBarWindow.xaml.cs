//
// --- FILE: ProgressBarWindow.xaml.cs ---
//
// Description:
// Code-behind for the ProgressBarWindow.xaml. It provides methods to update
// the progress and status messages, and handles the cancellation logic.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Change Log:
// - July 30, 2025: Initial creation of the code-behind.
//

#region Namespaces
using System;
using System.Diagnostics;
using System.Windows;
#endregion

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for ProgressBarWindow.xaml
    /// </summary>
    public partial class ProgressBarWindow : Window
    {
        private Stopwatch _stopwatch = new Stopwatch();
        private bool _isCancellationPending = false;

        public bool IsCancellationPending => _isCancellationPending;

        public ProgressBarWindow()
        {
            InitializeComponent();
            _stopwatch.Start();
        }

        /// <summary>
        /// Updates the progress bar and status text.
        /// </summary>
        /// <param name="current">The current item number being processed.</param>
        /// <param name="total">The total number of items.</param>
        public void UpdateProgress(int current, int total)
        {
            if (total == 0) return;

            // Update the progress bar value
            ProgressBar.Value = (double)current / total * 100;

            // Update the status text
            StatusText.Text = $"Processing element {current} of {total}...";

            // Update the time estimate (only after a few items to get a stable average)
            if (current > 5)
            {
                double elapsedSeconds = _stopwatch.Elapsed.TotalSeconds;
                double avgTimePerItem = elapsedSeconds / current;
                double remainingSeconds = (total - current) * avgTimePerItem;

                TimeEstimateText.Text = FormatTimeSpan(TimeSpan.FromSeconds(remainingSeconds));
            }

            // Allow the UI to update
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Formats a TimeSpan into a user-friendly string.
        /// </summary>
        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalSeconds < 1)
            {
                return "Less than a second remaining...";
            }
            if (ts.TotalMinutes < 1)
            {
                return $"About {ts.Seconds} seconds remaining...";
            }
            if (ts.TotalHours < 1)
            {
                return $"About {ts.Minutes} minute(s) and {ts.Seconds} second(s) remaining...";
            }
            return $"About {ts.Hours} hour(s) and {ts.Minutes} minute(s) remaining...";
        }

        /// <summary>
        /// Handles the click event for the Cancel button.
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _isCancellationPending = true;
            CancelButton.IsEnabled = false;
            StatusText.Text = "Cancelling operation...";
        }
    }
}
