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
// - July 31, 2025: Added UpdateRoomStatus method for more detailed feedback.
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
        /// Updates the main progress bar and status text.
        /// </summary>
        public void UpdateProgress(int current, int total)
        {
            if (total == 0) return;

            ProgressBar.Value = (double)current / total * 100;
            StatusText.Text = $"Processing element {current} of {total}...";

            if (current > 5)
            {
                double elapsedSeconds = _stopwatch.Elapsed.TotalSeconds;
                double avgTimePerItem = elapsedSeconds / current;
                double remainingSeconds = (total - current) * avgTimePerItem;
                TimeEstimateText.Text = FormatTimeSpan(TimeSpan.FromSeconds(remainingSeconds));
            }

            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Updates the secondary status text for room processing.
        /// </summary>
        public void UpdateRoomStatus(string roomName, int currentRoom, int totalRooms)
        {
            if (totalRooms > 0)
            {
                RoomStatusText.Text = $"Processing Room {currentRoom} of {totalRooms}: {roomName}";
            }
            else
            {
                RoomStatusText.Text = roomName;
            }
            System.Windows.Forms.Application.DoEvents();
        }

        private string FormatTimeSpan(TimeSpan ts)
        {
            if (ts.TotalSeconds < 1) return "Less than a second remaining...";
            if (ts.TotalMinutes < 1) return $"About {ts.Seconds} seconds remaining...";
            if (ts.TotalHours < 1) return $"About {ts.Minutes} minute(s) and {ts.Seconds} second(s) remaining...";
            return $"About {ts.Hours} hour(s) and {ts.Minutes} minute(s) remaining...";
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _isCancellationPending = true;
            CancelButton.IsEnabled = false;
            StatusText.Text = "Cancelling operation...";
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
    }
}
