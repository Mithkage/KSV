using System;
using System.Windows;

namespace RTS.UI
{
    public partial class RoutingReportProgressBarWindow : Window
    {
        public bool IsCancelled { get; private set; }

        public RoutingReportProgressBarWindow()
        {
            InitializeComponent();
            IsCancelled = false;
        }

        /// <summary>
        /// Updates the progress bar and cable info fields.
        /// </summary>
        /// <param name="currentRoute">Current route number</param>
        /// <param name="totalRoutes">Total number of routes</param>
        /// <param name="cableReference">Cable Reference (unique identifier)</param>
        /// <param name="fromDevice">Origin device or location</param>
        /// <param name="toDevice">Destination device or location</param>
        /// <param name="percent">Progress percent</param>
        /// <param name="taskDescription">Task description</param>
        public void UpdateProgress(int currentRoute, int totalRoutes, string cableReference, string fromDevice, string toDevice, double percent, string taskDescription)
        {
            RouteProgressText.Text = cableReference;
            ProgressBar.Value = percent;
            TaskDescriptionText.Text = taskDescription;

            CableReferenceText.Text = cableReference;
            FromText.Text = fromDevice;
            ToText.Text = toDevice;
        }

        public void ShowError(string message)
        {
            ErrorText.Text = message;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            CancelButton.IsEnabled = false;
            TaskDescriptionText.Text = "Cancelling...";
        }
    }
}