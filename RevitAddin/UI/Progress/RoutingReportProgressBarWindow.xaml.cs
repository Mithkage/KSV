//-----------------------------------------------------------------------------
// <copyright file="RoutingReportProgressBarWindow.xaml.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Progress bar window for routing report generation
// </summary>
//-----------------------------------------------------------------------------

using System.Windows;

namespace RTS.UI
{
    /// <summary>
    /// Interaction logic for RoutingReportProgressBarWindow.xaml
    /// </summary>
    public partial class RoutingReportProgressBarWindow : Window
    {
        public bool IsCancelled { get; private set; } = false;

        public RoutingReportProgressBarWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Updates the progress bar and text fields
        /// </summary>
        public void UpdateProgress(int current, int total, string cableRef, string from, string to, double progressPercent, string taskDescription)
        {
            RouteProgressText.Text = $"{current} of {total}";
            CableReferenceText.Text = cableRef;
            FromText.Text = from;
            ToText.Text = to;

            ProgressBar.Value = progressPercent;

            TaskDescriptionText.Text = taskDescription;
        }

        /// <summary>
        /// Shows an error message in the progress window
        /// </summary>
        public void ShowError(string errorMessage)
        {
            ErrorText.Text = errorMessage;
            TaskDescriptionText.Text = "Error encountered";
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            CancelButton.IsEnabled = false;
            TaskDescriptionText.Text = "Cancelling operation...";
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}