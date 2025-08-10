//
// --- FILE: FileDialogs.cs ---
//
// Description:
// A utility class providing standardized file and folder selection dialogs
// that work well within the Revit environment.
//
// Change Log:
// - August 16, 2025: Initial creation.
// - August 21, 2025: Added PromptForSaveFile method for ReportGeneratorBase.
// - August 22, 2025: Added PromptForCsvFile method for PC_Cable_ImporterClass.
//

using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace RTS.Utilities
{
    /// <summary>
    /// Provides standardized file and folder selection dialogs.
    /// </summary>
    public static class FileDialogs
    {
        /// <summary>
        /// Prompts the user to select a folder with a modern dialog that supports path pasting.
        /// </summary>
        /// <param name="description">The dialog description shown to the user.</param>
        /// <param name="initialPath">Optional initial path to display in the dialog.</param>
        /// <returns>The selected folder path or null if cancelled.</returns>
        public static string PromptForFolder(string description, string initialPath = null)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = description;
                dialog.ShowNewFolderButton = true;

                // Use the initial path if provided, otherwise default to My Documents
                if (!string.IsNullOrEmpty(initialPath))
                {
                    dialog.SelectedPath = initialPath;
                }
                else
                {
                    dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                DialogResult result = dialog.ShowDialog();
                return result == DialogResult.OK ? dialog.SelectedPath : null;
            }
        }

        /// <summary>
        /// Prompts the user to select a file with a modern dialog that supports path pasting.
        /// </summary>
        /// <param name="title">The dialog title shown to the user.</param>
        /// <param name="filter">The file filter (e.g., "CSV Files (*.csv)|*.csv|All files (*.*)|*.*").</param>
        /// <param name="initialDirectory">Optional initial directory to display in the dialog.</param>
        /// <returns>The selected file path or null if cancelled.</returns>
        public static string PromptForFile(string title, string filter, string initialDirectory = null)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = title;
                dialog.Filter = filter;
                dialog.RestoreDirectory = true;

                if (!string.IsNullOrEmpty(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                }

                DialogResult result = dialog.ShowDialog();
                return result == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Prompts the user to select a CSV file with a modern dialog that supports path pasting.
        /// </summary>
        /// <param name="title">The dialog title shown to the user.</param>
        /// <param name="initialDirectory">Optional initial directory to display in the dialog.</param>
        /// <returns>The selected CSV file path or null if cancelled.</returns>
        public static string PromptForCsvFile(string title, string initialDirectory = null)
        {
            return PromptForFile(
                title,
                "CSV Files (*.csv)|*.csv|All files (*.*)|*.*",
                initialDirectory
            );
        }

        /// <summary>
        /// Prompts the user for a location to save a file.
        /// </summary>
        /// <param name="title">The dialog title shown to the user.</param>
        /// <param name="filter">The file filter (e.g., "CSV Files (*.csv)|*.csv|All files (*.*)|*.*").</param>
        /// <param name="defaultFileName">Default filename to suggest.</param>
        /// <param name="initialDirectory">Optional initial directory to display in the dialog.</param>
        /// <returns>The selected save path or null if cancelled.</returns>
        public static string PromptForSaveLocation(string title, string filter, string defaultFileName = null, string initialDirectory = null)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = title;
                dialog.Filter = filter;
                dialog.RestoreDirectory = true;
                dialog.OverwritePrompt = true;

                if (!string.IsNullOrEmpty(defaultFileName))
                {
                    dialog.FileName = defaultFileName;
                }

                if (!string.IsNullOrEmpty(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                }

                DialogResult result = dialog.ShowDialog();
                return result == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Prompts the user for a save file location with the specified parameters.
        /// </summary>
        /// <param name="defaultFileName">Default filename to suggest</param>
        /// <param name="title">Dialog title</param>
        /// <param name="filter">File filter string</param>
        /// <param name="initialDirectory">Initial directory to show</param>
        /// <param name="defaultExt">Default file extension (without dot)</param>
        /// <returns>The selected save file path or null if cancelled</returns>
        public static string PromptForSaveFile(
            string defaultFileName,
            string title,
            string filter,
            string initialDirectory = null,
            string defaultExt = null)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = title;
                dialog.Filter = filter;
                dialog.RestoreDirectory = true;
                dialog.OverwritePrompt = true;

                if (!string.IsNullOrEmpty(defaultFileName))
                {
                    dialog.FileName = defaultFileName;
                }

                if (!string.IsNullOrEmpty(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                }

                if (!string.IsNullOrEmpty(defaultExt))
                {
                    dialog.DefaultExt = defaultExt;
                    dialog.AddExtension = true;
                }

                DialogResult result = dialog.ShowDialog();
                return result == DialogResult.OK ? dialog.FileName : null;
            }
        }
    }
}