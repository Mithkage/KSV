//
// File: SharedParameterCommand.cs
//
// Namespace: RTS.Commands.ProjectSetup
//
// Class: SharedParameterCommand
//
// Function: This Revit external command allows a user to generate a .txt shared parameter
//           file from the master definitions stored in SharedParameterFile.cs. It provides
//           a modern save dialog and post-creation options.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - Initial creation of the command.
// - Implemented a SaveFileDialog to allow custom file naming and location.
// - Added post-creation options to set the file as active and open its location.
//
// Author: ReTick Solutions
//
#region Namespaces
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Utilities; // Required for SharedParameters
#endregion

namespace RTS.Commands.ProjectSetup
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SharedParameterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            // --- 1. Prompt user for save location and file name ---
            string filePath;
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.Title = "Save RTS Shared Parameter File";
                saveFileDialog.FileName = "RTS_Shared_Parameters.txt"; // Seed with the default name
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return Result.Cancelled;
                }
                filePath = saveFileDialog.FileName;
            }

            // --- 2. Generate the shared parameter file ---
            try
            {
                // Call the public method from the SharedParameters utility class
                SharedParameters.CreateSharedParameterFile(filePath);
            }
            catch (Exception ex)
            {
                message = $"Failed to generate the shared parameter file: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // --- 3. Show post-generation options to the user ---
            TaskDialog taskDialog = new TaskDialog("Success");
            taskDialog.MainInstruction = "Shared Parameter File Generated Successfully";
            taskDialog.MainContent = $"The file has been saved to:\n{filePath}";

            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Set as Active Shared Parameter File", "Sets the newly created file as the active shared parameter file for this Revit session.");
            taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Open File Location", "Opens the folder containing the new file in Windows Explorer.");

            taskDialog.CommonButtons = TaskDialogCommonButtons.Close;
            taskDialog.DefaultButton = TaskDialogResult.Close;

            TaskDialogResult result = taskDialog.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                try
                {
                    app.SharedParametersFilename = filePath;
                    TaskDialog.Show("Active File Set", "The new shared parameter file has been set as active.");
                }
                catch (Exception ex)
                {
                    message = $"Could not set the active shared parameter file: {ex.Message}";
                    TaskDialog.Show("Error", message);
                }
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                try
                {
                    Process.Start(Path.GetDirectoryName(filePath));
                }
                catch (Exception ex)
                {
                    message = $"Could not open the file location: {ex.Message}";
                    TaskDialog.Show("Error", message);
                }
            }

            return Result.Succeeded;
        }
    }
}
