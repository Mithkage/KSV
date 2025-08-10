//
// --- FILE: LinkManager.cs ---
//
// File: LinkManager.cs
//
// Namespace: RTS.Commands
//
// Class: LinkManagerCommand
//
// Function: This Revit external command provides a WPF interface for managing Revit links.
//           It allows a user to view all linked models, assign a discipline, set the
//           coordinate system, and define a responsible person with contact details.
//           This data is intended to be saved as a "Project Profile" within the Revit
//           project's Extensible Storage for use by other project setup and auditing tools.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 17, 2025: Updated to correctly set the Revit window as the owner for the WPF dialog.
// - July 16, 2025: Initial creation of the command to launch the Link Manager WPF window.
//

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI; // The namespace for your WPF window
using System;
using System.Windows.Interop; // Required for WindowInteropHelper

namespace RTS.Commands.DocumentTools.ProjectManagement
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LinkManagerCommand : IExternalCommand
    {
        /// <summary>
        /// The main entry point for the Revit external command.
        /// This method is executed when the user clicks the button in the Revit UI.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, Revit application, and documents.</param>
        /// <param name="message">A message string which can be set by the external application 
        /// to report failure or success of the command.</param>
        /// <param name="elements">A set of elements which can be passed back by the external 
        /// application to highlight in the Revit UI.</param>
        /// <returns>The result of the command execution.</returns>
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the application and document from the command data
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Instantiate the LinkManagerWindow, passing the active document to its constructor.
                // This window is defined in the RTS.UI namespace.
                LinkManagerWindow linkManagerWindow = new LinkManagerWindow(doc);

                // *** UPDATE: Set the owner of the WPF window to be the main Revit window. ***
                // This is the correct way to launch a modal WPF dialog in Revit. It ensures
                // the window stays on top of Revit and behaves correctly.
                WindowInteropHelper wih = new WindowInteropHelper(linkManagerWindow)
                {
                    Owner = uiApp.MainWindowHandle
                };

                // Show the window as a modal dialog. This means the user cannot interact with
                // Revit until the window is closed.
                linkManagerWindow.ShowDialog();

                // The logic for saving data will be handled within the window's code-behind.
                // The main command's job is simply to launch the UI.

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If any error occurs, report it to the user.
                message = $"An unexpected error occurred: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}