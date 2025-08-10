//
// File: RTS_Reports.cs
//
// Namespace: RTS_Reports
//
// Class: RTS_ReportsClass
//
// Function: This Revit external command serves as the primary entry point for
//           accessing various reporting functionalities within the RTS_Reports add-in.
//           It launches a WPF window (ReportSelectionWindow) to allow the user
//           to select and generate different types of reports from data stored
//           in Revit's extensible storage.
//
// Author: Kyle Vorster
//
// Log:
// - July 2, 2025: Consolidated command execution to launch ReportSelectionWindow.
//                 Renamed file from RTS_ReportsClass.cs to RTS_Reports.cs.
// - July 15, 2025: Updated using directive for ReportSelectionWindow to RTS.UI.
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI; // UPDATED: Corrected using directive for ReportSelectionWindow
using System;
#endregion

namespace RTS.Commands.ProjectSetup
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_ReportsClass : IExternalCommand
    {
        /// <summary>
        /// The main entry point for the external command. Revit calls this method.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Create and show the WPF window
                ReportSelectionWindow window = new ReportSelectionWindow(commandData);
                window.ShowDialog(); // Use ShowDialog to make it modal

                // The result of the command will depend on how you want to handle
                // the window's closure. For simplicity, we assume success if no unhandled exception occurs.
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An unexpected error occurred: {ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}
