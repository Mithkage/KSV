//
// File: OpenLogFolderCommand.cs
//
// Namespace: RTS.Commands.Diagnostics
//
// Class: OpenLogFolderCommand
//
// Function: A simple external command that opens the RTS log directory in Windows Explorer.
//           This provides users with easy access to log files for troubleshooting.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - September 16, 2025: Initial creation of the command.
//

#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using RTS.Utilities; // Required for DiagnosticsManager
#endregion

namespace RTS.Commands.Diagnostics
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpenLogFolderCommand : IExternalCommand
    {
        /// <summary>
        /// Executes the command to open the log folder.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Use the diagnostics wrapper to log the operation and handle any potential errors.
                DiagnosticsManager.ExecuteWithDiagnostics("OpenLogFolder", () =>
                {
                    // Get the log directory path from the DiagnosticsManager and open it.
                    string logPath = DiagnosticsManager.LogDirectory;
                    Process.Start(logPath);
                });

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If something goes wrong (e.g., folder doesn't exist, permissions issue),
                // use the crash dialog to inform the user.
                DiagnosticsManager.ShowCrashDialog(ex, "Could not open log folder.");
                return Result.Failed;
            }
        }
    }
}
