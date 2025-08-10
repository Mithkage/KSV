//
// File: SetLogLevelCommand.cs
//
// Namespace: RTS.Commands.Diagnostics
//
// Class: SetLogLevelCommand
//
// Function: An external command that allows the user to dynamically change the
//           logging verbosity of the DiagnosticsManager via a TaskDialog.
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
using RTS.Utilities; // Required for DiagnosticsManager
#endregion

namespace RTS.Commands.Diagnostics
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SetLogLevelCommand : IExternalCommand
    {
        /// <summary>
        /// Executes the command to set the log level.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Use the diagnostics wrapper to log the operation.
                DiagnosticsManager.ExecuteWithDiagnostics("SetLogLevel", () =>
                {
                    TaskDialog dialog = new TaskDialog("Set RTS Log Level")
                    {
                        MainInstruction = "Select the desired logging level.",
                        MainContent = "More detailed levels can help troubleshoot issues but create larger log files.\n\n" +
                                      $"Current Level: {DiagnosticsManager.CurrentLogLevel}",
                        CommonButtons = TaskDialogCommonButtons.Cancel
                    };

                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Debug", "Logs everything, including performance data and verbose operational details.");
                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Info", "Logs standard operations and startup/shutdown. (Default)");
                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Warning", "Logs only warnings, errors, and crashes.");
                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Error", "Logs only critical errors and crashes.");

                    TaskDialogResult result = dialog.Show();

                    DiagnosticsManager.LogLevel newLevel = DiagnosticsManager.CurrentLogLevel;

                    if (result == TaskDialogResult.CommandLink1) newLevel = DiagnosticsManager.LogLevel.Debug;
                    else if (result == TaskDialogResult.CommandLink2) newLevel = DiagnosticsManager.LogLevel.Info;
                    else if (result == TaskDialogResult.CommandLink3) newLevel = DiagnosticsManager.LogLevel.Warning;
                    else if (result == TaskDialogResult.CommandLink4) newLevel = DiagnosticsManager.LogLevel.Error;

                    // Only update and notify if the level has actually changed.
                    if (newLevel != DiagnosticsManager.CurrentLogLevel)
                    {
                        DiagnosticsManager.CurrentLogLevel = newLevel;
                        // Log the change itself for future reference.
                        DiagnosticsManager.LogMessage(DiagnosticsManager.LogLevel.Info, $"Log level changed to {newLevel}");
                        TaskDialog.Show("Log Level Updated", $"The RTS log level has been set to {newLevel}.");
                    }
                });

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DiagnosticsManager.ShowCrashDialog(ex, "Failed to set log level.");
                return Result.Failed;
            }
        }
    }
}
