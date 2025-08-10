//
// File: CheckResourcesCommand.cs
//
// Namespace: RTS.Commands.Diagnostics
//
// Class: CheckResourcesCommand
//
// Function: An external command that informs the user about the automatic resource
//           verification process that runs at startup.
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
    public class CheckResourcesCommand : IExternalCommand
    {
        /// <summary>
        /// Executes the command to check resources.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                DiagnosticsManager.ExecuteWithDiagnostics("CheckResources", () =>
                {
                    // This command simply informs the user that the check is automatic
                    // and points them to the log file for the results. This avoids
                    // duplicating the resource list from the main App class.
                    TaskDialog.Show("Resource Check Information",
                        "A check for all required image resources is performed automatically every time Revit starts.\n\n" +
                        "Please check the latest log file for a 'VerifyEmbeddedResources' entry to see the results.\n\n" +
                        "You can open the log folder directly from this same ribbon panel.");
                });

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DiagnosticsManager.ShowCrashDialog(ex, "Failed to show resource check information.");
                return Result.Failed;
            }
        }
    }
}
