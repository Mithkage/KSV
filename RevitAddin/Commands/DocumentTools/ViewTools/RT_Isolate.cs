// File: RT_Isolate.cs

// Application: Revit 2022/2024 External Command

// Description: This script allows the user to input one or more comma-separated string values.
//              If the input is blank/empty or cancelled, it removes all existing graphic overrides from the active view.
//              Otherwise, it first removes any existing graphic overrides from all elements in the active view.
//              Then, it searches the active view for elements that have a matching value
//              in their 'RTS_ID' shared parameter or any of the 'RTS_Cable_XX' shared parameters.
//              All elements found with a matching value will have their graphics overridden to solid red.
//              All other visible elements in the active view will be overridden to be 40% transparent and half-tone.
//              Elements hidden before running the command will remain hidden unless specifically targeted and made visible by Revit's graphic rules.

// Log:
// - July 3, 2025: Implemented WPF InputWindow with "OK" (default/Return key), "Cancel" (closes window),
//                 and "Clear Overrides" buttons. Modified Execute method to handle these actions.
//                 Updated to use separate XAML and code-behind for InputWindow.
// - July 3, 2025: Added wildcard '*' functionality for search values.
//                 A single '*' matches any non-empty parameter value.
//                 '*' within a string acts as a multi-character wildcard (e.g., 'Cable-*' matches 'Cable-01', 'Cable-XYZ').
//                 Matching is case-insensitive.
// - July 3, 2025: Removed duplicate 'WindowAction' enum definition to resolve CS0101 error.
//                 The 'WindowAction' enum is now solely defined in 'InputWindow.xaml.cs'.
// - July 3, 2025: Fixed CS0103 error by correcting typo 'solidRedRedOverride' to 'solidRedOverride'.
// - July 4, 2025: Removed 'Clear Overrides' button from UI and corresponding logic.
//                 Clearing overrides is now handled by submitting a blank input.
// - July 15, 2025: Updated 'using' statement for InputWindow to reflect new namespace RTS.UI.
// - August 8, 2025: Added reference to RTS.Utilities namespace for SharedParameters class.
// - August 8, 2025: Removed dialog box for processing scope and restricted operation to active view only.
// - [APPLIED FIX]: Updated parameter access to use the new SharedParameters structure.

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;

namespace RTS.Commands.DocumentTools.ViewTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_IsolateClass : IExternalCommand
    {
        // Static variable to store the last used input within the session
        private static string _lastUserInput = string.Empty; // Initialize with empty string

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            if (activeView == null)
            {
                message = "No active view found. Please open a view.";
                TaskDialog.Show("RT_Isolate Error", message);
                return Result.Failed;
            }

            // Action to reset all graphic overrides in the active view
            Action resetAllGraphics = () =>
            {
                using (Transaction tResetGraphics = new Transaction(doc, "Reset View Graphic Overrides"))
                {
                    tResetGraphics.Start();
                    try
                    {
                        var allElementsInView = new FilteredElementCollector(doc, activeView.Id)
                            .WhereElementIsNotElementType()
                            .ToList();

                        foreach (Element elem in allElementsInView)
                        {
                            // Reset overrides to default for each element
                            activeView.SetElementOverrides(elem.Id, new OverrideGraphicSettings());
                        }
                        tResetGraphics.Commit();
                    }
                    catch (Exception ex)
                    {
                        if (tResetGraphics.HasStarted()) tResetGraphics.RollBack();
                        TaskDialog.Show("RT_Isolate Warning", $"Failed to reset view graphic overrides: {ex.Message}.");
                    }
                }
            };

            string inputString = "";
            WindowAction chosenAction = WindowAction.Cancel; // Default action if dialog is closed without explicit button click

            try
            {
                InputWindow inputWindow = new InputWindow();
                inputWindow.UserInput = _lastUserInput;
                new WindowInteropHelper(inputWindow).Owner = commandData.Application.MainWindowHandle;
                bool? dialogResult = inputWindow.ShowDialog();

                if (dialogResult == true) // User clicked OK
                {
                    chosenAction = inputWindow.ActionChosen;
                    inputString = inputWindow.UserInput;
                    _lastUserInput = inputString; // Store the current input for the next run
                }
                else // User clicked Cancel or closed the window
                {
                    chosenAction = WindowAction.Cancel;
                }
            }
            catch (Exception ex)
            {
                message = $"Error displaying input window: {ex.Message}";
                TaskDialog.Show("RT_Isolate Error", message);
                return Result.Failed;
            }

            if (chosenAction == WindowAction.Cancel)
            {
                return Result.Cancelled;
            }

            if (string.IsNullOrWhiteSpace(inputString))
            {
                resetAllGraphics();
                TaskDialog.Show("RT_Isolate", "Input was blank. All graphic overrides have been cleared from the active view.");
                return Result.Succeeded;
            }

            ICollection<Element> elementsForProcessing = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Id != activeView.Id)
                .ToList();

            List<string> searchValues = inputString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(s => s.Trim())
                                                    .Where(s => !string.IsNullOrEmpty(s))
                                                    .ToList();

            if (!searchValues.Any())
            {
                resetAllGraphics();
                TaskDialog.Show("RT_Isolate", "No valid search values provided. All graphic overrides have been cleared from the active view.");
                return Result.Succeeded;
            }

            // Define graphic overrides
            OverrideGraphicSettings solidRedOverride = new OverrideGraphicSettings();
            solidRedOverride.SetProjectionLineColor(new Color(255, 0, 0));
            solidRedOverride.SetSurfaceForegroundPatternColor(new Color(255, 0, 0));
            solidRedOverride.SetSurfaceForegroundPatternId(GetSolidFillPatternId(doc));
            solidRedOverride.SetSurfaceTransparency(0);

            OverrideGraphicSettings transparentHalftoneOverride = new OverrideGraphicSettings();
            transparentHalftoneOverride.SetHalftone(true);
            transparentHalftoneOverride.SetSurfaceTransparency(40);

            using (Transaction tApplyOverrides = new Transaction(doc, "Apply Isolation Overrides"))
            {
                tApplyOverrides.Start();
                try
                {
                    // First, reset all graphics to ensure a clean slate
                    foreach (Element elem in elementsForProcessing)
                    {
                        activeView.SetElementOverrides(elem.Id, new OverrideGraphicSettings());
                    }

                    // Now, apply the new overrides
                    foreach (Element elem in elementsForProcessing)
                    {
                        bool matchFound = false;

                        // Check RTS_ID parameter using the new compatibility layer
                        Parameter rtsIdParam = elem.get_Parameter(SharedParameters.General.RTS_ID);
                        if (rtsIdParam != null && !string.IsNullOrEmpty(rtsIdParam.AsString()))
                        {
                            string paramValue = rtsIdParam.AsString();
                            if (searchValues.Any(sv => MatchesWildcard(paramValue, sv)))
                            {
                                matchFound = true;
                            }
                        }

                        // If no match yet, check RTS_Cable_XX parameters
                        if (!matchFound)
                        {
                            // Use the new dynamic list of cable GUIDs from the compatibility layer
                            foreach (Guid cableParamGuid in SharedParameters.Cable.AllCableGuids)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    if (searchValues.Any(sv => MatchesWildcard(paramValue, sv)))
                                    {
                                        matchFound = true;
                                        break;
                                    }
                                }
                            }
                        }

                        // Apply graphic overrides based on match status.
                        if (matchFound)
                        {
                            activeView.SetElementOverrides(elem.Id, solidRedOverride);
                        }
                        else
                        {
                            if (!elem.IsHidden(activeView))
                            {
                                activeView.SetElementOverrides(elem.Id, transparentHalftoneOverride);
                            }
                        }
                    }

                    tApplyOverrides.Commit();
                }
                catch (Exception ex)
                {
                    if (tApplyOverrides.HasStarted()) tApplyOverrides.RollBack();
                    message = $"Error during element graphic modification: {ex.Message}";
                    TaskDialog.Show("RT_Isolate Error", message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        private ElementId GetSolidFillPatternId(Document doc)
        {
            FillPatternElement solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            if (solidFill != null)
            {
                return solidFill.Id;
            }
            throw new Exception("Solid fill pattern not found in the document.");
        }

        private bool MatchesWildcard(string input, string pattern)
        {
            if (string.IsNullOrEmpty(input))
            {
                return false;
            }

            if (pattern.Equals("*", StringComparison.OrdinalIgnoreCase))
            {
                return !string.IsNullOrEmpty(input);
            }

            string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
            return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase);
        }
    }
}
