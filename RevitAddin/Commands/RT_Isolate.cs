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
//                and "Clear Overrides" buttons. Modified Execute method to handle these actions.
//                Updated to use separate XAML and code-behind for InputWindow.
// - July 3, 2025: Added wildcard '*' functionality for search values.
//                A single '*' matches any non-empty parameter value.
//                '*' within a string acts as a multi-character wildcard (e.g., 'Cable-*' matches 'Cable-01', 'Cable-XYZ').
//                Matching is case-insensitive.
// - July 3, 2025: Removed duplicate 'WindowAction' enum definition to resolve CS0101 error.
//                The 'WindowAction' enum is now solely defined in 'InputWindow.xaml.cs'.
// - July 3, 2025: Fixed CS0103 error by correcting typo 'solidRedRedOverride' to 'solidRedOverride'.
// - July 4, 2025: Removed 'Clear Overrides' button from UI and corresponding logic.
//                Clearing overrides is now handled by submitting a blank input.
// - July 15, 2025: Updated 'using' statement for InputWindow to reflect new namespace RTS.UI.
// - August 8, 2025: Added reference to RTS.Utilities namespace for SharedParameters class.
// - August 8, 2025: Removed dialog box for processing scope and restricted operation to active view only.

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI; // UPDATED: Changed from 'RT_Isolate' to 'RTS.UI'
using RTS.Utilities; // Added to access SharedParameters class
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows; // Required for WPF Window, RoutedEventArgs
using System.Windows.Controls; // Required for WPF TextBox, Button, StackPanel, TextBlock
using System.Windows.Input; // Required for KeyBinding (though IsDefault/IsCancel handle this for buttons)
using System.Windows.Interop; // Required for HwndSource for WPF owner

namespace RTS.Commands
{
    // The 'WindowAction' enum definition is now in 'RTS.UI.InputWindow.xaml.cs'.

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
                    }
                    catch (Exception ex)
                    {
                        tResetGraphics.RollBack();
                        TaskDialog.Show("RT_Isolate Warning", $"Failed to reset view graphic overrides: {ex.Message}.");
                    }
                    tResetGraphics.Commit();
                }
            };

            string inputString = "";
            WindowAction chosenAction = WindowAction.Cancel; // Default action if dialog is closed without explicit button click

            try
            {
                InputWindow inputWindow = new InputWindow();
                // Pass the last used input to the window for display
                inputWindow.UserInput = _lastUserInput;

                // Set the Revit main window as the owner of the WPF dialog
                new WindowInteropHelper(inputWindow).Owner = commandData.Application.MainWindowHandle;

                // Show the WPF dialog
                bool? dialogResult = inputWindow.ShowDialog();

                // Determine the action based on DialogResult and the custom ActionChosen property
                if (dialogResult == true) // User clicked OK
                {
                    chosenAction = inputWindow.ActionChosen;
                    // If OK, inputString gets the user's input
                    inputString = inputWindow.UserInput;
                    _lastUserInput = inputString; // Store the current input for the next run
                }
                else // User clicked Cancel or closed the window (DialogResult is false or null)
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

            // Handle actions based on the chosen button or blank input
            if (chosenAction == WindowAction.Cancel)
            {
                // If cancelled, just exit the command. No graphic reset.
                return Result.Cancelled;
            }

            // If chosenAction is WindowAction.Ok, proceed with processing inputString
            if (string.IsNullOrWhiteSpace(inputString))
            {
                // If OK was clicked but the input was blank, reset graphics
                resetAllGraphics();
                TaskDialog.Show("RT_Isolate", "Input was blank. All graphic overrides have been cleared from the active view.");
                return Result.Succeeded;
            }

            // Always use elements from the active view only - no dialog to select scope
            ICollection<Element> elementsForProcessing = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Id != activeView.Id)
                .ToList();

            // Parse the input string into a list of search values
            List<string> searchValues = inputString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                  .Select(s => s.Trim())
                                                  .Where(s => !string.IsNullOrEmpty(s))
                                                  .ToList();

            if (!searchValues.Any())
            {
                // If after trimming and splitting, no valid search values remain,
                // also treat this as a signal to just reset graphics.
                resetAllGraphics();
                TaskDialog.Show("RT_Isolate", "No valid search values provided. All graphic overrides have been cleared from the active view.");
                return Result.Succeeded;
            }

            List<ElementId> targetedElements = new List<ElementId>();
            List<ElementId> nonTargetedElements = new List<ElementId>();

            // Define graphic overrides
            // For targeted elements: Solid Red
            OverrideGraphicSettings solidRedOverride = new OverrideGraphicSettings();
            solidRedOverride.SetProjectionLineColor(new Color(255, 0, 0)); // Red line color
            solidRedOverride.SetSurfaceForegroundPatternColor(new Color(255, 0, 0)); // Red fill color
            solidRedOverride.SetSurfaceForegroundPatternId(GetSolidFillPatternId(doc)); // Get solid fill pattern
            solidRedOverride.SetSurfaceTransparency(0); // Ensure not transparent (0 means opaque)

            // For non-targeted elements: 40% transparent, halftone
            OverrideGraphicSettings transparentHalftoneOverride = new OverrideGraphicSettings();
            transparentHalftoneOverride.SetHalftone(true);
            transparentHalftoneOverride.SetSurfaceTransparency(40); // 40% transparent

            int elementsFoundCount = 0; // Keep track of how many elements matched the criteria

            using (Transaction tApplyOverrides = new Transaction(doc, "Apply Isolation Overrides"))
            {
                tApplyOverrides.Start();
                try
                {
                    foreach (Element elem in elementsForProcessing)
                    {
                        bool matchFound = false;

                        // Check RTS_ID parameter
                        Parameter rtsIdParam = elem.get_Parameter(SharedParameters.Cable.RTS_ID_GUID);
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
                            foreach (Guid cableParamGuid in SharedParameters.Cable.RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    if (searchValues.Any(sv => MatchesWildcard(paramValue, sv)))
                                    {
                                        matchFound = true;
                                        break; // Found a match, no need to check other cable parameters for this element
                                    }
                                }
                            }
                        }

                        // Apply graphic overrides based on match status.
                        if (matchFound)
                        {
                            activeView.SetElementOverrides(elem.Id, solidRedOverride);
                            targetedElements.Add(elem.Id);
                            elementsFoundCount++;
                        }
                        else
                        {
                            // Only apply halftone/transparency if it's currently visible
                            // This prevents making previously hidden elements partially visible unexpectedly
                            if (!elem.IsHidden(activeView))
                            {
                                activeView.SetElementOverrides(elem.Id, transparentHalftoneOverride);
                                nonTargetedElements.Add(elem.Id);
                            }
                        }
                    }

                    tApplyOverrides.Commit();
                }
                catch (Exception ex)
                {
                    tApplyOverrides.RollBack();
                    message = $"Error during element graphic modification: {ex.Message}";
                    TaskDialog.Show("RT_Isolate Error", message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Helper method to get the ElementId of the solid fill pattern.
        /// </summary>
        private ElementId GetSolidFillPatternId(Document doc)
        {
            // Find the solid fill pattern element in the document.
            FillPatternElement solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            if (solidFill != null)
            {
                return solidFill.Id;
            }

            // This exception should ideally not be hit in a valid Revit document,
            // as the solid fill pattern is a standard built-in type.
            throw new Exception("Solid fill pattern not found in the document. This is unexpected.");
        }

        /// <summary>
        /// Checks if an input string matches a given pattern, supporting '*' as a wildcard.
        /// The '*' wildcard matches any sequence of zero or more characters.
        /// A pattern of a single '*' matches any non-empty input string.
        /// Matching is case-insensitive.
        /// </summary>
        /// <param name="input">The string to check (e.g., a parameter value).</param>
        /// <param name="pattern">The pattern to match against (e.g., user input with wildcards).</param>
        /// <returns>True if the input matches the pattern, false otherwise.</returns>
        private bool MatchesWildcard(string input, string pattern)
        {
            if (string.IsNullOrEmpty(input))
            {
                return false; // An empty input string cannot match any pattern (unless pattern is also empty, which is not the intent for wildcard functionality here).
            }

            // Special case: if the pattern is exactly "*", it matches any non-empty input.
            if (pattern.Equals("*", StringComparison.OrdinalIgnoreCase))
            {
                return !string.IsNullOrEmpty(input);
            }

            // Escape regex special characters in the pattern, then replace unescaped '*' with '.*'
            // Regex.Escape escapes most special characters, but we want '*' to be a wildcard,
            // so we escape the pattern first, then replace the escaped '*' ('\*') with '.*'.
            string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";

            // Create a Regex object for case-insensitive matching
            Regex regex = new Regex(regexPattern, RegexOptions.IgnoreCase);

            return regex.IsMatch(input);
        }
    }
}
