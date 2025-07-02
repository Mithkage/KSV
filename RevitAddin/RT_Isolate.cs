// File: RT_Isolate.cs

// Application: Revit 2022/2024 External Command

// Description: This script allows the user to input one or more comma-separated string values.
//              It first removes any existing graphic overrides from all elements in the active view.
//              Then, it searches the active view for elements that have a matching value
//              in their 'RTS_ID' shared parameter or any of the 'RTS_Cable_XX' shared parameters.
//              All elements found with a matching value will have their graphics overridden to solid red.
//              All other visible elements in the active view will be overridden to be 40% transparent and half-tone.
//              Elements hidden before running the command will remain hidden unless specifically targeted and made visible by Revit's graphic rules.

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Interop; // Required for HwndSource for WPF owner

namespace RT_Isolate
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_IsolateClass : IExternalCommand
    {
        // Define GUIDs for the shared parameters, copied from RT_TrayConduits.cs
        private static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
        private static readonly List<Guid> RTS_Cable_GUIDs = new List<Guid>
        {
            new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), // RTS_Cable_01
            new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), // RTS_Cable_02
            new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"), // RTS_Cable_03
            new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), // RTS_Cable_04
            new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"), // RTS_Cable_05
            new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), // RTS_Cable_06
            new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"), // RTS_Cable_07
            new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), // RTS_Cable_08
            new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"), // RTS_Cable_09
            new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), // RTS_Cable_10
            new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"), // RTS_Cable_11
            new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), // RTS_Cable_12
            new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"), // RTS_Cable_13
            new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), // RTS_Cable_14
            new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"), // RTS_Cable_15
            new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), // RTS_Cable_16
            new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"), // RTS_Cable_17
            new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), // RTS_Cable_18
            new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"), // RTS_Cable_19
            new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), // RTS_Cable_20
            new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"), // RTS_Cable_21
            new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), // RTS_Cable_22
            new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"), // RTS_Cable_23
            new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), // RTS_Cable_24
            new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"), // RTS_Cable_25
            new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), // RTS_Cable_26
            new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"), // RTS_Cable_27
            new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), // RTS_Cable_28
            new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"), // RTS_Cable_29
            new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b")  // RTS_Cable_30
        };

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

            // --- Reset graphic overrides for all elements at the start ---
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
                        // Removed CanOverrideGraphicsInView check.
                        activeView.SetElementOverrides(elem.Id, new OverrideGraphicSettings());
                    }
                }
                catch (Exception ex)
                {
                    tResetGraphics.RollBack();
                    // Keep this warning as it indicates an issue but allows the script to continue.
                    TaskDialog.Show("RT_Isolate Warning", $"Failed to reset view graphic overrides: {ex.Message}. Attempting to proceed with isolation.");
                }
                tResetGraphics.Commit();
            }
            // --- END Reset Graphics ---


            // Create and show the WPF input window
            string inputString = "";
            try
            {
                InputWindow inputWindow = new InputWindow();

                // Set the Revit main window as the owner of the WPF window
                // This ensures the WPF window stays on top of Revit
                new WindowInteropHelper(inputWindow).Owner = commandData.Application.MainWindowHandle;

                bool? dialogResult = inputWindow.ShowDialog();

                if (dialogResult == true)
                {
                    inputString = inputWindow.UserInput;
                }
                else
                {
                    // User clicked Cancel or closed the WPF window
                    TaskDialog.Show("RT_Isolate", "Input cancelled by user. Script cancelled.");
                    return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                message = $"Error displaying input window: {ex.Message}";
                TaskDialog.Show("RT_Isolate Error", message);
                return Result.Failed;
            }

            if (string.IsNullOrWhiteSpace(inputString))
            {
                TaskDialog.Show("RT_Isolate", "No values entered. Script cancelled.");
                return Result.Cancelled;
            }

            // Parse the input string into a list of search values
            List<string> searchValues = inputString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                   .Select(s => s.Trim())
                                                   .Where(s => !string.IsNullOrEmpty(s))
                                                   .ToList();

            if (!searchValues.Any())
            {
                TaskDialog.Show("RT_Isolate", "No valid search values parsed from input. Script cancelled.");
                return Result.Cancelled;
            }

            // Collect all elements in the active view for processing (excluding element types and the view itself)
            var allElementsInViewForProcessing = new FilteredElementCollector(doc, activeView.Id)
                                                       .WhereElementIsNotElementType()
                                                       .Where(e => e.Id != activeView.Id) // Exclude the view element itself
                                                       .ToList();

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
                    foreach (Element elem in allElementsInViewForProcessing)
                    {
                        bool matchFound = false;

                        // Check RTS_ID parameter
                        Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                        if (rtsIdParam != null && !string.IsNullOrEmpty(rtsIdParam.AsString()))
                        {
                            string paramValue = rtsIdParam.AsString();
                            if (searchValues.Any(sv => string.Equals(sv, paramValue, StringComparison.OrdinalIgnoreCase)))
                            {
                                matchFound = true;
                            }
                        }

                        // If no match yet, check RTS_Cable_XX parameters
                        if (!matchFound)
                        {
                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    if (searchValues.Any(sv => string.Equals(sv, paramValue, StringComparison.OrdinalIgnoreCase)))
                                    {
                                        matchFound = true;
                                        break; // Found a match, no need to check other cable parameters for this element
                                    }
                                }
                            }
                        }

                        // Apply graphic overrides based on match status.
                        // The CanOverrideGraphicsInView check is removed here, assuming all relevant elements are overrideable.
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

                    // Removed the final success TaskDialog.Show message.
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
    }
}