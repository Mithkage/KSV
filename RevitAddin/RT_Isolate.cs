// File: RT_Isolate.cs

// Application: Revit 2022 External Command

// Description: This script allows the user to input one or more comma-separated string values.
//              It then searches the active view for elements that have a matching value
//              in their 'RTS_ID' shared parameter or any of the 'RTS_Cable_XX' shared parameters.
//              All elements found with a matching value will remain visible, while
//              all other visible elements in the active view will be hidden.

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Interop; // Required for HwndSource for WPF owner
// No longer needed: using Autodesk.Revit.UI.Selection; // PromptStringOptions is not available in Revit 2022
// No longer needed: using System.Windows.Forms; // Replaced with WPF

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

            // Collect all visible elements in the active view
            // We'll exclude element types and the view itself
            var allVisibleElementsInView = new FilteredElementCollector(doc, activeView.Id)
                                            .WhereElementIsNotElementType()
                                            .Where(e => e.Id != activeView.Id) // Exclude the view element itself
                                            .ToList();

            List<ElementId> elementsToHide = new List<ElementId>();
            int elementsFoundCount = 0;

            using (Transaction t = new Transaction(doc, "Hide Non-Matching Elements"))
            {
                t.Start();
                try
                {
                    foreach (Element elem in allVisibleElementsInView)
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
                                elementsFoundCount++;
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
                                        elementsFoundCount++;
                                        break; // Found a match, no need to check other cable parameters for this element
                                    }
                                }
                            }
                        }

                        // If no match was found for this element, add it to the list to hide
                        if (!matchFound)
                        {
                            elementsToHide.Add(elem.Id);
                        }
                    }

                    if (elementsToHide.Any())
                    {
                        // Hide the non-matching elements using the HideElements method
                        activeView.HideElements(elementsToHide);
                        TaskDialog.Show("RT_Isolate", $"Successfully hid {elementsToHide.Count} non-matching elements. {elementsFoundCount} matching elements remain visible.");
                    }
                    else if (elementsFoundCount > 0)
                    {
                        TaskDialog.Show("RT_Isolate", $"All {elementsFoundCount} found elements are already visible. No elements were hidden.");
                    }
                    else
                    {
                        TaskDialog.Show("RT_Isolate", "No elements found matching the provided values. No elements were hidden.");
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = $"Error during element visibility modification: {ex.Message}";
                    TaskDialog.Show("RT_Isolate Error", message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
    }
}