//
// File: ViewportAlignmentTool.cs
//
// Namespace: RTS.Commands
//
// Class: ViewportAlignmentTool
//
// Function: This Revit external command allows the user to select a reference viewport on a sheet
//           and then select multiple target viewports on the same sheet. It aligns the center
//           of all selected target viewports to the center of the reference viewport.
//           The script automatically unpins target viewports before alignment and re-pins them
//           after alignment if they were originally pinned. A summary message is provided at the
//           end, detailing which viewports were moved and which were already correctly aligned.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 16, 2025: Initial creation with reference/target selection, center alignment, and transaction handling.
// - July 16, 2025: Added logic to unpin target viewports before alignment and re-pin them afterwards.
// - July 16, 2025: Implemented detailed user message at the end, listing updated viewports and their status (moved/already aligned).
// - July 16, 2025: Encapsulated the command and helper class within the 'RTS.Commands' namespace.
// - July 16, 2025: Corrected usage of ViewName. Viewports do not have a ViewName property;
//                  instead, retrieve the associated View element and use its Name property.
//
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; // Required for StringBuilder

namespace RTS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewportAlignmentTool : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Viewport referenceViewport = null;
            XYZ referenceCenter = null;
            string referenceViewName = string.Empty; // To store the reference view's name

            // 1. Reference Viewport Selection
            try
            {
                Reference refSelection = uidoc.Selection.PickObject(ObjectType.Element, new ViewportSelectionFilter(), "Select the reference viewport for alignment.");
                referenceViewport = doc.GetElement(refSelection) as Viewport;

                if (referenceViewport == null)
                {
                    message = "No reference viewport selected. Exiting.";
                    return Result.Cancelled;
                }

                // Get the View associated with the reference viewport to get its name
                View referenceView = doc.GetElement(referenceViewport.ViewId) as View;
                if (referenceView != null)
                {
                    referenceViewName = referenceView.Name;
                }
                else
                {
                    referenceViewName = "Unknown View"; // Fallback if view name can't be retrieved
                }


                // Ensure the reference viewport is on an active sheet
                ViewSheet sheet = doc.GetElement(referenceViewport.SheetId) as ViewSheet;
                if (sheet == null || sheet.Id != doc.ActiveView.Id)
                {
                    TaskDialog.Show("Error", "The selected reference viewport is not on the currently active sheet. Please activate the sheet containing the reference viewport.");
                    return Result.Cancelled;
                }

                // Get its precise position
                referenceCenter = referenceViewport.GetBoxCenter();
            }
            catch (OperationCanceledException)
            {
                message = "Reference viewport selection cancelled.";
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = "Error selecting reference viewport: " + ex.Message;
                return Result.Failed;
            }

            // 2. Target Viewport(s) Selection
            List<Viewport> targetViewports = new List<Viewport>();
            // Dictionary to store target viewport names for the summary
            Dictionary<ElementId, string> targetViewportNames = new Dictionary<ElementId, string>();


            TaskDialog.Show("Instructions", "Now, select the target viewports to align. Press ESC or Finish Selection when done.");

            try
            {
                IList<Reference> targetRefs = uidoc.Selection.PickObjects(ObjectType.Element, new ViewportSelectionFilter(), "Select target viewports to align (press ESC or Finish Selection when done).");

                foreach (Reference targetRef in targetRefs)
                {
                    Viewport targetVp = doc.GetElement(targetRef) as Viewport;
                    if (targetVp != null && targetVp.Id != referenceViewport.Id) // Exclude the reference viewport itself
                    {
                        // Ensure target viewports are on the same sheet as the reference
                        if (targetVp.SheetId == referenceViewport.SheetId)
                        {
                            targetViewports.Add(targetVp);
                            // Get and store the target view's name
                            View targetView = doc.GetElement(targetVp.ViewId) as View;
                            targetViewportNames[targetVp.Id] = targetView != null ? targetView.Name : "Unknown View";
                        }
                        else
                        {
                            // Use the stored name if available, otherwise a generic message
                            string skippedViewName = (doc.GetElement(targetVp.ViewId) as View)?.Name ?? "Unknown View";
                            TaskDialog.Show("Warning", $"Viewport '{skippedViewName}' is on a different sheet and will be skipped.");
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // User cancelled selection, proceed with already selected viewports
            }
            catch (Exception ex)
            {
                message = "Error selecting target viewports: " + ex.Message;
                return Result.Failed;
            }

            if (!targetViewports.Any())
            {
                message = "No target viewports selected or found on the same sheet as the reference. Exiting.";
                return Result.Cancelled;
            }

            // Store original pin status to restore later
            Dictionary<ElementId, bool> originalPinStatus = new Dictionary<ElementId, bool>();
            StringBuilder resultSummary = new StringBuilder();
            int viewportsMovedCount = 0;
            int viewportsAlreadyAlignedCount = 0;

            // 3. Alignment Logic & 4. Revit API Transaction
            using (Transaction trans = new Transaction(doc, "Align Viewports"))
            {
                try
                {
                    trans.Start();

                    // Unpin all target viewports if they are pinned
                    foreach (Viewport targetVp in targetViewports)
                    {
                        originalPinStatus[targetVp.Id] = targetVp.Pinned; // Store original status
                        if (targetVp.Pinned)
                        {
                            targetVp.Pinned = false; // Unpin
                        }
                    }

                    // Perform alignment
                    foreach (Viewport targetVp in targetViewports)
                    {
                        XYZ currentTargetCenter = targetVp.GetBoxCenter();
                        string currentTargetViewName = targetViewportNames.ContainsKey(targetVp.Id) ? targetViewportNames[targetVp.Id] : "Unknown View";

                        // Check if alignment is needed (using a small tolerance for floating point comparisons)
                        if (currentTargetCenter.IsAlmostEqualTo(referenceCenter, 1e-6)) // 1e-6 is a common small tolerance
                        {
                            resultSummary.AppendLine($"- Viewport '{currentTargetViewName}' was already correctly aligned.");
                            viewportsAlreadyAlignedCount++;
                        }
                        else
                        {
                            targetVp.SetBoxCenter(referenceCenter);
                            resultSummary.AppendLine($"- Viewport '{currentTargetViewName}' was moved to align.");
                            viewportsMovedCount++;
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = "Error aligning viewports: " + ex.Message;
                    return Result.Failed;
                }
            }

            // Re-pin viewports if they were originally pinned, in a new transaction
            if (originalPinStatus.Any(kvp => kvp.Value == true))
            {
                using (Transaction rePinTrans = new Transaction(doc, "Re-pin Aligned Viewports"))
                {
                    try
                    {
                        rePinTrans.Start();
                        foreach (Viewport targetVp in targetViewports)
                        {
                            if (originalPinStatus.ContainsKey(targetVp.Id) && originalPinStatus[targetVp.Id])
                            {
                                targetVp.Pinned = true; // Re-pin
                            }
                        }
                        rePinTrans.Commit();
                    }
                    catch (Exception ex)
                    {
                        rePinTrans.RollBack();
                        TaskDialog.Show("Warning", $"Some viewports could not be re-pinned after alignment: {ex.Message}");
                    }
                }
            }

            // Final user message
            string finalMessageTitle = "Viewport Alignment Results";
            StringBuilder finalMessage = new StringBuilder();
            finalMessage.AppendLine($"Alignment complete for {targetViewports.Count} target viewport(s) to reference viewport '{referenceViewName}':\n");
            finalMessage.AppendLine(resultSummary.ToString());
            finalMessage.AppendLine($"\nSummary: {viewportsMovedCount} viewport(s) moved, {viewportsAlreadyAlignedCount} viewport(s) already aligned.");

            TaskDialog.Show(finalMessageTitle, finalMessage.ToString());

            return Result.Succeeded;
        }
    }

    // Helper class for filtering Viewport elements during selection
    public class ViewportSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Viewport;
        }

        public bool AllowReference(Reference refer, XYZ position)
        {
            return false; // Not used for element selection
        }
    }
}