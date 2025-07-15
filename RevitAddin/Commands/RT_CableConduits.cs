/*
 * Copyright (c) 2024.
 * All rights reserved.
 *
 * This C# script is designed for Autodesk Revit 2022.
 * It creates a new conduit type by duplicating an existing one,
 * as the concept of a separate 'ConduitStandardType' does not exist in this API version.
 *
 * Namespace: RTS.Commands
 * Class: RT_CableConduitsClass
 * Author: Gemini
 */

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace RTS.Commands
{
    /// <summary>
    /// A Revit external command that creates a new Conduit Type for Revit 2022.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_CableConduitsClass : IExternalCommand
    {
        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, Revit application, and active document.</param>
        /// <param name="message">A message that can be set by the external application 
        /// to report status or error information.</param>
        /// <param name="elements">A set of elements to which the external application 
        /// can add elements that are to be highlighted in case of failure.</param>
        /// <returns>The result of the command execution.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the essential Revit application and document objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Define the name for the new conduit type
            string newTypeName = "RTS_CableConduits";

            // --- Step 1: Find an existing ConduitType to duplicate ---
            // In Revit 2022, we create new types by duplicating existing ones.
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ConduitType));

            // Use ToElements().FirstOrDefault() for compatibility with the Revit 2022 API
            ConduitType conduitTypeToDuplicate = collector.ToElements().FirstOrDefault() as ConduitType;

            if (conduitTypeToDuplicate == null)
            {
                message = "No existing Conduit Types found in the project to use as a template. Please load a conduit type first.";
                TaskDialog.Show("Action Needed", message);
                return Result.Failed;
            }

            // --- Pre-check: Verify the new type name doesn't already exist ---
            if (collector.Cast<ConduitType>().Any(ct => ct.Name.Equals(newTypeName, StringComparison.OrdinalIgnoreCase)))
            {
                TaskDialog.Show("Already Exists", $"A Conduit Type named '{newTypeName}' already exists in this project.");
                return Result.Cancelled;
            }

            // All modifications to the Revit database must be wrapped in a transaction.
            using (Transaction tx = new Transaction(doc, "Create RTS Conduit Type"))
            {
                try
                {
                    tx.Start();

                    // --- Step 2: Duplicate the existing type ---
                    // This creates a new ConduitType with all the same settings as the original.
                    ConduitType newConduitType = conduitTypeToDuplicate.Duplicate(newTypeName) as ConduitType;

                    if (newConduitType == null)
                    {
                        message = "Failed to duplicate the conduit type.";
                        tx.RollBack();
                        return Result.Failed;
                    }

                    // --- Regarding Sizes ---
                    // In Revit 2022, we cannot programmatically add/remove sizes in a simple way.
                    // The available sizes are determined by the 'Routing Preferences', which
                    // point to loaded conduit fitting families.
                    // The new type 'RTS_CableConduits' will inherit the sizes from the duplicated type.
                    // You can then modify these sizes manually:
                    // 1. In Revit, find the 'RTS_CableConduits' type in the Project Browser under Families > Conduits > Conduit Type.
                    // 2. Right-click and choose 'Type Properties'.
                    // 3. Click 'Edit' next to 'Routing Preferences' to add or remove sizes.

                    tx.Commit();

                    TaskDialog.Show("Success", $"Successfully created Conduit Type '{newTypeName}'.\n\nPlease manage its available sizes through the 'Edit Type' > 'Routing Preferences' dialog.");

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    // If any error occurs, set the failure message,
                    // roll back the transaction, and return a failed result.
                    message = ex.Message;
                    if (tx.HasStarted())
                    {
                        tx.RollBack();
                    }
                    return Result.Failed;
                }
            }
        }
    }
}
