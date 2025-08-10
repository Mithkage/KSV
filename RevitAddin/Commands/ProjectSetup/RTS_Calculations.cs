/*
 * Copyright (c) 2024.
 * All rights reserved.
 *
 * It creates a new user-created workset named "RTS_Calculations".
 * The script will first check if worksharing is enabled and if the
 * workset name is unique before attempting to create it.
 *
 * Namespace: RTS.Commands
 * Class: RTS_CalculationsClass
 * Author: Gemini
 */

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace RTS.Commands.ProjectSetup
{
    /// <summary>
    /// A Revit external command that creates a new workset.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_CalculationsClass : IExternalCommand
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

            // Define the name for the new workset
            string worksetName = "RTS_Calculations";

            // --- Step 1: Check if the model is workshared ---
            // Worksets can only be created in a workshared model.
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Not a Workshared Model", "This command requires an active workshared model. Please enable worksharing and try again.");
                return Result.Cancelled;
            }

            // All modifications to the Revit database must be wrapped in a transaction.
            // Note: Creating a workset requires a transaction.
            using (Transaction tx = new Transaction(doc, "Create Calculations Workset"))
            {
                try
                {
                    tx.Start();

                    // --- Step 2: Check if the workset name is unique ---
                    // This prevents creating duplicate worksets.
                    if (!WorksetTable.IsWorksetNameUnique(doc, worksetName))
                    {
                        TaskDialog.Show("Workset Exists", $"A workset named '{worksetName}' already exists in this project.");
                        // We don't need to roll back as no changes were made.
                        tx.RollBack();
                        return Result.Cancelled;
                    }

                    // --- Step 3: Create the new workset ---
                    // We create a user-visible workset of the 'UserWorkset' kind.
                    Workset.Create(doc, worksetName);

                    // Commit the transaction to save the changes.
                    tx.Commit();

                    TaskDialog.Show("Success", $"Successfully created workset '{worksetName}'.");

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
