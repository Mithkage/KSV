// RT_PanelConnect.cs
// Purpose: Imports a CSV file and powers Revit Panels based on the imported data.

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RT_PanelConnect
{
    /// <summary>
    /// Revit External Command to connect electrical panels based on a CSV input file.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_PanelConnectClass : IExternalCommand
    {
        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, Revit application, and documents.</param>
        /// <param name="message">A message string which can be set by the external application 
        /// to report status or result information back to Revit.</param>
        /// <param name="elements">A set of elements to which the command will be applied.</param>
        /// <returns>The result of the command execution.</returns>
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // Get the Revit application and document objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            // Fully qualify the Application type to avoid ambiguity with System.Windows.Forms.Application
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // --- 1. PROMPT USER TO SELECT CSV FILE ---
            string csvPath = ShowOpenFileDialog();
            if (string.IsNullOrEmpty(csvPath))
            {
                message = "Operation cancelled. No CSV file was selected.";
                return Result.Cancelled;
            }

            // --- 2. READ AND PARSE THE CSV DATA ---
            List<string[]> parsedData;
            try
            {
                parsedData = File.ReadAllLines(csvPath)
                                     .Skip(1) // Assuming a header row
                                     .Select(line => line.Split('\t'))
                                     .ToList();
            }
            catch (Exception ex)
            {
                message = $"Error reading or parsing CSV file: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // Lists for final reporting
            var updatedPanels = new List<string>();
            var skippedConnections = new List<string>();

            // --- 3. PROCESS DATA AND POWER PANELS WITHIN A TRANSACTION ---
            using (Transaction trans = new Transaction(doc, "Power Panels from CSV"))
            {
                try
                {
                    trans.Start();

                    foreach (var row in parsedData)
                    {
                        // --- EXPECTED CSV FORMAT: ---
                        // Column Index 1: "SWB From" (Source Panel)
                        // Column Index 2: "SWB To" (Panel to connect)

                        if (row.Length < 3) continue;

                        string supplyFromPanelName = row[1].Trim('(', ')').Trim();
                        string panelToPowerName = row[2].Trim('(', ')').Trim();

                        if (string.IsNullOrWhiteSpace(panelToPowerName) || string.IsNullOrWhiteSpace(supplyFromPanelName))
                        {
                            continue;
                        }

                        Element panelToPower = FindPanelByPanelName(doc, panelToPowerName);
                        Element supplyPanel = FindPanelByPanelName(doc, supplyFromPanelName);

                        if (panelToPower == null || supplyPanel == null)
                        {
                            string notFoundMessage = panelToPower == null
                                ? $"Panel to power '{panelToPowerName}' not found in the model."
                                : $"Supply panel '{supplyFromPanelName}' not found in the model.";

                            if (!skippedConnections.Contains(notFoundMessage))
                            {
                                skippedConnections.Add(notFoundMessage);
                            }
                            Debug.WriteLine($"{notFoundMessage} Skipping connection.");
                            continue;
                        }

                        // --- VOLTAGE SYSTEM VALIDATION ---
                        // Use RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM for Revit 2022 API compatibility
                        var panelToPowerDistSystemId = panelToPower.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)?.AsElementId();
                        var supplyPanelDistSystemId = supplyPanel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)?.AsElementId();

                        if (panelToPowerDistSystemId == null || supplyPanelDistSystemId == null || panelToPowerDistSystemId != supplyPanelDistSystemId)
                        {
                            string targetSystemName = panelToPowerDistSystemId != null && panelToPowerDistSystemId != ElementId.InvalidElementId ? doc.GetElement(panelToPowerDistSystemId).Name : "Not Defined";
                            string supplySystemName = supplyPanelDistSystemId != null && supplyPanelDistSystemId != ElementId.InvalidElementId ? doc.GetElement(supplyPanelDistSystemId).Name : "Not Defined";

                            string reason = $"Incompatible voltage systems: '{panelToPowerName}' ({targetSystemName}) and '{supplyFromPanelName}' ({supplySystemName}).";
                            if (!skippedConnections.Contains(reason))
                            {
                                skippedConnections.Add(reason);
                            }
                            Debug.WriteLine($"Skipping connection. {reason}");
                            continue;
                        }

                        var elecSystem = (panelToPower as FamilyInstance)?.MEPModel?.GetElectricalSystems()?.FirstOrDefault();
                        if (elecSystem == null)
                        {
                            Debug.WriteLine($"Element '{panelToPowerName}' does not have a connectable electrical system. Skipping.");
                            continue;
                        }

                        // Use BaseEquipment property for API compatibility and ELEM_PARTITION_PARAM for the Panel Name parameter
                        if (elecSystem.BaseEquipment != null && elecSystem.BaseEquipment.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)?.AsString().Equals(supplyFromPanelName, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Debug.WriteLine($"Panel '{panelToPowerName}' is already correctly powered by '{supplyFromPanelName}'. Skipping.");
                            continue;
                        }

                        // Use BaseEquipment property for API compatibility
                        string updateMessage = elecSystem.BaseEquipment != null
                            ? $"Updated '{panelToPowerName}' from '{elecSystem.BaseEquipment.Name}' to be supplied by '{supplyFromPanelName}'."
                            : $"Connected '{panelToPowerName}' to be supplied by '{supplyFromPanelName}'.";

                        Debug.WriteLine(updateMessage);

                        elecSystem.SelectPanel(supplyPanel as FamilyInstance);

                        // Add to the list of successful updates for reporting
                        updatedPanels.Add($"'{panelToPowerName}' connected to '{supplyFromPanelName}'");
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    message = $"An error occurred: {ex.Message}";
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }
            }

            // --- 4. FINAL REPORTING ---
            string summaryMessage = "Panel connection process complete.\n";
            bool changesMade = false;

            if (updatedPanels.Any())
            {
                summaryMessage += "\nSuccessfully updated connections:\n- " + string.Join("\n- ", updatedPanels);
                changesMade = true;
            }

            if (skippedConnections.Any())
            {
                summaryMessage += "\n\nSkipped connections:\n- " + string.Join("\n- ", skippedConnections);
                changesMade = true;
            }

            if (!changesMade)
            {
                summaryMessage = "Process complete. No panel connections needed to be changed.";
            }

            TaskDialog.Show("Process Report", summaryMessage);


            return Result.Succeeded;
        }

        /// <summary>
        /// Displays an OpenFileDialog to let the user choose a CSV file.
        /// </summary>
        /// <returns>The file path of the selected CSV, or null if cancelled.</returns>
        private string ShowOpenFileDialog()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// Finds an electrical panel (FamilyInstance) in the document by its "Panel Name" parameter.
        /// This is more reliable than checking Element.Name.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="panelName">The value of the "Panel Name" parameter to find.</param>
        /// <returns>The found Element, or null if not found.</returns>
        private Element FindPanelByPanelName(Document doc, string panelName)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .OfClass(typeof(FamilyInstance))
                .FirstOrDefault(e => e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)?.AsString()?.Equals(panelName, StringComparison.OrdinalIgnoreCase) ?? false);
        }
    }
}
