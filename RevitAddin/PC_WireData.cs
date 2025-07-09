//
// File: PC_WireData.cs
//
// Namespace: PC_WireData
//
// Class: PC_WireDataClass
//
// Function: This Revit external command reads cleaned cable data (PC_Data) from
// project extensible storage and uses it to update corresponding electrical
// wires in the active Revit model.
//
// Author: AI (Based on user requirements)
//
// Date: June 28, 2025
//
#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms; // For TaskDialog
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Electrical; // For Electrical System, Wire
using Autodesk.Revit.DB.ExtensibleStorage;
using System.Text.Json; // For JSON deserialization
#endregion

namespace PC_WireData
{
    /// <summary>
    /// The main class for the Revit external command to update electrical wires from stored cable data.
    /// Implements the IExternalCommand interface.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_WireDataClass : IExternalCommand
    {
        // IMPORTANT: These GUIDs and names MUST match those used in PC_Extensible.cs
        private static readonly Guid SchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336");
        private const string FieldName = "PC_DataJson";
        private const string SchemaName = "PC_ExtensibleDataSchema";

        // GUID for the RTS_ID parameter on Electrical Fixtures/Equipment and Wires
        private static readonly Guid RtsIdParameterGuid = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");

        /// <summary>
        /// The main entry point for the external command. Revit calls this method when the user clicks the button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1. Recall cleaned cable data (PC_Data) from extensible storage
                List<CableData> cleanedCableData = RecallCableDataFromExtensibleStorage(doc);

                if (cleanedCableData == null || cleanedCableData.Count == 0)
                {
                    TaskDialog.Show("No Data Found", "No cleaned cable data (PC_Data) was found in the project's extensible storage. Please run 'Process & Save Cable Data' first.");
                    return Result.Cancelled;
                }

                // 2. Process and update electrical wires in the model
                // This operation needs to be wrapped in a Revit transaction
                using (Transaction tx = new Transaction(doc, "Update Electrical Wires from PC_Data"))
                {
                    tx.Start();
                    try
                    {
                        int updatedWiresCount = UpdateElectricalWires(doc, cleanedCableData);
                        tx.Commit();
                        TaskDialog.Show("Wire Update Complete", $"Successfully updated {updatedWiresCount} electrical wires based on PC_Data. " +
                                                               "Please check the updated parameters in your model.");
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        message = $"Failed to update electrical wires: {ex.Message}";
                        TaskDialog.Show("Wire Update Error", message);
                        return Result.Failed;
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An unexpected error occurred: {ex.Message}\n\nStackTrace: {ex.StackTrace}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }

        #region Extensible Storage Data Recall (Copied from PC_Extensible for self-containment)

        /// <summary>
        /// Recalls the cleaned cable data (PC_Data) from extensible storage.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <returns>A List of CableData, or an empty list if no data is found or an error occurs during recall.</returns>
        private List<CableData> RecallCableDataFromExtensibleStorage(Document doc)
        {
            Schema schema = Schema.Lookup(SchemaGuid); // Look up the schema by its GUID

            if (schema == null)
            {
                return new List<CableData>();
            }

            // Find existing DataStorage elements associated with our schema
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                // Check if the DataStorage element contains an entity for our schema
                if (ds.GetEntity(schema) != null)
                {
                    dataStorage = ds;
                    break;
                }
            }

            if (dataStorage == null)
            {
                return new List<CableData>();
            }

            // Get the Entity from the DataStorage
            Entity entity = dataStorage.GetEntity(schema);

            if (!entity.IsValid())
            {
                return new List<CableData>();
            }

            // Get the JSON string from the field
            string jsonString = entity.Get<string>(schema.GetField(FieldName));

            if (string.IsNullOrEmpty(jsonString))
            {
                return new List<CableData>();
            }

            // Deserialize the JSON string back into a List<CableData>
            try
            {
                return JsonSerializer.Deserialize<List<CableData>>(jsonString) ?? new List<CableData>();
            }
            catch (JsonException ex)
            {
                TaskDialog.Show("Data Recall Error", $"Failed to deserialize stored PC_Data: {ex.Message}. The stored data might be corrupt or incompatible.");
                return new List<CableData>();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Data Recall Error", $"An unexpected error occurred during PC_Data recall: {ex.Message}");
                return new List<CableData>();
            }
        }
        #endregion

        #region Wire Update Logic

        /// <summary>
        /// Updates electrical wires in the Revit model based on the provided cable data.
        /// This method assumes a transaction is already open.
        /// </summary>
        /// <param name="doc">The Revit Document.</param>
        /// <param name="cleanedCableData">The list of cleaned cable data.</param>
        /// <returns>The count of updated wires.</returns>
        private int UpdateElectricalWires(Document doc, List<CableData> cleanedCableData)
        {
            int updatedWiresCount = 0;
            List<string> issuesSummary = new List<string>();

            // Create a dictionary for quick lookup of cable data by "To" value.
            // If multiple CableData entries have the same "To" value, the first one encountered will be used.
            var cableDataByTo = cleanedCableData
                                    .Where(cd => !string.IsNullOrEmpty(cd.To))
                                    .GroupBy(cd => cd.To)
                                    .ToDictionary(g => g.Key, g => g.First());

            // Collector for electrical wires
            FilteredElementCollector wireCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Wire)
                .WhereElementIsNotElementType();

            foreach (Element wireElement in wireCollector)
            {
                Wire wire = wireElement as Wire;
                if (wire == null) continue;

                // Get the ElectricalSystem that this wire belongs to using MEPSystem property
                Element systemElement = wire.MEPSystem;
                if (systemElement == null)
                {
                    issuesSummary.Add($"Wire ID: {wire.Id} has no associated Electrical System.");
                    continue;
                }

                ElectricalSystem elecSystem = systemElement as ElectricalSystem;
                if (elecSystem == null)
                {
                    issuesSummary.Add($"Wire ID: {wire.Id} system (ID: {systemElement.Id}) could not be cast to ElectricalSystem.");
                    continue;
                }

                // Identify the "To" (load-side) element connected by this system
                // The user specified: "The wire will be looked up using 'To' which represents the load-side."
                // And: "The 'To' element is always one of the ElectricalSystem.Elements (excluding the panel itself)."
                // And: "Assume that a wire's ElectricalSystem will contain only one load element that corresponds to the CableData.To."
                Element toElement = null;
                if (elecSystem.Elements != null && elecSystem.BaseEquipment != null)
                {
                    foreach (Element sysElem in elecSystem.Elements)
                    {
                        // Ensure it's not the source panel itself
                        if (sysElem.Id != elecSystem.BaseEquipment.Id)
                        {
                            toElement = sysElem;
                            break; // Found the load-side element
                        }
                    }
                }

                if (toElement == null)
                {
                    issuesSummary.Add($"Wire ID: {wire.Id} (System: {elecSystem.Id}) - No distinct load-side element found in its electrical system.");
                    continue;
                }

                // Get the RTS_ID parameter from the "To" element
                Parameter toElementRtsIdParam = toElement.get_Parameter(RtsIdParameterGuid);
                string toElementRtsIdValue = null;

                if (toElementRtsIdParam != null && toElementRtsIdParam.HasValue)
                {
                    toElementRtsIdValue = toElementRtsIdParam.AsString();
                }

                if (string.IsNullOrEmpty(toElementRtsIdValue))
                {
                    issuesSummary.Add($"Wire ID: {wire.Id} (To Element ID: {toElement.Id}, Name: {toElement.Name}) - 'RTS_ID' parameter is null or empty on the load element.");
                    continue;
                }

                // Now try to match this RTS_ID value with the "To" column in our cable data
                if (cableDataByTo.TryGetValue(toElementRtsIdValue, out CableData correspondingCableData))
                {
                    // Found matching cable data. Now update the wire's RTS_ID parameter.
                    Parameter wireRtsIdParam = wire.get_Parameter(RtsIdParameterGuid); // This is the RTS_ID param ON THE WIRE

                    if (wireRtsIdParam != null && !wireRtsIdParam.IsReadOnly)
                    {
                        string valueToSet = string.IsNullOrEmpty(correspondingCableData.CableReference) ?
                                            correspondingCableData.To :
                                            correspondingCableData.CableReference;
                        wireRtsIdParam.Set(valueToSet);
                        updatedWiresCount++;
                    }
                    else
                    {
                        issuesSummary.Add($"Wire ID: {wire.Id} (To Element RTS_ID: '{toElementRtsIdValue}') - 'RTS_ID' parameter on wire is read-only or not found, cannot update.");
                    }
                }
                else
                {
                    issuesSummary.Add($"Wire ID: {wire.Id} (To Element RTS_ID: '{toElementRtsIdValue}') - No matching CableData 'To' value found in stored data.");
                }
            }

            // Provide a summary of issues to the user
            if (issuesSummary.Any())
            {
                string messageHeader = "Some wires could not be fully processed or matched:";
                string detailedIssues = string.Join("\n- ", issuesSummary.Take(10)); // Show up to 10 issues
                string moreIssuesIndicator = issuesSummary.Count > 10 ? $"\n...and {issuesSummary.Count - 10} more issues." : "";

                TaskDialog.Show("Wire Update Details & Issues", $"{messageHeader}\n\n- {detailedIssues}{moreIssuesIndicator}\n\nPlease review these wires and your data sources.");
            }

            return updatedWiresCount;
        }

        #endregion

        #region Data Classes (Must match PC_Extensible.cs)
        /// <summary>
        /// A data class to hold the values for a single row of processed cable data from the CSV.
        /// Public for JSON serialization.
        /// </summary>
        public class CableData
        {
            public string CableReference { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string CableType { get; set; }
            public string CableCode { get; set; }
            public string CableConfiguration { get; set; }
            public string Cores { get; set; }
            public string ConductorActive { get; set; }
            public string Insulation { get; set; }
            public string ConductorEarth { get; set; }
            public string SeparateEarthForMulticore { get; set; }
            public string CableLength { get; set; }
            public string TotalCableRunWeight { get; set; }
            public string NominalOverallDiameter { get; set; }
            public string NumberOfActiveCables { get; set; }
            public string ActiveCableSize { get; set; }
            public string NumberOfNeutralCables { get; set; }
            public string NeutralCableSize { get; set; }
            public string NumberOfEarthCables { get; set; }
            public string EarthCableSize { get; set; }
            public string CablesKgPerM { get; set; }
            public string AsNsz3008CableDeratingFactor { get; set; }
        }
        #endregion
    }
}
