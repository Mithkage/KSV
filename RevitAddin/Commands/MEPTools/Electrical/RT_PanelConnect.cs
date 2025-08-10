// RT_PanelConnect.cs
// Purpose: Reads cleaned cable data (PC_Data) from project extensible storage
//          and connects Revit Electrical Equipment and Fixtures based on this data
//          using a shared parameter (RTS_ID).
//
// Author: ReTick Solutions (Modified by AI)
// Date: June 28, 2025 (Updated to read from Extensible Storage)

#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.ExtensibleStorage; // Added for extensible storage
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO; // Still needed for Path.Combine, etc., if other methods use it for local files, but not for core data reading
using System.Linq;
using System.Text;
using System.Windows.Forms; // For TaskDialog, OpenFileDialog removed for core function
using System.Text.Json;
#endregion

namespace RTS.Commands.MEPTools.Electrical
{
    /// <summary>
    /// Revit External Command to connect electrical elements based on data recalled from Extensible Storage.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_PanelConnectClass : IExternalCommand
    {
        // Shared Parameter GUID for RTS_ID (3175a27e-d386-4567-bf10-2da1a9cbb73b)
        private static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
        private const string RTS_ID_NAME = "RTS_ID"; // For reporting clarity in messages

        // IMPORTANT: These GUIDs and names MUST match those used in PC_Extensible.cs
        private static readonly Guid SchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336");
        private const string FieldName = "PC_DataJson";
        private const string SchemaName = "PC_ExtensibleDataSchema";

        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the Revit application and document objects
            UIApplication uiapp = commandData.Application;
            // UIDocument uidoc = uiapp.ActiveUIDocument; // Not directly used after getting doc
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application; // Not directly used
            Document doc = uiapp.ActiveUIDocument.Document;

            // Lists for final reporting
            var updatedConnections = new List<string>();
            var skippedConnections = new List<string>();
            var duplicateDataEntries = new List<string>(); // To report duplicate (From, To) pairs in PC_Data

            try
            {
                // --- 1. RECALL DATA FROM EXTENSIBLE STORAGE ---
                List<CableData> cleanedCableData = RecallCableDataFromExtensibleStorage(doc);

                if (cleanedCableData == null || !cleanedCableData.Any())
                {
                    TaskDialog.Show("No Data Found", "No cleaned cable data (PC_Data) was found in the project's extensible storage. Please run 'Process & Save Cable Data' first.");
                    return Result.Cancelled;
                }

                // --- 2. PRE-PROCESS DATA: Identify unique connection pairs and duplicates ---
                var processedConnectionPairs = new List<Tuple<string, string>>();
                var seenPairs = new HashSet<Tuple<string, string>>();

                foreach (var cableDataEntry in cleanedCableData)
                {
                    string supplyFromRtsId = cableDataEntry.From?.Trim();
                    string elementToPowerRtsId = cableDataEntry.To?.Trim();

                    // Skip entries where From or To are null/empty, as these cannot form a connection
                    if (string.IsNullOrWhiteSpace(supplyFromRtsId) || string.IsNullOrWhiteSpace(elementToPowerRtsId))
                    {
                        skippedConnections.Add($"Skipped CableData entry due to empty 'From' or 'To' values. Cable Ref: '{cableDataEntry.CableReference ?? "N/A"}', From: '{supplyFromRtsId ?? "N/A"}', To: '{elementToPowerRtsId ?? "N/A"}'");
                        continue;
                    }

                    var currentPair = Tuple.Create(supplyFromRtsId, elementToPowerRtsId);

                    if (!seenPairs.Add(currentPair))
                    {
                        // This pair has been seen before (duplicate in terms of From-To connection)
                        duplicateDataEntries.Add($"Duplicate connection pair found in PC_Data: From='{supplyFromRtsId}', To='{elementToPowerRtsId}'. Skipping additional processing for this pair.");
                        continue; // Skip processing this duplicate
                    }
                    processedConnectionPairs.Add(currentPair); // Add unique pair for processing
                }

                if (!processedConnectionPairs.Any())
                {
                    TaskDialog.Show("No Valid Connections", "No valid (From, To) connection pairs were found in the stored PC_Data after filtering. Process cancelled.");
                    return Result.Cancelled;
                }

                // --- 3. PROCESS DATA AND CONNECT ELEMENTS WITHIN A TRANSACTION ---
                using (Transaction trans = new Transaction(doc, "Connect Electrical Elements from PC_Data"))
                {
                    trans.Start();

                    foreach (var connectionPair in processedConnectionPairs)
                    {
                        string supplyFromRtsId = connectionPair.Item1;
                        string elementToPowerRtsId = connectionPair.Item2;

                        // Find elements by RTS_ID
                        Element elementToPower = FindElementByRTSID(doc, elementToPowerRtsId);
                        Element supplyElement = FindElementByRTSID(doc, supplyFromRtsId);

                        // Error Handling: Element not found by RTS_ID
                        if (elementToPower == null)
                        {
                            string notFoundMessage = $"Element with {RTS_ID_NAME} '{elementToPowerRtsId}' (To Connect) not found in the model or its {RTS_ID_NAME} parameter is null/empty. Skipping connection.";
                            skippedConnections.Add(notFoundMessage);
                            Debug.WriteLine(notFoundMessage);
                            continue;
                        }

                        if (supplyElement == null)
                        {
                            string notFoundMessage = $"Element with {RTS_ID_NAME} '{supplyFromRtsId}' (Supply From) not found in the model or its {RTS_ID_NAME} parameter is null/empty. Skipping connection.";
                            skippedConnections.Add(notFoundMessage);
                            Debug.WriteLine(notFoundMessage);
                            continue;
                        }

                        // Ensure both elements are FamilyInstances, as electrical connections typically operate on them
                        FamilyInstance fiToPower = elementToPower as FamilyInstance;
                        FamilyInstance fiSupply = supplyElement as FamilyInstance;

                        if (fiToPower == null || fiToPower.MEPModel == null || !fiToPower.MEPModel.GetElectricalSystems().Any())
                        {
                            string reason = $"Element '{elementToPowerRtsId}' is not a valid electrical FamilyInstance (or has no MEPModel/ElectricalSystems). Skipping.";
                            skippedConnections.Add(reason);
                            Debug.WriteLine(reason);
                            continue;
                        }

                        // The supply element must be a FamilyInstance to be used in SelectPanel
                        if (fiSupply == null)
                        {
                            string reason = $"Supply element '{supplyFromRtsId}' is not a valid FamilyInstance. Skipping.";
                            skippedConnections.Add(reason);
                            Debug.WriteLine(reason);
                            continue;
                        }

                        // Get the electrical system of the element to be powered
                        ElectricalSystem elecSystem = fiToPower.MEPModel.GetElectricalSystems()
                            .FirstOrDefault(s => s.SystemType == ElectricalSystemType.PowerCircuit);

                        if (elecSystem == null)
                        {
                            string reason = $"Element '{elementToPowerRtsId}' does not have a suitable electrical system (Power Circuit) to connect. Skipping.";
                            skippedConnections.Add(reason);
                            Debug.WriteLine(reason);
                            continue;
                        }

                        // --- VOLTAGE SYSTEM VALIDATION ---
                        // Ensure distribution systems (voltages) match
                        var elementToPowerDistSystemId = elementToPower.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)?.AsElementId();
                        var supplyElementDistSystemId = supplyElement.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)?.AsElementId();

                        if (elementToPowerDistSystemId == null || supplyElementDistSystemId == null || !elementToPowerDistSystemId.Equals(supplyElementDistSystemId))
                        {
                            string targetSystemName = elementToPowerDistSystemId != null && doc.GetElement(elementToPowerDistSystemId) != null ? doc.GetElement(elementToPowerDistSystemId).Name : "Not Defined";
                            string supplySystemName = supplyElementDistSystemId != null && doc.GetElement(supplyElementDistSystemId) != null ? doc.GetElement(supplyElementDistSystemId).Name : "Not Defined";

                            string reason = $"Incompatible voltage systems: '{elementToPowerRtsId}' ({targetSystemName}) and '{supplyFromRtsId}' ({supplySystemName}). Skipping connection.";
                            skippedConnections.Add(reason);
                            Debug.WriteLine($"Skipping connection. {reason}");
                            continue;
                        }

                        // Check if already correctly connected to prevent unnecessary updates
                        string currentSupplyRtsId = null;
                        if (elecSystem.BaseEquipment != null)
                        {
                            // Get the RTS_ID of the current supplying equipment
                            Parameter currentSupplyRtsParam = elecSystem.BaseEquipment.get_Parameter(RTS_ID_GUID);
                            if (currentSupplyRtsParam != null && currentSupplyRtsParam.HasValue)
                            {
                                currentSupplyRtsId = currentSupplyRtsParam.AsString();
                            }
                        }

                        if (currentSupplyRtsId != null && currentSupplyRtsId.Equals(supplyFromRtsId, StringComparison.OrdinalIgnoreCase))
                        {
                            Debug.WriteLine($"Element '{elementToPowerRtsId}' is already correctly powered by '{supplyFromRtsId}'. Skipping.");
                            continue; // No change needed
                        }

                        // Perform the connection
                        string updateMessage = elecSystem.BaseEquipment != null
                            ? $"Updated '{elementToPowerRtsId}' (was supplied by '{currentSupplyRtsId ?? "unknown"}') to be supplied by '{supplyFromRtsId}'."
                            : $"Connected '{elementToPowerRtsId}' to be supplied by '{supplyFromRtsId}'.";

                        Debug.WriteLine(updateMessage);

                        // SelectPanel sets the supplying element for the electrical system
                        elecSystem.SelectPanel(fiSupply);

                        // Add to the list of successful updates for reporting
                        updatedConnections.Add($"'{elementToPowerRtsId}' connected to '{supplyFromRtsId}'");
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                // This catch block handles exceptions outside the transaction scope (e.g., data recall issues)
                message = $"An error occurred during process: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }

            // --- 4. FINAL REPORTING ---
            string summaryMessage = "Electrical element connection process complete.\n";
            bool changesMade = false;

            if (updatedConnections.Any())
            {
                summaryMessage += "\nSuccessfully updated connections:\n- " + string.Join("\n- ", updatedConnections);
                changesMade = true;
            }

            if (duplicateDataEntries.Any())
            {
                summaryMessage += "\n\nDuplicate connection pairs found in stored PC_Data (skipped):\n- " + string.Join("\n- ", duplicateDataEntries.Take(10));
                if (duplicateDataEntries.Count > 10) summaryMessage += $"\n...and {duplicateDataEntries.Count - 10} more duplicate entries.";
                changesMade = true;
            }

            if (skippedConnections.Any())
            {
                summaryMessage += "\n\nSkipped connections (with reasons):\n- " + string.Join("\n- ", skippedConnections.Take(10));
                if (skippedConnections.Count > 10) summaryMessage += $"\n...and {skippedConnections.Count - 10} more skipped connections.";
                changesMade = true;
            }

            if (!changesMade)
            {
                summaryMessage = "Process complete. No electrical element connections needed to be changed or valid connections found.";
            }

            TaskDialog.Show("Process Report", summaryMessage);

            return Result.Succeeded;
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

        /// <summary>
        /// Finds an electrical element (Electrical Equipment or Electrical Fixture) in the document by its RTS_ID shared parameter.
        /// Returns null if the element is not found, or if the RTS_ID parameter is missing or empty on the element.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="rtsIdValue">The value of the RTS_ID shared parameter to find.</param>
        /// <returns>The found Element (FamilyInstance), or null if not found or RTS_ID is invalid/missing.</returns>
        private Element FindElementByRTSID(Document doc, string rtsIdValue)
        {
            // If the lookup value is empty, we cannot find an element.
            if (string.IsNullOrWhiteSpace(rtsIdValue))
            {
                Debug.WriteLine($"Attempted to find element with empty {RTS_ID_NAME} value. Returning null.");
                return null;
            }

            // Create a multi-category filter for Electrical Equipment and Electrical Fixtures
            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures
            };
            ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(categories);

            // Collect all FamilyInstances that belong to the specified categories
            var collector = new FilteredElementCollector(doc)
                .WherePasses(categoryFilter)
                .OfClass(typeof(FamilyInstance)) // Ensure we are only looking at instances
                .WhereElementIsNotElementType(); // Exclude family types

            foreach (Element e in collector)
            {
                // Attempt to get the shared parameter by its GUID
                Parameter rtsIdParam = e.get_Parameter(RTS_ID_GUID);

                // Check if the parameter exists and has a value
                if (rtsIdParam != null && rtsIdParam.HasValue)
                {
                    string paramValue = rtsIdParam.AsString();
                    // Perform a case-insensitive comparison
                    if (!string.IsNullOrEmpty(paramValue) && paramValue.Equals(rtsIdValue, StringComparison.OrdinalIgnoreCase))
                    {
                        return e; // Found a matching element
                    }
                }
            }
            return null; // No matching element found or RTS_ID parameter was missing/empty/not matching
        }

        #region Data Classes (Must match PC_Extensible.cs and PC_WireData.cs)
        /// <summary>
        /// A data class to hold the values for a single row of processed cable data.
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
