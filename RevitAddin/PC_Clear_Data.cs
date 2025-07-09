using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PC_Clear_Data
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_Clear_DataClass : IExternalCommand 
    {
        // Helper class to store parameter information
        private class ParameterInfo
        {
            public Guid Guid { get; }
            public string DataType { get; } // "TEXT" or "YESNO"

            public ParameterInfo(string guidString, string dataType)
            {
                // Clean the GUID string by removing spaces and hyphens if they are not standard
                string cleanedGuidString = Regex.Replace(guidString, @"\s+", "");
                if (!Guid.TryParse(cleanedGuidString, out Guid parsedGuid))
                {
                    throw new ArgumentException($"Invalid GUID format: {guidString}");
                }
                Guid = parsedGuid;
                DataType = dataType.ToUpper();
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Define the GUID for the PC_PowerCAD parameter (condition)
            Guid pcPowerCadConditionGuid;
            try
            {
                pcPowerCadConditionGuid = new Guid("8f31d68f-60c9-4ec6-a7ff-78a6e3bdaab6");
            }
            catch (FormatException)
            {
                TaskDialog.Show("Error", "Invalid GUID format for PC_PowerCAD condition parameter.");
                return Result.Failed;
            }

            // List of Shared Parameters to iterate through and set to null
            // GUIDs are cleaned of spaces.
            // PC_PowerCAD, PC_SWB To, and PC_SWB From have been removed from this list.
            var parametersToClearInfo = new List<ParameterInfo>();
            try
            {
                parametersToClearInfo.AddRange(new List<ParameterInfo>
                {
                    new ParameterInfo("c3c68f41-8c3b-4f20-b355-bc8263ab557c", "TEXT"), // Protective Device Rating (A)
                    new ParameterInfo("00804d39-1e83-40ba-bc2e-7309f0a093eb", "TEXT"), // PC_Working Distance (mm)
                    new ParameterInfo("92be8f85-927e-4279-a421-346f5d7e1cbb", "TEXT"), // PC_Upstream Diversity
                    new ParameterInfo("29197273-0ba6-42b9-bce3-3b598520276a", "TEXT"), // PC_Switchgear Trip Unit Type
                    new ParameterInfo("1f959198-b192-4a99-af24-729fbcceeb68", "TEXT"), // PC_Switchgear Manufacturer
                    new ParameterInfo("9d5ab9c2-09e2-4d42-ae2f-2a5fba6f7131", "TEXT"), // PC_SWB Type
                    // new ParameterInfo("e142b0ed-d084-447a-991b-d9a3a3f67a8d", "TEXT"), // PC_SWB To (REMOVED)
                    new ParameterInfo("5032629e-1aa1-4cfd-80c0-6533dde3bcb5", "TEXT"), // PC_SWB PF
                    new ParameterInfo("d325322e-9981-450f-8d41-11bc850b499f", "TEXT"), // PC_SWB Load Scope
                    new ParameterInfo("60f670e1-7d54-4ffc-b0f5-ed62c08d3b90", "TEXT"), // PC_SWB Load
                    // new ParameterInfo("5dd52911-7bcd-4a06-869d-73fcef59951c", "TEXT"), // PC_SWB From (REMOVED)
                    new ParameterInfo("d748433f-b90d-4fef-ac6a-3f6e5bce18e9", "TEXT"), // PC_Standard
                    new ParameterInfo("dbd6ec80-8f5c-417f-921a-2df866c1337c", "TEXT"), // PC_Protective Device Type
                    new ParameterInfo("c4c77239-bb4a-48d0-acf1-06e49f1773fa", "TEXT"), // PC_Protective Device Trip Setting (A)
                    new ParameterInfo("50889a75-86ce-4640-9ee2-6495eee39ccf", "TEXT"), // PC_Protective Device Settings
                    new ParameterInfo("c8cd91c1-3d04-46db-b103-700475ba47fd", "TEXT"), // PC_Protective Device OCR/Trip Unit
                    new ParameterInfo("043ac161-7a1c-4a9e-90b1-850e797b39e8", "TEXT"), // PC_Protective Device Model
                    new ParameterInfo("285d845a-a1d9-4821-a3d5-09a9ee956c28", "TEXT"), // PC_Protective Device Manufacturer
                    new ParameterInfo("9eb0f777-3bf0-4398-94f9-e7ab26d1fc19", "TEXT"), // PC_Protective Device Description
                    new ParameterInfo("9a1a8f15-5316-4666-84ee-d84b8c84d8f1", "TEXT"), // PC_Protective Device Breaking Capacity (kA)
                    new ParameterInfo("e1fe521c-673d-4d5a-a69a-407ef8f82518", "TEXT"), // PC_Prospective Fault 3ø Isc (kA)
                    new ParameterInfo("bdd53b50-3fa1-40fe-8fcc-e04641efd658", "TEXT"), // PC_Position
                    new ParameterInfo("43e35e25-e0ab-4c22-86ca-cb3fb382f5bb", "TEXT"), // PC_Minimum PPE Rating
                    new ParameterInfo("bc3f0ed3-e9f5-456f-83c6-d7f2dc0db169", "TEXT"), // PC_Isolator Type
                    new ParameterInfo("5ef35b2c-4ad3-4cad-b099-27efe2df1c92", "TEXT"), // PC_Isolator Rating (A)
                    new ParameterInfo("62ba4120-7d8c-4dcb-8eac-feec66010e05", "TEXT"), // PC_Installation Method
                    new ParameterInfo("217e2d3a-b6e8-4eda-93e4-8ad41f3883c4", "TEXT"), // PC_Incident Energy
                    new ParameterInfo("a07e63e4-dafa-4da6-aaf8-f3dca5fea842", "TEXT"), // PC_Enclosure Width (mm)
                    new ParameterInfo("a63e0421-4fee-49ca-92ce-fdb2c513d814", "TEXT"), // PC_Enclosure Height (mm)
                    new ParameterInfo("4bff1aed-3203-4c84-91aa-de7eba537b8d", "TEXT"), // PC_Enclosure Depth(mm)
                    new ParameterInfo("104f0a87-7c20-4c35-8329-14fd63bcf6b9", "YESNO"),// PC_Enabled
                    new ParameterInfo("1aff59c0-8add-4dae-851c-64ecd7543f75", "TEXT"), // PC_Electrode Configuration
                    new ParameterInfo("214a1cbb-3f5d-48fd-b23b-7915d8e28c6f", "TEXT"), // PC_Conductor Gap (mm)
                    new ParameterInfo("4dc884d6-07f4-4546-b3aa-98a13d5ae6f1", "TEXT"), // PC_Clearance Time (sec)
                    new ParameterInfo("ed6e7cdf-0a9f-400f-a661-48f6c16593df", "TEXT"), // PC_Catalogue Number
                    new ParameterInfo("5efc9508-f2e9-40ee-b353-d68429766a37", "TEXT"), // PC_Cable Type
                    new ParameterInfo("dbd8ce49-4c2b-4ce5-a3ba-fca4c28e3d15", "TEXT"), // PC_Cable Size - Neutral conductors
                    new ParameterInfo("50b6ef99-8f5e-42e1-8645-cce97f6b94b6", "TEXT"), // PC_Cable Size - Earthing conductor
                    new ParameterInfo("91c8321f-1342-4efa-b648-7ba5e95c0085", "TEXT"), // PC_Cable Size - Active conductors
                    // new ParameterInfo("da8a3228-00aa-4f2d-a472-1ba675284cef", "TEXT"), // PC_Cable Reference
                    new ParameterInfo("09007343-dd0a-44c3-b04d-145118778ac3", "TEXT"), // PC_Cable Length
                    new ParameterInfo("dc328944-dcda-4612-8a2e-ecaed994c876", "TEXT"), // PC_Cable Insulation
                    new ParameterInfo("33a25660-bed5-4c55-b0f1-cc2809b3879a", "TEXT"), // PC_Bus/Chassis Rating (A)
                    new ParameterInfo("dc2bcf24-804e-4407-b80d-2b2869dcfbdb", "TEXT"), // PC_Bus Type
                    new ParameterInfo("46482cbb-5247-4dc4-970d-2854a559bca5", "TEXT"), // PC_Arc Flash PPE Category
                    new ParameterInfo("4fc57404-37fd-459a-b432-0769b13fc26c", "TEXT"), // PC_Arc Flash Boundary (mm)
                    new ParameterInfo("adcdc2ac-8720-43ca-b4c4-3e9a3c3f53a3", "TEXT"), // PC_Arc Fault (kA)
                    new ParameterInfo("62ba1846-99f5-4a7f-985a-9255d54c5b93", "TEXT"), // PC_Active Conductor material
                    new ParameterInfo("f5d583ec-f511-4653-828e-9b45281baa54", "TEXT"), // PC_# of Phases
                    // new ParameterInfo("8f31d68f-60c9-4ec6-a7ff-78a6e3bdaab6", "YESNO"),// PC_PowerCAD (REMOVED - it's the condition parameter)
                    new ParameterInfo("c9ae575e-648e-4152-ab07-22fc1c895542", "TEXT"), // PC_Frame Size
                    new ParameterInfo("f3c6038d-7300-46a5-8787-14560678f531", "TEXT"), // PC_Cores
                    new ParameterInfo("ee5c5f1a-7e6e-480b-9591-c1391cf0990b", "TEXT"), // PC_Earth Conductor material
                    new ParameterInfo("c2631a30-511f-450a-92b6-b5c95274ae7b", "TEXT"), // PC_Accum Volt Drop Incl FSC
                    new ParameterInfo("5b362b2d-ceec-42db-93a0-ca5ccb630e9e", "TEXT"), // PC_Prospective Fault at End of Cable
                    new ParameterInfo("4bf4ee16-1be4-47a4-bcd8-9fe79429e9f0", "TEXT"), // PC_No. of Conduits
                    new ParameterInfo("0b552581-2c97-48e4-b71a-0fdc302bdb4b", "TEXT"), // PC_Conduit Size
                    new ParameterInfo("5eb8da2d-a094-4f80-8626-9960fb6b4aa3", "TEXT"), // PC_Design Progress
                    new ParameterInfo("98fa8f80-0219-4b99-bf3d-8da7c74f356d", "TEXT"), // PC_Nominal Overall Diameter
                    new ParameterInfo("cbd9da23-c8eb-4233-af6b-6337c05c2f12", "TEXT")  // PC_Separate Earth for Multicore
                });
            }
            catch (ArgumentException ex)
            {
                TaskDialog.Show("Error", $"Error initializing parameter list: {ex.Message}");
                return Result.Failed;
            }


            // Filter for Detail Items
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_DetailComponents);
            collector.OfClass(typeof(FamilyInstance));

            IList<Element> detailItems = collector.ToElements();

            if (!detailItems.Any())
            {
                TaskDialog.Show("Information", "No Detail Items found in the document.");
                return Result.Succeeded;
            }

            int itemsProcessedCount = 0;
            int parametersClearedCount = 0;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Clear Detail Item Parameters");

                foreach (Element el in detailItems)
                {
                    FamilyInstance detailItem = el as FamilyInstance;
                    if (detailItem == null) continue;

                    // Check the PC_PowerCAD parameter
                    Parameter pcPowerCadParam = detailItem.get_Parameter(pcPowerCadConditionGuid);

                    // PC_PowerCAD is YESNO type; 1 means Yes, 0 means No.
                    if (pcPowerCadParam != null && pcPowerCadParam.HasValue && pcPowerCadParam.AsInteger() == 1)
                    {
                        itemsProcessedCount++;
                        // Iterate through the list of parameters to clear them
                        foreach (ParameterInfo paramInfo in parametersToClearInfo)
                        {
                            Parameter targetParam = detailItem.get_Parameter(paramInfo.Guid);

                            if (targetParam != null && !targetParam.IsReadOnly)
                            {
                                try
                                {
                                    if (paramInfo.DataType == "TEXT")
                                    {
                                        if (targetParam.StorageType == StorageType.String)
                                        {
                                            targetParam.Set(string.Empty);
                                            parametersClearedCount++;
                                        }
                                    }
                                    else if (paramInfo.DataType == "YESNO")
                                    {
                                        if (targetParam.StorageType == StorageType.Integer)
                                        {
                                            targetParam.Set(0); // Set to No (0) to "nullify"
                                            parametersClearedCount++;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Log or inform user about specific parameter setting failure
                                    // For brevity, this example skips detailed per-parameter error logging
                                    System.Diagnostics.Debug.WriteLine($"Failed to set parameter {targetParam.Definition.Name} on element {detailItem.Id}: {ex.Message}");
                                }
                            }
                        }
                    }
                }
                tx.Commit();
            }

            TaskDialog.Show("Success", $"Processed {itemsProcessedCount} Detail Items. Cleared {parametersClearedCount} parameter instances.");
            return Result.Succeeded;
        }
    }
}