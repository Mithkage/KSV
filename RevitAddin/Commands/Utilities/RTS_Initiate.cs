//
// File: RTS_Initiate.cs
//
// Namespace: RTS_Initiate
//
// Class: RTS_InitiateClass
//
// Function: Initiates shared parameters in a Revit 2022 and 2024 project for Electrical Detail Items, Cable/Conduit elements, and various electrical/lighting/communication categories, including Wires.
//
// Author: Kyle Vorster
//
// Date: June 18, 2024 (Updated June 29, 2025)
//

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.Support;
using System;
using System.Collections.Generic; // Required for List
using System.IO;
using System.Linq;
using System.Text; // Required for StringBuilder

namespace RTS.Commands.Utilities
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_InitiateClass : IExternalCommand
    {
        // Helper class to store parameter information
        private class SharedParameterInfo
        {
            public string Name { get; }
            public Guid Guid { get; }
            public string GuidString { get; }

            public SharedParameterInfo(string name, string guidString)
            {
                Name = name;
                GuidString = guidString;
                Guid = new Guid(guidString);
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // --- General Configuration ---
            string sharedParamFileNameContains = "RTS_Shared Parameters";
            BuiltInParameterGroup parameterGroup = BuiltInParameterGroup.PG_ELECTRICAL;

            // --- Configuration for Detail Item Parameters ---
            BuiltInCategory detailItemTargetCategory = BuiltInCategory.OST_DetailComponents;
            List<SharedParameterInfo> detailItemParametersToAdd = new List<SharedParameterInfo>
            {
                new SharedParameterInfo("PC_SWB To", "e142b0ed-d084-447a-991b-d9a3a3f67a8d"),
                new SharedParameterInfo("PC_# of Phases", "f5d583ec-f511-4653-828e-9b45281baa54"),
                new SharedParameterInfo("PC_Cable Length", "09007343-dd0a-44c3-b04d-145118778ac3"),
                new SharedParameterInfo("PC_Cable Reference", "da8a3228-00aa-4f2d-a472-1ba675284cef"),
                new SharedParameterInfo("PC_Isolator Type", "bc3f0ed3-e9f5-456f-83c6-d7f2dc0db169"),
                new SharedParameterInfo("PC_PowerCAD", "8f31d68f-60c9-4ec6-a7ff-78a6e3bdaab6"),
                new SharedParameterInfo("PC_Prospective Fault at End of Cable", "5b362b2d-ceec-42db-93a0-ca5ccb630e9e"),
                new SharedParameterInfo("PC_SWB From", "5dd52911-7bcd-4a06-869d-73fcef59951c"),
                new SharedParameterInfo("PC_SWB Load", "60f670e1-7d54-4ffc-b0f5-ed62c08d3b90"),
                new SharedParameterInfo("PC_SWB Type", "9d5ab9c2-09e2-4d42-ae2f-2a5fba6f7131")
            };

            // --- Configuration for Cable/Conduit Parameters ---
            List<BuiltInCategory> cableConduitTargetCategoriesEnum = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayRun, // Note: Runs are system elements, direct parameter binding might behave differently.
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_ConduitRun,    // Note: Runs are system elements, direct parameter binding might behave differently.
                BuiltInCategory.OST_Conduit // Corrected from OST_Conduits
            };
            List<SharedParameterInfo> cableConduitParametersToAdd = new List<SharedParameterInfo>
            {
                new SharedParameterInfo("RTS_Cable_01", "cf0d478e-1e98-4e83-ab80-6ee867f61798"),
                new SharedParameterInfo("RTS_Cable_02", "2551d308-44ed-405c-8aad-fb78624d086e"),
                new SharedParameterInfo("RTS_Cable_03", "c1dfc402-2101-4e53-8f52-f6af64584a9f"),
                new SharedParameterInfo("RTS_Cable_04", "f297daa6-a9e0-4dd5-bda3-c628db7c28bd"),
                new SharedParameterInfo("RTS_Cable_05", "b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"),
                new SharedParameterInfo("RTS_Cable_06", "7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"),
                new SharedParameterInfo("RTS_Cable_07", "9bc78bce-0d39-4538-b507-7b98e8a13404"),
                new SharedParameterInfo("RTS_Cable_08", "e9d50153-a0e9-4685-bc92-d89f244f7e8e"),
                new SharedParameterInfo("RTS_Cable_09", "5713d65a-91df-4d2e-97bf-1c3a10ea5225"),
                new SharedParameterInfo("RTS_Cable_10", "64af3105-b2fd-44bc-9ad3-17264049ff62"),
                new SharedParameterInfo("RTS_Cable_11", "f3626002-0e62-4b75-93cc-35d0b11dfd67"),
                new SharedParameterInfo("RTS_Cable_12", "63dc0a2e-0770-4002-a859-a9d40a2ce023"),
                new SharedParameterInfo("RTS_Cable_13", "eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"),
                new SharedParameterInfo("RTS_Cable_14", "0e0572e5-c568-42b7-8730-a97433bd9b54"),
                new SharedParameterInfo("RTS_Cable_15", "bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"),
                new SharedParameterInfo("RTS_Cable_16", "f6d2af67-027e-4b9c-9def-336ebaa87336"),
                new SharedParameterInfo("RTS_Cable_17", "f6a4459d-46a1-44c0-8545-ee44e4778854"),
                new SharedParameterInfo("RTS_Cable_18", "0d66d2fa-f261-4daa-8041-9eadeefac49a"),
                new SharedParameterInfo("RTS_Cable_19", "af483914-c8d2-4ce6-be6e-ab81661e5bf1"),
                new SharedParameterInfo("RTS_Cable_20", "c8d2d2fc-c248-483f-8d52-e630eb730cd7"),
                new SharedParameterInfo("RTS_Cable_21", "aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"),
                new SharedParameterInfo("RTS_Cable_22", "6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"),
                new SharedParameterInfo("RTS_Cable_23", "7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"),
                new SharedParameterInfo("RTS_Cable_24", "7f745b2b-a537-42d9-8838-7a5521cc7d0c"),
                new SharedParameterInfo("RTS_Cable_25", "9a76c2dc-1022-4a54-ab66-5ca625b50365"),
                new SharedParameterInfo("RTS_Cable_26", "658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"),
                new SharedParameterInfo("RTS_Cable_27", "8ad24640-036b-44d2-af9c-b891f6e64271"),
                new SharedParameterInfo("RTS_Cable_28", "c046c4d7-e1fd-4cf7-a99f-14ae96b722be"),
                new SharedParameterInfo("RTS_Cable_29", "cdf00587-7e11-4af4-8e54-48586481cf22"),
                new SharedParameterInfo("RTS_Cable_30", "a92bb0f9-2781-4971-a3b1-9c47d62b947b"),
                new SharedParameterInfo("RTS_Cables On Tray", "c7430aff-c4ee-4354-9601-a060364b43d5"),
                new SharedParameterInfo("Section Tag", "bc3d8d0a-9ee3-43fa-bca5-0bc414306316"),
                new SharedParameterInfo("Fire Saftey", "4e8047d8-023b-4ae9-ae96-1f871cf51f4e"),
                new SharedParameterInfo("Conduit Number", "d44ab6c4-aa8c-4abb-b2cb-16038885a7f9"),
                new SharedParameterInfo("Branch Number", "3ea1a3bb-8416-45ed-b606-3f3a3f87d4be"),
                new SharedParameterInfo("Cable Tray Tier", "cea7b3ba-f72c-403e-a6c0-0a414b793b9d"),
                new SharedParameterInfo("String Supply", "a9613056-877e-42bd-ad74-73707c1ad24e"),
                new SharedParameterInfo("RT_Tray Occupancy", "a6f087c7-cecc-4335-831b-249cb9398abc"),
                new SharedParameterInfo("RT_Cables Weight", "51d670fa-0338-42e7-ac9e-f2c44a56ffcc"),
                new SharedParameterInfo("RT_Tray Min Size", "5ed6b64c-af5c-4425-ab69-85a7fa5fdffe"),
                new SharedParameterInfo("RTS_Comment", "f8a844ce-cb1a-4d95-bb11-d48d15a84a8e")
            };

            // --- Configuration for RTS_ID Parameter for various categories ---
            List<BuiltInCategory> rtsIdTargetCategoriesEnum = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_Wire // Added Wires category for RTS_ID
            };

            // RTS_ID parameter definition
            List<SharedParameterInfo> rtsIdParametersToAdd = new List<SharedParameterInfo>
            {
                new SharedParameterInfo("RTS_ID", "3175a27e-d386-4567-bf10-2da1a9cbb73b")
            };

            // --- Configuration for PC_ Parameters for Wires category ---
            BuiltInCategory wireTargetCategory = BuiltInCategory.OST_Wire;
            List<SharedParameterInfo> wireParametersToAdd = new List<SharedParameterInfo>
            {
                new SharedParameterInfo("PC_Separate Earth for Multicore", "cbd9da23-c8eb-4233-af6b-6337c05c2f12"),
                new SharedParameterInfo("PC_PowerCAD", "8f31d68f-60c9-4ec6-a7ff-78a6e3bdaab6"),
                new SharedParameterInfo("PC_Nominal Overall Diameter", "98fa8f80-0219-4b99-bf3d-8da7c74f356d"),
                new SharedParameterInfo("PC_Earth Conductor material", "ee5c5f1a-7e6e-480b-9591-c1391cf0990b"),
                new SharedParameterInfo("PC_Design Progress", "5eb8da2d-a094-4f80-8626-9960fb6b4aa3"),
                new SharedParameterInfo("PC_Cores", "f3c6038d-7300-46a5-8787-14560678f531"),
                new SharedParameterInfo("PC_Cable Type", "5efc9508-f2e9-40ee-b353-d68429766a37"),
                new SharedParameterInfo("PC_Cable Size - Neutral conductors", "dbd8ce49-4c2b-4ce5-a3ba-fca4c28e3d15"),
                new SharedParameterInfo("PC_Cable Size - Earthing conductor", "50b6ef99-8f5e-42e1-8645-cce97f6b94b6"),
                new SharedParameterInfo("PC_Cable Size - Active conductors", "91c8321f-1342-4efa-b648-7ba5e95c0085"),
                new SharedParameterInfo("PC_Cable Reference", "da8a3228-00aa-4f2d-a472-1ba675284cef"),
                new SharedParameterInfo("PC_Cable Length", "09007343-dd0a-44c3-b04d-145118778ac3"),
                new SharedParameterInfo("PC_Cable Insulation", "dc328944-dcda-4612-8a2e-ecaed994c876"),
                new SharedParameterInfo("PC_Active Conductor material", "62ba1846-99f5-4a7f-985a-9255d54c5b93"),
                new SharedParameterInfo("PC_Accum Volt Drop Incl FSC", "c2631a30-511f-450a-92b6-b5c95274ae7b"),
                new SharedParameterInfo("PC_# of Phases", "f5d583ec-f511-4653-828e-9b45281baa54")
            };
            // --- End Configuration ---

            StringBuilder summaryMessage = new StringBuilder();
            summaryMessage.AppendLine("RTS Initiate Parameters - Processing Summary:");
            summaryMessage.AppendLine("--------------------------------------------");

            DefinitionFile sharedParamFile = null;

            try
            {
                // 1. Get and validate the Shared Parameter File
                summaryMessage.AppendLine("Step 1: Validating Shared Parameter File...");
                string currentSharedParamFile = app.SharedParametersFilename;
                if (string.IsNullOrEmpty(currentSharedParamFile) || !File.Exists(currentSharedParamFile))
                {
                    message = "No Shared Parameter file is currently set in Revit.";
                    CustomTaskDialog.Show("Error - Prerequisite Failed", message);
                    return Result.Failed;
                }

                if (!Path.GetFileName(currentSharedParamFile).Contains(sharedParamFileNameContains))
                {
                    message = $"The current Shared Parameter file '{Path.GetFileName(currentSharedParamFile)}' does not seem to be the correct one (expected name containing '{sharedParamFileNameContains}'). Please set the correct file and try again.";
                    CustomTaskDialog.Show("Error - Prerequisite Failed", message);
                    return Result.Failed;
                }
                summaryMessage.AppendLine("    -> Shared Parameter File is valid.");

                app.SharedParametersFilename = currentSharedParamFile;
                sharedParamFile = app.OpenSharedParameterFile();

                if (sharedParamFile == null)
                {
                    message = "Could not open the Shared Parameter file: " + currentSharedParamFile;
                    CustomTaskDialog.Show("Error - Prerequisite Failed", message);
                    return Result.Failed;
                }
                summaryMessage.AppendLine("    -> Shared Parameter File opened successfully.");

                // 2. Process Detail Item Parameters
                summaryMessage.AppendLine("\nStep 2: Processing Detail Item Parameters...");
                CategorySet detailItemCategories = app.Create.NewCategorySet();
                Category detailItemCatObj = doc.Settings.Categories.get_Item(detailItemTargetCategory);
                if (detailItemCatObj == null)
                {
                    summaryMessage.AppendLine($"ERROR: Could not find the '{detailItemTargetCategory.ToString()}' category in the project.");
                }
                else
                {
                    detailItemCategories.Insert(detailItemCatObj);
                    ProcessParameters(doc, app, sharedParamFile, detailItemParametersToAdd, detailItemCategories, parameterGroup, summaryMessage, "Detail Items");
                }

                // 3. Process Cable/Conduit Parameters
                summaryMessage.AppendLine("\nStep 3: Processing Cable/Conduit Parameters...");
                CategorySet cableConduitCategoriesSet = app.Create.NewCategorySet();
                bool allCableConduitCategoriesFound = true;
                foreach (BuiltInCategory catEnum in cableConduitTargetCategoriesEnum)
                {
                    Category catObj = doc.Settings.Categories.get_Item(catEnum);
                    if (catObj == null)
                    {
                        summaryMessage.AppendLine($"ERROR: Could not find the '{catEnum.ToString()}' category in the project.");
                        allCableConduitCategoriesFound = false;
                    }
                    else
                    {
                        cableConduitCategoriesSet.Insert(catObj);
                    }
                }

                if (cableConduitCategoriesSet.IsEmpty)
                {
                    summaryMessage.AppendLine("ERROR: No valid categories found or specified for Cable/Conduit parameters. Skipping this set.");
                }
                else
                {
                    if (!allCableConduitCategoriesFound)
                    {
                        summaryMessage.AppendLine("WARNING: Not all specified Cable/Conduit categories were found. Parameters will be applied to found categories only.");
                    }
                    ProcessParameters(doc, app, sharedParamFile, cableConduitParametersToAdd, cableConduitCategoriesSet, parameterGroup, summaryMessage, "Cable/Conduit");
                }

                // 4. Process RTS_ID Parameter for Electrical Equipment, Electrical Fixtures, Lighting Devices, Lighting Fixtures, Conduit Fittings, Conduits, Communication Devices, and Wires
                summaryMessage.AppendLine("\nStep 4: Processing RTS_ID Parameter for Electrical Equipment, Electrical Fixtures, Lighting Devices, Lighting Fixtures, Conduit Fittings, Conduits, Communication Devices, and Wires...");
                CategorySet rtsIdCategoriesSet = app.Create.NewCategorySet();
                bool allRTSIdCategoriesFound = true;
                foreach (BuiltInCategory catEnum in rtsIdTargetCategoriesEnum)
                {
                    Category catObj = doc.Settings.Categories.get_Item(catEnum);
                    if (catObj == null)
                    {
                        summaryMessage.AppendLine($"ERROR: Could not find the '{catEnum.ToString()}' category in the project.");
                        allRTSIdCategoriesFound = false;
                    }
                    else
                    {
                        rtsIdCategoriesSet.Insert(catObj);
                    }
                }

                if (rtsIdCategoriesSet.IsEmpty)
                {
                    summaryMessage.AppendLine("ERROR: No valid categories found or specified for RTS_ID parameters. Skipping this set.");
                }
                else
                {
                    if (!allRTSIdCategoriesFound)
                    {
                        summaryMessage.AppendLine("WARNING: Not all specified RTS_ID categories were found. Parameters will be applied to found categories only.");
                    }
                    ProcessParameters(doc, app, sharedParamFile, rtsIdParametersToAdd, rtsIdCategoriesSet, parameterGroup, summaryMessage, "RTS_ID Categories");
                }

                // 5. Process PC_ Parameters for Wires Category
                summaryMessage.AppendLine("\nStep 5: Processing PC_ Parameters for Wires Category...");
                CategorySet wireCategoriesSet = app.Create.NewCategorySet();
                Category wireCatObj = doc.Settings.Categories.get_Item(wireTargetCategory);
                if (wireCatObj == null)
                {
                    summaryMessage.AppendLine($"ERROR: Could not find the '{wireTargetCategory.ToString()}' category in the project.");
                }
                else
                {
                    wireCategoriesSet.Insert(wireCatObj);
                    ProcessParameters(doc, app, sharedParamFile, wireParametersToAdd, wireCategoriesSet, parameterGroup, summaryMessage, "Wires PC_ Parameters");
                }

                CustomTaskDialog.Show("RTS Initiate Parameters - Results", summaryMessage.ToString());
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "A critical unexpected error occurred. See the summary for details.";
                summaryMessage.AppendLine($"\nCRITICAL ERROR in main execution block.");
                LogExceptionDetails(ex, summaryMessage, "Main Execute Block");
                CustomTaskDialog.Show("RTS Initiate Parameters - Critical Error", summaryMessage.ToString());
                return Result.Failed;
            }
        }

        /// <summary>
        /// Processes a list of shared parameters, binding them to the specified categories.
        /// </summary>
        private void ProcessParameters(Document doc, Application app, DefinitionFile sharedParamFile,
                               List<SharedParameterInfo> parametersToProcess,
                               CategorySet targetCategories,
                               BuiltInParameterGroup parameterGroup,
                               StringBuilder summaryMessage,
                               string categorySetNameForMessages)
        {
            if (targetCategories.IsEmpty)
            {
                summaryMessage.AppendLine($"INFO: No categories provided for '{categorySetNameForMessages}', skipping.");
                return;
            }

            BindingMap bindingMap = doc.ParameterBindings;

            foreach (SharedParameterInfo paramInfo in parametersToProcess)
            {
                Transaction t = null;
                try
                {
                    summaryMessage.AppendLine($"\nProcessing '{paramInfo.Name}' for {categorySetNameForMessages}...");

                    // Find the definition in the shared parameter file.
                    ExternalDefinition externalDefinition = sharedParamFile.Groups
                        .SelectMany(g => g.Definitions)
                        .OfType<ExternalDefinition>()
                        .FirstOrDefault(d => d.GUID == paramInfo.Guid);

                    if (externalDefinition == null)
                    {
                        summaryMessage.AppendLine($"    -> ERROR: Definition not found in shared parameter file for GUID: {paramInfo.GuidString}. Please ensure the shared parameter file is correct and contains this parameter.");
                        continue;
                    }

                    // Start a single transaction for this parameter.
                    t = new Transaction(doc, $"Bind/Update Parameter: {paramInfo.Name}");
                    t.Start();

                    bool bindingSuccess = false;
                    InstanceBinding existingBinding = bindingMap.get_Item(externalDefinition) as InstanceBinding;

                    if (existingBinding != null)
                    {
                        // A binding for this parameter already exists. Check and update it.
                        bool needsRebind = false;
                        CategorySet categoriesToInsert = app.Create.NewCategorySet();
                        foreach (Category cat in targetCategories)
                        {
                            if (!existingBinding.Categories.Contains(cat))
                            {
                                categoriesToInsert.Insert(cat);
                                needsRebind = true;
                            }
                            else
                            {
                                categoriesToInsert.Insert(cat); // Keep existing categories in the new set
                            }
                        }

                        if (needsRebind)
                        {
                            summaryMessage.AppendLine($"    -> INFO: Existing binding found. Updating categories...");
                            // ReInsert implicitly updates the categories for an existing binding
                            if (bindingMap.ReInsert(externalDefinition, app.Create.NewInstanceBinding(categoriesToInsert), parameterGroup))
                            {
                                summaryMessage.AppendLine($"      -> SUCCESS: Updated binding to include new categories.");
                                bindingSuccess = true;
                            }
                            else
                            {
                                summaryMessage.AppendLine($"      -> ERROR: Failed to update (ReInsert) binding.");
                            }
                        }
                        else
                        {
                            summaryMessage.AppendLine($"    -> INFO: Existing binding is already correct. No changes needed.");
                            bindingSuccess = true; // No action needed, but consider it a success.
                        }
                    }
                    else
                    {
                        // The parameter does not have an existing binding. Create a new one.
                        summaryMessage.AppendLine($"    -> INFO: No existing binding found. Creating new one...");
                        InstanceBinding newBinding = app.Create.NewInstanceBinding(targetCategories);
                        if (bindingMap.Insert(externalDefinition, newBinding, parameterGroup))
                        {
                            summaryMessage.AppendLine($"      -> SUCCESS: Created new binding.");
                            bindingSuccess = true;
                        }
                        else
                        {
                            summaryMessage.AppendLine($"      -> ERROR: Failed to create (Insert) new binding.");
                        }
                    }

                    if (bindingSuccess)
                    {
                        // Set 'Vary by Group' property only if it's an instance parameter.
                        // Type parameters do not support this property and will throw an ArgumentException.
                        SharedParameterElement spElement = SharedParameterElement.Lookup(doc, paramInfo.Guid);
                        if (spElement != null)
                        {
                            InternalDefinition internalDef = spElement.GetDefinition();
                            if (internalDef != null)
                            {
                                if (internalDef.Name != paramInfo.Name)
                                {
                                    summaryMessage.AppendLine($"      -> WARNING: In-project name is '{internalDef.Name}'; shared file name is '{paramInfo.Name}'.");
                                }

                                // Use GetDataType() instead of ParameterType to check for instance parameter compatibility
                                // SpecTypeId.Text for text, but you'll generally check if it's NOT a type parameter data type
                                // GetDataType() returns a ForgeTypeId
                                ForgeTypeId dataType = internalDef.GetDataType();

                                // Check if it's an instance parameter. A simple heuristic is that type parameters
                                // typically have a specific data type related to type-level properties,
                                // while many instance parameters are more general (like text, number, etc.)
                                // If SetAllowVaryBetweenGroups throws an error, it's definitely a type parameter.
                                try
                                {
                                    // This line will still throw if it's a Type Parameter,
                                    // but we're now attempting it based on the new API recommendation
                                    // and catching the specific ArgumentException if it fails.
                                    internalDef.SetAllowVaryBetweenGroups(doc, true);
                                    summaryMessage.AppendLine("      -> INFO: 'Vary by Group' property set.");
                                }
                                catch (ArgumentException exSetVary)
                                {
                                    // This catch block is for when SetAllowVaryBetweenGroups truly isn't applicable
                                    summaryMessage.AppendLine($"      -> WARNING: Could not set 'Vary by Group' for '{paramInfo.Name}'. This parameter might be a Type Parameter, which does not support this property. (Error: {exSetVary.Message})");
                                }
                            }
                        }
                    }

                    if (t.GetStatus() == TransactionStatus.Started)
                    {
                        t.Commit();
                    }
                }
                catch (Exception exParam)
                {
                    string context = $"Processing parameter '{paramInfo.Name}' for '{categorySetNameForMessages}'";
                    summaryMessage.AppendLine($"\nCRITICAL ERROR: An exception was thrown while {context}.");
                    LogExceptionDetails(exParam, summaryMessage, context);

                    if (t != null && t.GetStatus() == TransactionStatus.Started)
                    {
                        t.RollBack();
                        summaryMessage.AppendLine("    -> INFO: Transaction was rolled back.");
                    }
                }
            }
        }

        /// <summary>
        /// Creates a detailed log of an exception, including inner exceptions and stack trace.
        /// </summary>
        private void LogExceptionDetails(Exception ex, StringBuilder sb, string context = "")
        {
            sb.AppendLine($"\n--- DETAILED ERROR LOG ---");
            if (!string.IsNullOrEmpty(context))
            {
                sb.AppendLine($"Context: {context}");
            }
            sb.AppendLine($"Exception Type: {ex.GetType().Name}");
            sb.AppendLine($"Message: {ex.Message}");
            sb.AppendLine($"Stack Trace:\n{ex.StackTrace}");

            Exception inner = ex.InnerException;
            int level = 1;
            while (inner != null)
            {
                sb.AppendLine($"\n--- Inner Exception (Level {level}) ---");
                sb.AppendLine($"Type: {inner.GetType().Name}");
                sb.AppendLine($"Message: {inner.Message}");
                sb.AppendLine($"Stack Trace:\n{inner.StackTrace}");
                inner = inner.InnerException;
                level++;
            }
            sb.AppendLine("--- END OF DETAILED ERROR LOG ---");
        }
    }
}