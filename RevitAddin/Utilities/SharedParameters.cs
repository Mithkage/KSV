// SharedParameters.cs

// --- FILE DESCRIPTION ---
// This file contains the logic for managing shared parameters in Revit.
// It programmatically generates a shared parameter file using data from the
// SharedParameterData class, binds parameters to a project, and provides
// helper methods for setting parameter values. This file no longer contains
// hardcoded parameter definitions; they are now in SharedParameterFile.cs.

// --- CHANGE LOG ---
// - [2025-08-13]: Changed CreateSharedParameterFile to public to allow external access.
// - [2025-08-12]: Merged old SharedParameters.cs with the SharedParameterHelper script.
// - Replaced hardcoded GUIDs in static classes with a structured list of objects.
// - Implemented methods to programmatically create the shared parameter file.
// - Added comprehensive error handling for file I/O and Revit API calls.
// - Added a new public method, GetParameterGuidByName, to retrieve a GUID by its name.
// - Removed the old GetRTS_CableGUIDByIndex method and RTS_Cable_GUIDs list.
// - [APPLIED FIX]: Removed duplicate parameter definitions from the list.
// - [APPLIED FIX]: Corrected the format of the generated shared parameter .txt file.
// - [APPLIED FIX]: Implemented a check to prevent errors when binding already-existing parameters.
// - [VERSION UPDATE]: Added conditional compilation directives for Revit 2022 and Revit 2024 support.
// - [CONTENT UPDATE]: Sorted parameters alphabetically within each group and added AI-generated descriptions.
// - [API FIX]: Resolved compiler errors for Revit 2022/2024 API differences.
// - [COMPATIBILITY FIX]: Added a backward-compatibility layer to support legacy static property access.
// - [UPDATE]: Added 'IsInstance' property to SharedParameterDefinition to control binding as Type or Instance.
// - [UPDATE]: Added SetParameterValue and SetTypeParameterValue helper methods for modifying parameter values.
// - [REFACTOR]: Separated parameter data definitions into SharedParameterFile.cs. This file now only contains logic.

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using static RTS.Utilities.SharedParameterData; // Import the static data class

namespace RTS.Utilities
{
    /// <summary>
    /// Manages shared parameter definitions and provides methods to programmatically
    /// create, apply, and modify shared parameters in a Revit project.
    /// </summary>
    public static class SharedParameters
    {
        /// <summary>
        /// A public method to retrieve a parameter's GUID by its name.
        /// This replaces the functionality of the old GetRTS_CableGUIDByIndex method.
        /// </summary>
        /// <param name="paramName">The name of the parameter to find.</param>
        /// <returns>The Guid of the parameter.</returns>
        /// <exception cref="ArgumentException">Thrown when the parameter name is not found.</exception>
        public static Guid GetParameterGuidByName(string paramName)
        {
            var parameter = MySharedParameters.FirstOrDefault(p => p.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            if (parameter == null)
            {
                throw new ArgumentException($"Shared parameter with name '{paramName}' not found in the definition list.");
            }
            return parameter.Guid;
        }

        // --- Backward Compatibility Layer ---
        // This section re-creates the old static properties to avoid breaking
        // existing code that relies on them. They now fetch the GUIDs from the
        // central list, ensuring a single source of truth.

        private static string ToIdentifier(string name) => Regex.Replace(name, @"[^a-zA-Z0-9_]", "_");

        public static class General
        {
            public static Guid RTS_Appendix_Name => GetParameterGuidByName("RTS_Appendix_Name");
            public static Guid RTS_Approved => GetParameterGuidByName("RTS_Approved");
            public static Guid RTS_Approved_Date => GetParameterGuidByName("RTS_Approved_Date");
            public static Guid RTS_Bell_Diameter => GetParameterGuidByName("RTS_Bell_Diameter");
            public static Guid RTS_Building => GetParameterGuidByName("RTS_Building");
            public static Guid RTS_Building_Class => GetParameterGuidByName("RTS_Building_Class");
            public static Guid RTS_Cable_17 => GetParameterGuidByName("RTS_Cable_17");
            public static Guid RTS_Checked => GetParameterGuidByName("RTS_Checked");
            public static Guid RTS_Checked_Date => GetParameterGuidByName("RTS_Checked_Date");
            public static Guid RTS_Code => GetParameterGuidByName("RTS_Code");
            public static Guid RTS_Comment => GetParameterGuidByName("RTS_Comment");
            public static Guid RTS_CopyReference_ID => GetParameterGuidByName("RTS_CopyReference_ID");
            public static Guid RTS_CopyRelative_Position => GetParameterGuidByName("RTS_CopyRelative_Position");
            public static Guid RTS_CopyReference_Type => GetParameterGuidByName("RTS_CopyReference_Type");
            public static Guid RTS_Created_By => GetParameterGuidByName("RTS_Created_By");
            public static Guid RTS_Depth => GetParameterGuidByName("RTS_Depth");
            public static Guid RTS_Design_Date => GetParameterGuidByName("RTS_Design_Date");
            public static Guid RTS_Designation => GetParameterGuidByName("RTS_Designation");
            public static Guid RTS_Diameter => GetParameterGuidByName("RTS_Diameter");
            public static Guid RTS_Discipline => GetParameterGuidByName("RTS_Discipline");
            public static Guid RTS_Drawn => GetParameterGuidByName("RTS_Drawn");
            public static Guid RTS_Elevation_From_Level => GetParameterGuidByName("RTS_Elevation_From_Level");
            public static Guid RTS_Error => GetParameterGuidByName("RTS_Error");
            public static Guid RTS_Filled_Region_Visible => GetParameterGuidByName("RTS_Filled_Region_Visible");
            public static Guid RTS_Filter => GetParameterGuidByName("RTS_Filter");
            public static Guid RTS_FRR => GetParameterGuidByName("RTS_FRR");
            public static Guid RTS_Grid_Reference => GetParameterGuidByName("RTS_Grid_Reference");
            public static Guid RTS_Height => GetParameterGuidByName("RTS_Height");
            public static Guid RTS_ID => GetParameterGuidByName("RTS_ID");
            public static Guid RTS_ITR => GetParameterGuidByName("RTS_ITR");
            public static Guid RTS_Item_Number => GetParameterGuidByName("RTS_Item_Number");
            public static Guid RTS_Left_End_Detail => GetParameterGuidByName("RTS_Left_End_Detail");
            public static Guid RTS_Level => GetParameterGuidByName("RTS_Level");
            public static Guid RTS_Location => GetParameterGuidByName("RTS_Location");
            public static Guid RTS_Mass => GetParameterGuidByName("RTS_Mass");
            public static Guid RTS_Material => GetParameterGuidByName("RTS_Material");
            public static Guid RTS_MaterialSecondary => GetParameterGuidByName("RTS_MaterialSecondary");
            public static Guid RTS_Meter_Brand => GetParameterGuidByName("RTS_Meter_Brand");
            public static Guid RTS_Meter_Description => GetParameterGuidByName("RTS_Meter_Description");
            public static Guid RTS_Meter_Number => GetParameterGuidByName("RTS_Meter_Number");
            public static Guid RTS_Meter_Supplier => GetParameterGuidByName("RTS_Meter_Supplier");
            public static Guid RTS_Model_Manager => GetParameterGuidByName("RTS_Model_Manager");
            public static Guid RTS_Note => GetParameterGuidByName("RTS_Note");
            public static Guid RTS_Pipe_Loading_Cap___Drain => GetParameterGuidByName("RTS_Pipe_Loading_Cap_-_Drain");
            public static Guid RTS_QA_ASBUILT => GetParameterGuidByName("RTS_QA_ASBUILT");
            public static Guid RTS_QA_ARC => GetParameterGuidByName("RTS_QA_ARC");
            public static Guid RTS_QA_ELEC => GetParameterGuidByName("RTS_QA_ELEC");
            public static Guid RTS_QA_STRUC => GetParameterGuidByName("RTS_QA_STRUC");
            public static Guid RTS_QA_VIEW_RANGE => GetParameterGuidByName("RTS_QA_VIEW RANGE");
            public static Guid RTS_Radius => GetParameterGuidByName("RTS_Radius");
            public static Guid RTS_Revision => GetParameterGuidByName("RTS_Revision");
            public static Guid RTS_Sample_Number => GetParameterGuidByName("RTS_Sample_Number");
            public static Guid RTS_Service_Type => GetParameterGuidByName("RTS_Service_Type");
            public static Guid RTS_Sheet_Scale => GetParameterGuidByName("RTS_Sheet_Scale");
            public static Guid RTS_Sheet_Series => GetParameterGuidByName("RTS_Sheet_Series");
            public static Guid RTS_Status => GetParameterGuidByName("RTS_Status");
            public static Guid RTS_Stock_Location => GetParameterGuidByName("RTS_Stock_Location");
            public static Guid RTS_Sub_Discipline => GetParameterGuidByName("RTS_Sub_Discipline");
            public static Guid RTS_Survey_Point => GetParameterGuidByName("RTS_Survey Point");
            public static Guid RTS_Thickness => GetParameterGuidByName("RTS_Thickness");
            public static Guid RTS_Verified_By => GetParameterGuidByName("RTS_Verified_By");
            public static Guid RTS_Verifier => GetParameterGuidByName("RTS_Verifier");
            public static Guid RTS_Version_No_ => GetParameterGuidByName("RTS_Version_No.");
            public static Guid RTS_View_Group => GetParameterGuidByName("RTS_View_Group");
            public static Guid RTS_Watermark => GetParameterGuidByName("RTS_Watermark");
            public static Guid RTS_Width => GetParameterGuidByName("RTS_Width");
            public static Guid RTS_Zone => GetParameterGuidByName("RTS_Zone");
        }

        public static class Cable
        {
            public static Guid Cables_On_Tray => GetParameterGuidByName("Cables On Tray");
            public static Guid Conduit_Number => GetParameterGuidByName("Conduit Number");
            public static Guid Fire_Saftey => GetParameterGuidByName("Fire Saftey");
            public static Guid RT_Cables_Weight => GetParameterGuidByName("RT_Cables Weight");
            public static Guid RT_Tray_Occupancy => GetParameterGuidByName("RT_Tray Occupancy");
            public static Guid RT_Tray_Min_Size => GetParameterGuidByName("RT_Tray Min Size");
            public static Guid RTS_Branch_Number => GetParameterGuidByName("RTS_Branch Number");
            public static Guid RTS_Cable_01 => GetParameterGuidByName("RTS_Cable_01");
            public static Guid RTS_Cable_02 => GetParameterGuidByName("RTS_Cable_02");
            public static Guid RTS_Cable_03 => GetParameterGuidByName("RTS_Cable_03");
            public static Guid RTS_Cable_04 => GetParameterGuidByName("RTS_Cable_04");
            public static Guid RTS_Cable_05 => GetParameterGuidByName("RTS_Cable_05");
            public static Guid RTS_Cable_06 => GetParameterGuidByName("RTS_Cable_06");
            public static Guid RTS_Cable_07 => GetParameterGuidByName("RTS_Cable_07");
            public static Guid RTS_Cable_08 => GetParameterGuidByName("RTS_Cable_08");
            public static Guid RTS_Cable_09 => GetParameterGuidByName("RTS_Cable_09");
            public static Guid RTS_Cable_10 => GetParameterGuidByName("RTS_Cable_10");
            public static Guid RTS_Cable_11 => GetParameterGuidByName("RTS_Cable_11");
            public static Guid RTS_Cable_12 => GetParameterGuidByName("RTS_Cable_12");
            public static Guid RTS_Cable_13 => GetParameterGuidByName("RTS_Cable_13");
            public static Guid RTS_Cable_14 => GetParameterGuidByName("RTS_Cable_14");
            public static Guid RTS_Cable_15 => GetParameterGuidByName("RTS_Cable_15");
            public static Guid RTS_Cable_16 => GetParameterGuidByName("RTS_Cable_16");
            public static Guid RTS_Cable_18 => GetParameterGuidByName("RTS_Cable_18");
            public static Guid RTS_Cable_19 => GetParameterGuidByName("RTS_Cable_19");
            public static Guid RTS_Cable_20 => GetParameterGuidByName("RTS_Cable_20");
            public static Guid RTS_Cable_21 => GetParameterGuidByName("RTS_Cable_21");
            public static Guid RTS_Cable_22 => GetParameterGuidByName("RTS_Cable_22");
            public static Guid RTS_Cable_23 => GetParameterGuidByName("RTS_Cable_23");
            public static Guid RTS_Cable_24 => GetParameterGuidByName("RTS_Cable_24");
            public static Guid RTS_Cable_25 => GetParameterGuidByName("RTS_Cable_25");
            public static Guid RTS_Cable_26 => GetParameterGuidByName("RTS_Cable_26");
            public static Guid RTS_Cable_27 => GetParameterGuidByName("RTS_Cable_27");
            public static Guid RTS_Cable_28 => GetParameterGuidByName("RTS_Cable_28");
            public static Guid RTS_Cable_29 => GetParameterGuidByName("RTS_Cable_29");
            public static Guid RTS_Cable_30 => GetParameterGuidByName("RTS_Cable_30");
            public static Guid RTS_Variance => GetParameterGuidByName("RTS_Variance");
            public static Guid Section_Tag => GetParameterGuidByName("Section Tag");
            public static Guid Spare_Capacity => GetParameterGuidByName("Spare Capacity");
            public static Guid Spare_Capacity_on_Circuit => GetParameterGuidByName("Spare Capacity on Circuit");
            public static Guid String_Supply => GetParameterGuidByName("String Supply");

            public static List<Guid> AllCableGuids => MySharedParameters
                .Where(p => p.Name.StartsWith("RTS_Cable_", StringComparison.OrdinalIgnoreCase))
                .Select(p => p.Guid)
                .ToList();
        }

        public static class PowerCAD
        {
            public static Guid PC_of_Phases => GetParameterGuidByName("PC_# of Phases");
            public static Guid PC_Cable_Size___Active_conductors => GetParameterGuidByName("PC_Cable Size - Active conductors");
            public static Guid PC_Cable_Size___Earthing_conductor => GetParameterGuidByName("PC_Cable Size - Earthing conductor");
            public static Guid PC_Clearance_Time_sec => GetParameterGuidByName("PC_Clearance Time (sec)");
            public static Guid PC_Conductor_Gap_mm => GetParameterGuidByName("PC_Conductor Gap (mm)");
            public static Guid PC_Cores => GetParameterGuidByName("PC_Cores");
            public static Guid PC_Design_Progress => GetParameterGuidByName("PC_Design Progress");
            public static Guid PC_Earth_Conductor_material => GetParameterGuidByName("PC_Earth Conductor material");
            public static Guid PC_Electrode_Configuration => GetParameterGuidByName("PC_Electrode Configuration");
            public static Guid PC_Enclosure_Depth_mm => GetParameterGuidByName("PC_Enclosure Depth (mm)");
            public static Guid PC_Enclosure_Width_mm => GetParameterGuidByName("PC_Enclosure Width (mm)");
            public static Guid PC_Isolator_Type => GetParameterGuidByName("PC_Isolator Type");
            public static Guid PC_No_of_Conduits => GetParameterGuidByName("PC_No. of Conduits");
            public static Guid PC_Prospective_Fault_3ø_Isc_kA => GetParameterGuidByName("PC_Prospective Fault 3ø Isc (kA)");
            public static Guid PC_Protective_Device_Breaking_Capacity_kA => GetParameterGuidByName("PC_Protective Device Breaking Capacity (kA)");
            public static Guid PC_Protective_Device_Description => GetParameterGuidByName("PC_Protective Device Description");
            public static Guid PC_Protective_Device_Manufacturer => GetParameterGuidByName("PC_Protective Device Manufacturer");
            public static Guid PC_Protective_Device_OCR_Trip_Unit => GetParameterGuidByName("PC_Protective Device OCR/Trip Unit");
            public static Guid PC_Protective_Device_Settings => GetParameterGuidByName("PC_Protective Device Settings");
            public static Guid PC_SWB_From => GetParameterGuidByName("PC_SWB From");
            public static Guid PC_SWB_Load => GetParameterGuidByName("PC_SWB Load");
            public static Guid PC_SWB_Load_Scope => GetParameterGuidByName("PC_SWB Load Scope");
            public static Guid PC_SWB_To => GetParameterGuidByName("PC_SWB To");
            public static Guid PC_SWB_Type => GetParameterGuidByName("PC_SWB Type");
            public static Guid PC_Active_Conductor_material => GetParameterGuidByName("PC_Active Conductor material");
            public static Guid Protective_Device_Rating_A => GetParameterGuidByName("Protective Device Rating (A)");
        }

        // --- Core Methods ---

        public static void AddMyParametersToProject(Document doc)
        {
            Application app = doc.Application;
            string originalSharedFilePath = app.SharedParametersFilename;
            string tempFilePath = string.Empty;

            try
            {
                tempFilePath = Path.Combine(Path.GetTempPath(), "RTS_Parameters_" + Path.GetRandomFileName() + ".txt");
                CreateSharedParameterFile(tempFilePath);

                app.SharedParametersFilename = tempFilePath;
                DefinitionFile sharedParamFile = app.OpenSharedParameterFile();
                if (sharedParamFile == null)
                {
                    throw new InvalidOperationException("Failed to open temporary shared parameter file.");
                }

                using (Transaction trans = new Transaction(doc, "Add RTS Parameters"))
                {
                    trans.Start();
                    try
                    {
                        BindingMap bindingMap = doc.ParameterBindings;
                        CategorySet cats = app.Create.NewCategorySet();
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalEquipment));
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Conduit));
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CableTray));
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_LightingFixtures));
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));

#if REVIT2024_OR_GREATER
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_ConduitFitting));
                        cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CableTrayFitting));
#endif

                        foreach (var paramDef in MySharedParameters)
                        {
                            DefinitionGroup group = sharedParamFile.Groups.get_Item(
                                MySharedParameterGroups.First(g => g.Id == paramDef.GroupId).Name);
                            if (group == null) continue;

                            ExternalDefinition externalDef = group.Definitions.get_Item(paramDef.Name) as ExternalDefinition;
                            if (externalDef == null) continue;
                            if (bindingMap.Contains(externalDef)) continue;

                            string uiGroupName = MySharedParameterGroups.First(g => g.Id == paramDef.GroupId).Name;
                            Binding binding = paramDef.IsInstance ? (Binding)app.Create.NewInstanceBinding(cats) : app.Create.NewTypeBinding(cats);

#if REVIT2024_OR_GREATER
                            ForgeTypeId groupTypeId = GetParameterGroupTypeId(uiGroupName);
                            bindingMap.Insert(externalDef, binding, groupTypeId);
#else
#pragma warning disable CS0618
                            BuiltInParameterGroup bipg = GetBuiltInParameterGroup(uiGroupName);
                            bindingMap.Insert(externalDef, binding, bipg);
#pragma warning restore CS0618
#endif
                        }
                        trans.Commit();
                    }
                    catch
                    {
                        if (trans.GetStatus() == TransactionStatus.Started) trans.RollBack();
                        throw;
                    }
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(originalSharedFilePath))
                {
                    try { app.SharedParametersFilename = originalSharedFilePath; } catch { }
                }
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { }
                }
            }
        }

        /// <summary>
        /// Generates a valid Revit shared parameter .txt file from the internal parameter definitions.
        /// Includes error handling for file I/O operations.
        /// </summary>
        /// <param name="filePath">The full path where the file will be created.</param>
        public static void CreateSharedParameterFile(string filePath)
        {
            try
            {
                var sortedParameters = MySharedParameters.OrderBy(p => p.GroupId).ThenBy(p => p.Name).ToList();
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    writer.WriteLine("# This is a Revit shared parameter file.");
                    writer.WriteLine("*META\tVERSION\tMINVERSION");
                    writer.WriteLine("META\t2\t1");
                    writer.WriteLine("*GROUP\tID\tNAME");
                    foreach (var group in MySharedParameterGroups)
                    {
                        writer.WriteLine($"GROUP\t{group.Id}\t{group.Name}");
                    }
                    writer.WriteLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE");
                    foreach (var paramDef in sortedParameters)
                    {
                        string datatype = paramDef.DataType.ToString().ToUpper();
                        string description = paramDef.Description ?? "";
                        string userModifiable = paramDef.UserModifiable ? "1" : "0";
                        string visible = paramDef.Visible ? "1" : "0";
                        string hideWhenNoValue = paramDef.HideWhenNoValue ? "1" : "0";
                        writer.WriteLine($"PARAM\t{paramDef.Guid}\t{paramDef.Name}\t{datatype}\t\t{paramDef.GroupId}\t{visible}\t{description}\t{userModifiable}\t{hideWhenNoValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create the temporary shared parameter file.", ex);
            }
        }

#if REVIT2024_OR_GREATER
        private static ForgeTypeId GetParameterGroupTypeId(string groupName)
        {
            switch (groupName)
            {
                case "_General": return GroupTypeId.General;
                case "RSGx": return GroupTypeId.Data;
                case "PowerCAD": return GroupTypeId.ElectricalAnalysis;
                default: return GroupTypeId.General;
            }
        }
#else
#pragma warning disable CS0618
        private static BuiltInParameterGroup GetBuiltInParameterGroup(string groupName)
        {
            switch (groupName)
            {
                case "_General": return BuiltInParameterGroup.PG_GENERAL;
                case "RSGx": return BuiltInParameterGroup.PG_DATA;
                case "PowerCAD": return BuiltInParameterGroup.PG_ELECTRICAL_ANALYSIS;
                default: return BuiltInParameterGroup.PG_GENERAL;
            }
        }
#pragma warning restore CS0618
#endif

        #region --- Parameter Setter Methods ---
        public static bool SetParameterValue(Element element, Guid paramGuid, object value)
        {
            if (element == null) return false;
            Parameter param = element.get_Parameter(paramGuid);
            if (param == null || param.IsReadOnly) return false;
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String: return param.Set(value as string);
                    case StorageType.Double:
                        if (value is int intVal) return param.Set(Convert.ToDouble(intVal));
                        if (value is double dblVal) return param.Set(dblVal);
                        break;
                    case StorageType.Integer:
                        if (value is bool boolVal) return param.Set(boolVal ? 1 : 0);
                        if (value is int intValue) return param.Set(intValue);
                        break;
                    case StorageType.ElementId:
                        if (value is ElementId idVal) return param.Set(idVal);
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to set parameter {param.Definition.Name}: {ex.Message}");
                return false;
            }
            return false;
        }

        public static bool SetTypeParameterValue(Element element, Guid paramGuid, object value)
        {
            if (element == null) return false;
            Element elementType = element.Document.GetElement(element.GetTypeId());
            if (elementType == null) return false;
            return SetParameterValue(elementType, paramGuid, value);
        }
        #endregion
    }
}
