//
// File: RTS_Initiate.cs
//
// Namespace: RTS.Commands
//
// Class: RTS_InitiateClass
//
// Function: Initiates shared parameters in a Revit project by calling the centralized
//           SharedParameters and ScheduleManager utilities.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Refactored to use the centralized ScheduleManager for all schedule-related parameter initiation.
// - [APPLIED FIX]: Added a granular option to initiate parameters for the CopyRelative script in families.
//
// Author: Kyle Vorster
//
#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI;
using RTS.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#endregion

namespace RTS.Commands.ProjectSetup
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_InitiateClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            var selectionWindow = new InitiateParametersWindow(doc);
            bool? dialogResult = selectionWindow.ShowDialog();

            if (dialogResult != true)
            {
                return Result.Cancelled;
            }

            bool anyActionTaken = false;
            var reportBuilder = new StringBuilder();
            reportBuilder.AppendLine("RTS Initiate Parameters - Processing Summary:");
            reportBuilder.AppendLine("--------------------------------------------");

            // --- Family-Specific Actions ---
            if (selectionWindow.AddRtsTypeToFamily)
            {
                AddParametersToFamily(commandData, new List<string> { "RTS_Type" }, "RTS_Type", reportBuilder);
                anyActionTaken = true;
            }
            if (selectionWindow.InitiateCopyRelativeParameters)
            {
                var copyRelativeParams = new List<string> { "RTS_CopyReference_ID", "RTS_CopyRelative_Position", "RTS_CopyReference_Type" };
                AddParametersToFamily(commandData, copyRelativeParams, "Copy Relative", reportBuilder);
                anyActionTaken = true;
            }

            // --- Project-Specific Actions ---
            if (selectionWindow.InitiateAllProjectParameters)
            {
                InitiateProjectParameters(commandData, reportBuilder);
                anyActionTaken = true;
            }

            if (selectionWindow.InitiateMapCablesParameters)
            {
                InitiateMapCablesParameters(commandData, reportBuilder);
                anyActionTaken = true;
            }

            if (selectionWindow.InitiateScheduleParameters)
            {
                InitiateScheduleParameters(commandData, reportBuilder, selectionWindow);
                anyActionTaken = true;
            }

            if (!anyActionTaken)
            {
                TaskDialog.Show("Information", "No action was selected.");
            }
            else
            {
                TaskDialog.Show("Initiate Parameters Complete", reportBuilder.ToString());
            }

            return Result.Succeeded;
        }

        private void InitiateProjectParameters(ExternalCommandData commandData, StringBuilder report)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            try
            {
                SharedParameters.AddMyParametersToProject(doc);
                report.AppendLine("\n- Initiated all RTS Shared Parameters in the project.");
            }
            catch (Exception ex)
            {
                report.AppendLine($"\n- ERROR initiating all project parameters: {ex.Message}");
            }
        }

        private void InitiateMapCablesParameters(ExternalCommandData commandData, StringBuilder report)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            var mapCablesParamGuids = new List<Guid>();
            mapCablesParamGuids.AddRange(SharedParameters.Cable.AllCableGuids);
            mapCablesParamGuids.Add(SharedParameters.General.RTS_Note);

            var targetCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            BindSpecificParameters(doc, mapCablesParamGuids, targetCategories, "Cable Mapping", report);
        }

        private void InitiateScheduleParameters(ExternalCommandData commandData, StringBuilder report, InitiateParametersWindow selections)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            var paramGuids = new HashSet<Guid>();
            var targetCategories = new HashSet<BuiltInCategory>();

            // Use a helper to map selections to schedule definitions from the manager
            Action<bool, string> processScheduleSelection = (isSelected, scheduleName) =>
            {
                if (isSelected)
                {
                    var scheduleDef = ScheduleManager.StandardSchedules.FirstOrDefault(s => s.Name == scheduleName);
                    if (scheduleDef != null)
                    {
                        foreach (var guid in scheduleDef.RequiredSharedParameterGuids)
                        {
                            paramGuids.Add(guid);
                        }
                        if (scheduleDef.Category != BuiltInCategory.INVALID)
                        {
                            targetCategories.Add(scheduleDef.Category);
                        }
                    }
                }
            };

            processScheduleSelection(selections.InitiateScheduleDetailItems, "RTS_Sc_Detail Items");
            processScheduleSelection(selections.InitiateScheduleElecEquip, "RTS_Sc_Electrical Equipment");
            processScheduleSelection(selections.InitiateScheduleLighting, "RTS_Sc_Lighting Fixtures");
            processScheduleSelection(selections.InitiateScheduleElecFixtures, "RTS_Sc_Electrical Fixtures");
            processScheduleSelection(selections.InitiateScheduleCableTrays, "RTS_Sc_Cable Trays");
            processScheduleSelection(selections.InitiateScheduleCableTrays, "RTS_Sc_Cable Tray Fittings");
            processScheduleSelection(selections.InitiateScheduleCableTrays, "RTS_Sc_Conduits");


            if (!paramGuids.Any())
            {
                report.AppendLine("\n- No schedule types were selected to initiate parameters.");
                return;
            }

            BindSpecificParameters(doc, paramGuids.ToList(), targetCategories.ToList(), "Standard Schedules", report);
        }


        private void BindSpecificParameters(Document doc, List<Guid> paramGuids, List<BuiltInCategory> builtInCategories, string processName, StringBuilder report)
        {
            Application app = doc.Application;
            var targetCategories = new CategorySet();
            foreach (var bic in builtInCategories)
            {
                Category cat = doc.Settings.Categories.get_Item(bic);
                if (cat != null)
                {
                    targetCategories.Insert(cat);
                }
            }

            if (targetCategories.IsEmpty)
            {
                report.AppendLine($"\n- ERROR for '{processName}': Could not find the necessary categories in this project.");
                return;
            }

            try
            {
                string sharedParamFilepath = app.SharedParametersFilename;
                if (string.IsNullOrEmpty(sharedParamFilepath) || !File.Exists(sharedParamFilepath))
                {
                    report.AppendLine($"\n- ERROR for '{processName}': A valid Shared Parameter file is not configured.");
                    return;
                }

                app.SharedParametersFilename = sharedParamFilepath;
                DefinitionFile sharedParamFile = app.OpenSharedParameterFile();
                if (sharedParamFile == null) throw new InvalidOperationException("Failed to open shared parameter file.");

                using (var tx = new Transaction(doc, $"Initiate Parameters for {processName}"))
                {
                    tx.Start();
                    BindingMap bindingMap = doc.ParameterBindings;
                    int boundCount = 0;
                    int alreadyBoundCount = 0;

                    foreach (Guid paramGuid in paramGuids)
                    {
                        ExternalDefinition def = sharedParamFile.Groups
                            .SelectMany(g => g.Definitions)
                            .OfType<ExternalDefinition>()
                            .FirstOrDefault(d => d.GUID == paramGuid);

                        if (def != null)
                        {
                            if (!bindingMap.Contains(def))
                            {
                                InstanceBinding newBinding = app.Create.NewInstanceBinding(targetCategories);
                                bindingMap.Insert(def, newBinding, GroupTypeId.Data);
                                boundCount++;
                            }
                            else
                            {
                                alreadyBoundCount++;
                            }
                        }
                    }
                    tx.Commit();
                    report.AppendLine($"\n- For '{processName}': {boundCount} new parameters were bound. ({alreadyBoundCount} were already bound).");
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"\n- ERROR for '{processName}': {ex.Message}");
            }
        }

        private void AddParametersToFamily(ExternalCommandData commandData, List<string> paramNames, string processName, StringBuilder report)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Application app = doc.Application;

            if (!doc.IsFamilyDocument || doc.OwnerFamily == null)
            {
                report.AppendLine($"\n- ERROR for '{processName}': This action can only be run in a family document.");
                return;
            }

            string sharedParamFilepath = app.SharedParametersFilename;
            if (string.IsNullOrEmpty(sharedParamFilepath) || !File.Exists(sharedParamFilepath))
            {
                report.AppendLine($"\n- ERROR for '{processName}': A valid Shared Parameter file is not configured.");
                return;
            }

            try
            {
                app.SharedParametersFilename = sharedParamFilepath;
                DefinitionFile sharedParamFile = app.OpenSharedParameterFile();
                if (sharedParamFile == null) throw new Exception("Failed to open shared parameter file.");

                using (Transaction t = new Transaction(doc, $"Add {processName} Family Parameters"))
                {
                    t.Start();
                    int addedCount = 0;
                    int existingCount = 0;
                    foreach (string paramName in paramNames)
                    {
                        Guid paramGuid = SharedParameters.GetParameterGuidByName(paramName);
                        ExternalDefinition def = sharedParamFile.Groups.SelectMany(g => g.Definitions).OfType<ExternalDefinition>().FirstOrDefault(d => d.GUID == paramGuid);

                        if (def != null)
                        {
                            if (!doc.FamilyManager.Parameters.Cast<FamilyParameter>().Any(fp => fp.GUID == def.GUID))
                            {
                                bool isInstance = SharedParameterData.MySharedParameters.First(p => p.Guid == paramGuid).IsInstance;
                                doc.FamilyManager.AddParameter(def, GroupTypeId.IdentityData, isInstance);
                                addedCount++;
                            }
                            else
                            {
                                existingCount++;
                            }
                        }
                    }
                    t.Commit();
                    report.AppendLine($"\n- For '{processName}': {addedCount} new parameter(s) were added to the family. ({existingCount} already existed).");
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"\n- ERROR adding parameters for '{processName}': {ex.Message}");
            }
        }
    }
}
