//
// Copyright (c) 2025. All rights reserved.
//
// Author: ReTick Solutions
//
// This script is a Revit External Command designed to automate the creation
// of a standard set of project schedules. When executed, the script will:
// 1. Read a list of standard schedules from the centralized ScheduleManager.
// 2. Present a WPF UI to the user to select which schedules to generate or remove.
// 3. For selected schedules, it checks if any are currently open and closes them.
// 4. If an existing schedule is found, it is deleted before recreation.
// 5. A new schedule is then created with the specified name and type.
// 6. Each newly created schedule is initialized with fields and settings defined in the ScheduleManager.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - [APPLIED FIX]: Corrected Revit API usage for getting a ScheduleField to resolve CS0103 compiler error.
// - [APPLIED FIX]: Refactored to use the centralized ScheduleManager for all schedule definitions.
// - [APPLIED FIX]: Resolved namespace collision by using an alias for the custom ScheduleDefinition class.
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.UI;
using RTS.Utilities; // Required for ScheduleManager and SharedParameters
using System;
using System.Collections.Generic;
using System.Linq;
using RtsScheduleDef = RTS.Utilities.ScheduleDefinition; // Alias to resolve namespace conflict
#endregion

namespace RTS.Commands.ProjectSetup
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_SchedulesClass : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --- 1. Get schedule definitions from the centralized ScheduleManager ---
            var allSchedules = ScheduleManager.StandardSchedules
                .Select(def => new ScheduleItem { Name = def.Name, Category = def.Category, IsSelected = false })
                .ToList();

            // --- 2. Show UI for user selection ---
            ScheduleSelectionWindow ui = new ScheduleSelectionWindow(allSchedules);
            bool? dialogResult = ui.ShowDialog();

            if (dialogResult != true)
            {
                return Result.Cancelled;
            }

            List<string> reportMessages = new List<string>();
            List<ScheduleItem> schedulesToProcess = ui.Schedules.Where(s => s.IsSelected).ToList();

            if (schedulesToProcess.Count == 0 && !ui.RemoveAllSelected)
            {
                TaskDialog.Show("RTS Schedule Generation", "No schedules were selected. Operation cancelled.");
                return Result.Cancelled;
            }

            // --- 3. Close any open schedules that will be modified/deleted ---
            List<string> scheduleNamesToProcess = schedulesToProcess.Select(s => s.Name).ToList();
            if (ui.RemoveAllSelected)
            {
                scheduleNamesToProcess = allSchedules.Select(s => s.Name).ToList();
            }

            IList<UIView> openUIViews = uidoc.GetOpenUIViews();
            foreach (UIView uiView in openUIViews)
            {
                if (doc.GetElement(uiView.ViewId) is ViewSchedule scheduleView && scheduleNamesToProcess.Contains(scheduleView.Name))
                {
                    uiView.Close();
                }
            }

            // --- 4. Modify the database ---
            using (Transaction tx = new Transaction(doc, "Manage RTS Schedules"))
            {
                tx.Start();
                try
                {
                    if (ui.RemoveAllSelected)
                    {
                        reportMessages.Add("--- Removing All Predefined Schedules ---");
                        foreach (var scheduleItem in allSchedules)
                        {
                            DeleteExistingSchedule(doc, scheduleItem.Name, reportMessages);
                        }
                    }
                    else
                    {
                        foreach (var scheduleItem in schedulesToProcess)
                        {
                            DeleteExistingSchedule(doc, scheduleItem.Name, reportMessages);
                            CreateAndConfigureSchedule(doc, scheduleItem);
                        }
                    }
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    tx.RollBack();
                    TaskDialog.Show("RTS Schedule Management Error", $"An unexpected error occurred: {ex.Message}");
                    return Result.Failed;
                }
            }

            string finalReport = string.Join("\n", reportMessages);
            TaskDialog.Show("RTS Schedule Management Complete", "The following schedules were processed successfully:\n\n" + finalReport);
            return Result.Succeeded;
        }

        private void DeleteExistingSchedule(Document doc, string scheduleName, List<string> reportMessages)
        {
            ViewSchedule existingSchedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(s => s.Name.Equals(scheduleName, StringComparison.OrdinalIgnoreCase));

            if (existingSchedule != null)
            {
                doc.Delete(existingSchedule.Id);
                reportMessages.Add($"{scheduleName}: Updated (previous version removed)");
            }
            else
            {
                reportMessages.Add($"{scheduleName}: Created");
            }
        }

        private void CreateAndConfigureSchedule(Document doc, ScheduleItem scheduleItem)
        {
            ViewSchedule newSchedule = null;
            if (scheduleItem.Category == BuiltInCategory.OST_Sheets)
            {
                newSchedule = ViewSchedule.CreateSheetList(doc);
            }
            else if (scheduleItem.Category == BuiltInCategory.INVALID)
            {
                newSchedule = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId);
            }
            else
            {
                newSchedule = ViewSchedule.CreateSchedule(doc, new ElementId(scheduleItem.Category));
            }

            if (newSchedule != null)
            {
                newSchedule.Name = scheduleItem.Name;
                AddFieldsAndSetWidths(newSchedule, scheduleItem.Category, doc);
            }
        }

        private void AddFieldsAndSetWidths(ViewSchedule schedule, BuiltInCategory category, Document doc)
        {
            if (schedule == null) return;

            Autodesk.Revit.DB.ScheduleDefinition definition = schedule.Definition;

            // Add common fields
            if (category != BuiltInCategory.OST_Sheets && category != BuiltInCategory.INVALID)
            {
                FindAndAddField(definition, BuiltInParameter.ELEM_FAMILY_PARAM);
                FindAndAddField(definition, BuiltInParameter.ELEM_TYPE_PARAM);
            }
            else if (category == BuiltInCategory.OST_Sheets)
            {
                FindAndAddField(definition, BuiltInParameter.SHEET_NUMBER);
                FindAndAddField(definition, BuiltInParameter.SHEET_NAME);
            }

            // Get the corresponding definition from the manager to add required shared parameters
            RtsScheduleDef rtsDef = ScheduleManager.StandardSchedules.FirstOrDefault(s => s.Name == schedule.Name);
            if (rtsDef != null)
            {
                foreach (var guid in rtsDef.RequiredSharedParameterGuids)
                {
                    FindAndAddSharedField(definition, doc, guid);
                }
            }

            // Set widths and sorting/filtering
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            ElementId rtsIdParamId = SharedParameterElement.Lookup(doc, SharedParameters.General.RTS_ID)?.Id;
            ElementId typeParamId = new ElementId(BuiltInParameter.ELEM_TYPE_PARAM);

            ScheduleField rtsIdField = null;
            ScheduleField typeField = null;

            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                ScheduleField field = definition.GetField(i);
                ElementId paramId = field.ParameterId;

                if (paramId == new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM))
                {
                    sectionData.SetColumnWidth(i, MmToFeet(60));
                }
                else if (paramId == typeParamId)
                {
                    sectionData.SetColumnWidth(i, MmToFeet(80));
                    typeField = field;
                }
                else if (rtsIdParamId != null && paramId == rtsIdParamId)
                {
                    sectionData.SetColumnWidth(i, MmToFeet(50));
                    rtsIdField = field;
                }
            }

            if (rtsIdField != null)
            {
                definition.AddFilter(new ScheduleFilter(rtsIdField.FieldId, ScheduleFilterType.HasValue));
                definition.ClearSortGroupFields();
                definition.AddSortGroupField(new ScheduleSortGroupField(rtsIdField.FieldId, ScheduleSortOrder.Ascending));
                if (typeField != null)
                {
                    definition.AddSortGroupField(new ScheduleSortGroupField(typeField.FieldId, ScheduleSortOrder.Ascending));
                }
            }
        }

        private void FindAndAddField(Autodesk.Revit.DB.ScheduleDefinition definition, BuiltInParameter param)
        {
            SchedulableField fieldToAdd = definition.GetSchedulableFields().FirstOrDefault(sf => sf.ParameterId == new ElementId(param));
            if (fieldToAdd != null)
            {
                definition.AddField(fieldToAdd);
            }
        }

        private void FindAndAddSharedField(Autodesk.Revit.DB.ScheduleDefinition definition, Document doc, Guid paramGuid)
        {
            SharedParameterElement sharedParamElem = SharedParameterElement.Lookup(doc, paramGuid);
            if (sharedParamElem != null)
            {
                SchedulableField fieldToAdd = definition.GetSchedulableFields().FirstOrDefault(sf => sf.ParameterId == sharedParamElem.Id);
                if (fieldToAdd != null)
                {
                    definition.AddField(fieldToAdd);
                }
            }
        }

        private double MmToFeet(double mm)
        {
            return mm / 304.8;
        }
    }
}
