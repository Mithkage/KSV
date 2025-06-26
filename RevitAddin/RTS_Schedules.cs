//
// Copyright (c) 2025. All rights reserved.
//
// Author: ReTick Solutions
//
// This script is a Revit External Command designed to automate the creation
// of a standard set of project schedules. When executed, the script will:
// 1. Define a list of ten standard schedules with specific names and categories.
// 2. Check if any of the target schedules are currently open in the UI and close them.
// 3. For each schedule in the list, it checks if a schedule with the same name
//    already exists in the active Revit project.
// 4. If an existing schedule is found, it is deleted.
// 5. A new schedule is then created with the specified name and type.
// 6. Each newly created schedule is initialized with default fields and column widths.
//    - "Family" column width is set to 60mm.
//    - "Type" column width is set to 80mm.
//    - "RTS_ID" column width is set to 50mm.
// 7. For Cable Tray, Cable Tray Fitting, and Conduit schedules, a specific list of 
//    shared parameters is also added. These schedules are filtered to only show items
//    where 'RTS_ID' has a value, and are sorted by 'RTS_ID' then by 'Type'.
// 8. A report is displayed to the user summarizing the creation/update status of each schedule.
//
// This ensures a consistent and standardized set of schedules is present in the project,
// ready for further customization.
//

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace RTS_Schedules
{
    /// <summary>
    /// This is the main class for the external command. It implements the IExternalCommand interface.
    /// The command finds and deletes a predefined list of schedules and recreates them with
    //  default fields and column widths to ensure project standards are met.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_SchedulesClass : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // Get the application and document from the command data.
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Define the list of schedules to be created.
            var schedulesToCreate = new List<(string Name, BuiltInCategory Category)>
            {
                ("RTS_Sheet List", BuiltInCategory.OST_Sheets),
                ("RTS_Detail Items", BuiltInCategory.OST_DetailComponents),
                ("RTS_Spaces", BuiltInCategory.OST_MEPSpaces),
                ("RTS_Sample Register", BuiltInCategory.INVALID),
                ("RTS_Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
                ("RTS_Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
                ("RTS_Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
                ("RTS_Cable Trays", BuiltInCategory.OST_CableTray),
                ("RTS_Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting),
                ("RTS_Conduits", BuiltInCategory.OST_Conduit)
            };

            // --- Phase 1: Close any open schedules that will be modified (UI Operation - No Transaction) ---
            List<string> scheduleNamesToProcess = schedulesToCreate.Select(s => s.Name).ToList();
            IList<UIView> openUIViews = uidoc.GetOpenUIViews();
            List<UIView> viewsToClose = new List<UIView>();

            foreach (UIView uiView in openUIViews)
            {
                if (uiView == null) continue;
                Element viewElement = doc.GetElement(uiView.ViewId);

                // Check if the open view is a schedule and if its name is in our target list.
                if (viewElement is ViewSchedule scheduleView && scheduleNamesToProcess.Contains(scheduleView.Name))
                {
                    viewsToClose.Add(uiView);
                }
            }

            // Close the identified views.
            foreach (UIView viewToClose in viewsToClose)
            {
                viewToClose.Close();
            }

            // --- Phase 2: Modify the database (DB Operation - With Transaction) ---
            (string Name, BuiltInCategory Category) currentScheduleInfo = default;
            List<string> reportMessages = new List<string>();

            using (Transaction tx = new Transaction(doc, "Generate RTS Schedules"))
            {
                tx.Start();

                try
                {
                    foreach (var scheduleInfo in schedulesToCreate)
                    {
                        currentScheduleInfo = scheduleInfo;

                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        ViewSchedule existingSchedule = collector.OfClass(typeof(ViewSchedule))
                            .Cast<ViewSchedule>()
                            .FirstOrDefault(s => s.Name.Equals(scheduleInfo.Name, StringComparison.OrdinalIgnoreCase));

                        if (existingSchedule != null)
                        {
                            doc.Delete(existingSchedule.Id);
                            reportMessages.Add($"{scheduleInfo.Name}: Updated");
                        }
                        else
                        {
                            reportMessages.Add($"{scheduleInfo.Name}: Created");
                        }

                        ViewSchedule newSchedule = null;
                        if (scheduleInfo.Category == BuiltInCategory.OST_Sheets)
                        {
                            newSchedule = ViewSchedule.CreateSheetList(doc);
                        }
                        else if (scheduleInfo.Category == BuiltInCategory.INVALID)
                        {
                            newSchedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.INVALID));
                        }
                        else
                        {
                            newSchedule = ViewSchedule.CreateSchedule(doc, new ElementId(scheduleInfo.Category));
                        }

                        if (newSchedule != null)
                        {
                            newSchedule.Name = scheduleInfo.Name;
                            AddFieldsAndSetWidths(newSchedule, scheduleInfo.Category, doc);
                        }
                    }

                    tx.Commit();

                    string finalReport = string.Join("\n", reportMessages);
                    TaskDialog.Show("RTS Schedule Generation Complete", "The following schedules were processed successfully:\n\n" + finalReport);

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    string errorDetails = $"An unexpected error occurred while processing schedule: '{currentScheduleInfo.Name}'\n\n" +
                                          $"Error: {ex.Message}\n\n" +
                                          $"StackTrace:\n{ex.StackTrace}";
                    TaskDialog.Show("RTS Schedule Generation Error", errorDetails);
                    message = ex.Message;
                    tx.RollBack();
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Adds fields to a schedule and then sets column widths, filters, and sorting.
        /// </summary>
        private void AddFieldsAndSetWidths(ViewSchedule schedule, BuiltInCategory category, Document doc)
        {
            if (schedule == null) return;

            ScheduleDefinition definition = schedule.Definition;

            // --- STEP 1: Add all fields to the schedule definition first ---

            if (category == BuiltInCategory.OST_Sheets)
            {
                FindAndAddField(definition, BuiltInParameter.SHEET_NUMBER);
                FindAndAddField(definition, BuiltInParameter.SHEET_NAME);
            }
            else if (category == BuiltInCategory.OST_MEPSpaces)
            {
                FindAndAddField(definition, BuiltInParameter.ROOM_NUMBER);
                FindAndAddField(definition, BuiltInParameter.ROOM_NAME);
            }
            else if (category == BuiltInCategory.OST_CableTray || category == BuiltInCategory.OST_CableTrayFitting || category == BuiltInCategory.OST_Conduit)
            {
                FindAndAddField(definition, BuiltInParameter.ELEM_FAMILY_PARAM);
                FindAndAddField(definition, BuiltInParameter.ELEM_TYPE_PARAM);

                var sharedParameterGuids = new List<string>
                {
                    "3175a27e-d386-4567-bf10-2da1a9cbb73b", // RTS_ID
                    "51d670fa-0338-42e7-ac9e-f2c44a56ffcc", // RT_Cables Weight
                    "5ed6b64c-af5c-4425-ab69-85a7fa5fdffe", // RT_Tray Min Size
                    "a6f087c7-cecc-4335-831b-249cb9398abc", // RT_Tray Occupancy
                    "cf0d478e-1e98-4e83-ab80-6ee867f61798", // RTS_Cable_01
                    "2551d308-44ed-405c-8aad-fb78624d086e", // RTS_Cable_02
                    "c1dfc402-2101-4e53-8f52-f6af64584a9f", // RTS_Cable_03
                    "f297daa6-a9e0-4dd5-bda3-c628db7c28bd", // RTS_Cable_04
                    "b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab", // RTS_Cable_05
                    "7c08095a-a3b2-4b78-ba15-dde09a7bc3a9", // RTS_Cable_06
                    "9bc78bce-0d39-4538-b507-7b98e8a13404", // RTS_Cable_07
                    "e9d50153-a0e9-4685-bc92-d89f244f7e8e", // RTS_Cable_08
                    "5713d65a-91df-4d2e-97bf-1c3a10ea5225", // RTS_Cable_09
                    "64af3105-b2fd-44bc-9ad3-17264049ff62", // RTS_Cable_10
                    "f3626002-0e62-4b75-93cc-35d0b11dfd67", // RTS_Cable_11
                    "63dc0a2e-0770-4002-a859-a9d40a2ce023", // RTS_Cable_12
                    "eb7c4b98-d676-4e2b-a408-e3578b2c0ef2", // RTS_Cable_13
                    "0e0572e5-c568-42b7-8730-a97433bd9b54", // RTS_Cable_14
                    "bf9cd3e8-e38f-4250-9daa-c0fc67eca10f", // RTS_Cable_15
                    "f6d2af67-027e-4b9c-9def-336ebaa87336", // RTS_Cable_16
                    "f6a4459d-46a1-44c0-8545-ee44e4778854", // RTS_Cable_17
                    "0d66d2fa-f261-4daa-8041-9eadeefac49a", // RTS_Cable_18
                    "af483914-c8d2-4ce6-be6e-ab81661e5bf1", // RTS_Cable_19
                    "c8d2d2fc-c248-483f-8d52-e630eb730cd7", // RTS_Cable_20
                    "aa41bc4a-e3e7-45b0-81fa-74d3e71ca506", // RTS_Cable_21
                    "6cffdb25-8270-4b34-8bb4-cf5d0a224dc2", // RTS_Cable_22
                    "7fdaad3a-454e-47f3-8189-7eda9cb9f6a2", // RTS_Cable_23
                    "7f745b2b-a537-42d9-8838-7a5521cc7d0c", // RTS_Cable_24
                    "9a76c2dc-1022-4a54-ab66-5ca625b50365", // RTS_Cable_25
                    "658e39c4-bbac-4e2e-b649-2f2f5dd05b5e", // RTS_Cable_26
                    "8ad24640-036b-44d2-af9c-b891f6e64271", // RTS_Cable_27
                    "c046c4d7-e1fd-4cf7-a99f-14ae96b722be", // RTS_Cable_28
                    "cdf00587-7e11-4af4-8e54-48586481cf22", // RTS_Cable_29
                    "a92bb0f9-2781-4971-a3b1-9c47d62b947b"  // RTS_Cable_30
                };

                foreach (var guidString in sharedParameterGuids)
                {
                    FindAndAddSharedField(definition, doc, new Guid(guidString));
                }
            }
            else
            {
                FindAndAddField(definition, BuiltInParameter.ELEM_FAMILY_PARAM);
                FindAndAddField(definition, BuiltInParameter.ELEM_TYPE_PARAM);
            }

            // --- STEP 2: After all fields are added, get table data and set widths ---
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            ElementId rtsIdParamId = SharedParameterElement.Lookup(doc, new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b"))?.Id;
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
                    typeField = field; // Capture the 'Type' field for sorting
                }
                else if (rtsIdParamId != null && paramId == rtsIdParamId)
                {
                    sectionData.SetColumnWidth(i, MmToFeet(50));
                    rtsIdField = field; // Capture the 'RTS_ID' field for filtering/sorting
                }
            }

            // --- STEP 3: Add Sorting and Filtering for relevant schedules ---
            if (category == BuiltInCategory.OST_CableTray || category == BuiltInCategory.OST_CableTrayFitting || category == BuiltInCategory.OST_Conduit)
            {
                // Apply Filter: RTS_ID has a value
                if (rtsIdField != null)
                {
                    ScheduleFilter filter = new ScheduleFilter(rtsIdField.FieldId, ScheduleFilterType.HasValue);
                    definition.AddFilter(filter);
                }

                // Apply Sorting: First by RTS_ID, then by Type
                if (rtsIdField != null)
                {
                    ScheduleSortGroupField sortById = new ScheduleSortGroupField(rtsIdField.FieldId, ScheduleSortOrder.Ascending);
                    definition.AddSortGroupField(sortById);
                }
                if (typeField != null)
                {
                    ScheduleSortGroupField sortByType = new ScheduleSortGroupField(typeField.FieldId, ScheduleSortOrder.Ascending);
                    definition.AddSortGroupField(sortByType);
                }
            }
        }

        /// <summary>
        /// Finds and adds a schedulable field by its BuiltInParameter.
        /// </summary>
        private void FindAndAddField(ScheduleDefinition definition, BuiltInParameter param)
        {
            IList<SchedulableField> schedulableFields = definition.GetSchedulableFields();
            SchedulableField fieldToAdd = schedulableFields.FirstOrDefault(sf => sf.ParameterId == new ElementId(param));
            if (fieldToAdd != null)
            {
                definition.AddField(fieldToAdd);
            }
        }

        /// <summary>
        /// Finds and adds a Shared Parameter by its GUID.
        /// </summary>
        private void FindAndAddSharedField(ScheduleDefinition definition, Document doc, Guid paramGuid)
        {
            SharedParameterElement sharedParamElem = SharedParameterElement.Lookup(doc, paramGuid);
            if (sharedParamElem != null)
            {
                SchedulableField fieldToAdd = definition.GetSchedulableFields()
                    .FirstOrDefault(sf => sf.ParameterId == sharedParamElem.Id);
                if (fieldToAdd != null)
                {
                    definition.AddField(fieldToAdd);
                }
            }
        }

        /// <summary>
        /// Converts a value from millimeters to decimal feet for the Revit API.
        /// </summary>
        private double MmToFeet(double mm)
        {
            return mm / 304.8;
        }
    }
}
