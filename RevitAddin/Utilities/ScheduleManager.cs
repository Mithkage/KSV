//
// File: ScheduleManager.cs
//
// Namespace: RTS.Utilities
//
// Class: ScheduleManager
//
// Function: This file acts as a centralized repository for all standard schedule
//           definitions used across the RTS add-in suite. It provides a single
//           source of truth for schedule names, categories, and their required
//           shared parameter fields, making other commands easier to maintain.
//
// --- CHANGE LOG ---
// 2024-08-13:
// - Initial creation of the ScheduleManager.
// - Populated with standard schedule definitions from RTS_Schedules.cs.
//
// Author: ReTick Solutions
//
#region Namespaces
using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
#endregion

namespace RTS.Utilities
{
    /// <summary>
    /// Defines a standard project schedule, including its name, category, and required fields.
    /// </summary>
    public class ScheduleDefinition
    {
        public string Name { get; set; }
        public BuiltInCategory Category { get; set; }
        public List<Guid> RequiredSharedParameterGuids { get; set; } = new List<Guid>();
        public List<BuiltInParameter> RequiredBuiltInParameters { get; set; } = new List<BuiltInParameter>();
    }

    /// <summary>
    /// Provides a centralized list of all standard schedule definitions.
    /// </summary>
    public static class ScheduleManager
    {
        public static readonly List<ScheduleDefinition> StandardSchedules;

        static ScheduleManager()
        {
            StandardSchedules = new List<ScheduleDefinition>
            {
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Sheet List",
                    Category = BuiltInCategory.OST_Sheets
                    // No specific shared parameters needed, but RTS_ID will be added if applicable.
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Detail Items",
                    Category = BuiltInCategory.OST_DetailComponents,
                    RequiredSharedParameterGuids = new List<Guid>
                    {
                        SharedParameters.General.RTS_ID
                    }
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Spaces",
                    Category = BuiltInCategory.OST_MEPSpaces
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Sample Register",
                    Category = BuiltInCategory.INVALID // This creates a multi-category schedule
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Electrical Equipment",
                    Category = BuiltInCategory.OST_ElectricalEquipment,
                    RequiredSharedParameterGuids = new List<Guid>
                    {
                        SharedParameters.General.RTS_ID
                    },
                    RequiredBuiltInParameters = new List<BuiltInParameter>
                    {
                        BuiltInParameter.RBS_ELEC_PANEL_NAME,
                        BuiltInParameter.FAMILY_LEVEL_PARAM,
                        BuiltInParameter.ELEM_PARTITION_PARAM, // Workset
                        BuiltInParameter.FAMILY_HEIGHT_PARAM,
                        BuiltInParameter.FAMILY_WIDTH_PARAM,
                        BuiltInParameter.INSTANCE_LENGTH_PARAM
                    }
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Lighting Fixtures",
                    Category = BuiltInCategory.OST_LightingFixtures,
                    RequiredSharedParameterGuids = new List<Guid>
                    {
                        SharedParameters.General.RTS_ID
                    }
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Electrical Fixtures",
                    Category = BuiltInCategory.OST_ElectricalFixtures,
                    RequiredSharedParameterGuids = new List<Guid>
                    {
                        SharedParameters.General.RTS_ID
                    }
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Cable Trays",
                    Category = BuiltInCategory.OST_CableTray,
                    RequiredSharedParameterGuids = GetAllCableAndTrayParams()
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Cable Tray Fittings",
                    Category = BuiltInCategory.OST_CableTrayFitting,
                    RequiredSharedParameterGuids = GetAllCableAndTrayParams()
                },
                new ScheduleDefinition
                {
                    Name = "RTS_Sc_Conduits",
                    Category = BuiltInCategory.OST_Conduit,
                    RequiredSharedParameterGuids = GetAllCableAndTrayParams()
                }
            };
        }

        /// <summary>
        /// Helper method to get all parameters for tray and conduit schedules.
        /// </summary>
        private static List<Guid> GetAllCableAndTrayParams()
        {
            var guids = new List<Guid>
            {
                SharedParameters.General.RTS_ID,
                SharedParameters.Cable.RT_Cables_Weight,
                SharedParameters.Cable.RT_Tray_Min_Size,
                SharedParameters.Cable.RT_Tray_Occupancy
            };
            guids.AddRange(SharedParameters.Cable.AllCableGuids);
            return guids;
        }
    }
}
