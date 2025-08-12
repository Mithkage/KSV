// SharedParameterFile.cs

// --- FILE DESCRIPTION ---
// This file contains the hardcoded definitions for all shared parameter groups and parameters
// used within the RTS add-in ecosystem. It centralizes the data, making it easier to
// update and manage parameter definitions without altering the functional logic
// in the main SharedParameters.cs file.

using System;
using System.Collections.Generic;

namespace RTS.Utilities
{
    /// <summary>
    /// Contains the static data definitions for shared parameter groups and parameters.
    /// </summary>
    public static class SharedParameterData
    {
        // --- Structured Data Classes for Easy Management ---

        /// <summary>
        /// Represents a shared parameter group from the .txt file.
        /// </summary>
        public class SharedParameterGroup
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        /// <summary>
        /// Represents a single shared parameter definition from the .txt file.
        /// </summary>
        public class SharedParameterDefinition
        {
            public Guid Guid { get; set; }
            public string Name { get; set; }
            public string DataType { get; set; }
            public int GroupId { get; set; }
            public bool Visible { get; set; } = true;
            public string Description { get; set; }
            public bool UserModifiable { get; set; } = true;
            public bool HideWhenNoValue { get; set; } = false;
            /// <summary>
            /// If true, the parameter is bound as an Instance parameter. If false, it is bound as a Type parameter.
            /// </summary>
            public bool IsInstance { get; set; } = false;
        }

        // --- Data Population (from RTS_Shared Parameters_9.txt) ---
        // This is where you would add or remove entries to manage your parameters.

        public static readonly List<SharedParameterGroup> MySharedParameterGroups = new List<SharedParameterGroup>
        {
            new SharedParameterGroup { Id = 1, Name = "_General" },
            new SharedParameterGroup { Id = 2, Name = "RSGx" },
            new SharedParameterGroup { Id = 3, Name = "PowerCAD" },
            new SharedParameterGroup { Id = 4, Name = "Lighting" },
            new SharedParameterGroup { Id = 5, Name = "Manufacture" },
            new SharedParameterGroup { Id = 6, Name = "_Initiate" },
            new SharedParameterGroup { Id = 7, Name = "Electrical" },
            new SharedParameterGroup { Id = 8, Name = "Cable Tray" },
            new SharedParameterGroup { Id = 9, Name = "Fire" },
            new SharedParameterGroup { Id = 10, Name = "Internal" }
        };

        // NOTE: Parameters are now sorted alphabetically by name within each group.
        // The 'IsInstance' property determines if a parameter is for an instance or a type.
        public static readonly List<SharedParameterDefinition> MySharedParameters = new List<SharedParameterDefinition>
        {
            // Group 1: _General
            new SharedParameterDefinition { Guid = new Guid("1c140da3-85a3-4aa7-a8d6-5880d7528613"), Name = "RTS_Appendix_Name", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The name of the appendix document associated with the element." },
            new SharedParameterDefinition { Guid = new Guid("12f08408-64ca-49cf-aadc-047c95e44fc4"), Name = "RTS_Approved", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: Records the approval status of an element." },
            new SharedParameterDefinition { Guid = new Guid("8f1a6083-a6c4-4e2f-abaa-54ebe4c9a9f2"), Name = "RTS_Approved_Date", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The date when the element was approved." },
            new SharedParameterDefinition { Guid = new Guid("1c5f2502-237d-4761-9763-9cabb623d92d"), Name = "RTS_Bell_Diameter", DataType = "LENGTH", GroupId = 1, Description = "AI: The diameter of the bell end of a fitting." },
            new SharedParameterDefinition { Guid = new Guid("21837930-2b6d-43e0-a293-e5f7611efc90"), Name = "RTS_Building", DataType = "TEXT", GroupId = 1, Description = "AI: The building name or identifier where the element is located." },
            new SharedParameterDefinition { Guid = new Guid("c4016d84-2a0a-455b-83fa-beb9cdc32870"), Name = "RTS_Building_Class", DataType = "TEXT", GroupId = 1, Description = "AI: The classification of the building (e.g., Class 2, Class 9a)." },
            new SharedParameterDefinition { Guid = new Guid("c0dd173f-5455-42f2-82ad-ab266fc37355"), Name = "RTS_Checked", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: Records the name or initials of the person who checked the element." },
            new SharedParameterDefinition { Guid = new Guid("e2370fde-3a64-4acd-8172-ebf3ad8af015"), Name = "RTS_Checked_Date", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The date when the element was checked." },
            new SharedParameterDefinition { Guid = new Guid("68dba510-98de-4ab8-9629-12a33f65aa98"), Name = "RTS_Code", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A code or identifier for the element." },
            new SharedParameterDefinition { Guid = new Guid("f8a844ce-cb1a-4d95-bb11-d48d15a84a8e"), Name = "RTS_Comment", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: General comments or notes about the element." },
            new SharedParameterDefinition { Guid = new Guid("a377a731-1886-4b93-b477-a6208f672987"), Name = "RTS_CopyReference_ID", DataType = "TEXT", GroupId = 1, IsInstance = true, Visible = false, UserModifiable = false, Description = "AI: Internal ID for tracking copied elements." },
            new SharedParameterDefinition { Guid = new Guid("04fd355f-7943-4020-956c-18411bddc6ba"), Name = "RTS_CopyRelative_Position", DataType = "TEXT", GroupId = 1, IsInstance = true, Visible = false, UserModifiable = false, Description = "AI: Internal data for relative positioning of copied elements." },
            new SharedParameterDefinition { Guid = new Guid("4d6ce1ad-eb55-47e2-acb6-69490634990e"), Name = "RTS_CopyReference_Type", DataType = "TEXT", GroupId = 1, IsInstance = true, Visible = false, UserModifiable = false, Description = "AI: Internal type for tracking copied elements." },
            new SharedParameterDefinition { Guid = new Guid("ad81cebc-41cb-4749-b213-b918fa6f938e"), Name = "RTS_Created_By", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The user who created the element." },
            new SharedParameterDefinition { Guid = new Guid("0e44a1d2-ed20-4892-a9da-645a5a9fbb2b"), Name = "RTS_Depth", DataType = "LENGTH", GroupId = 1, Description = "AI: The depth of the element." },
            new SharedParameterDefinition { Guid = new Guid("8004b3a7-87ae-4e7b-9f2a-de6d98c51f5e"), Name = "RTS_Design_Date", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The date of the design." },
            new SharedParameterDefinition { Guid = new Guid("83407bc9-c9e8-4f52-9713-398e73040e74"), Name = "RTS_Designation", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Used for designating a purpose to a tray or conduit" },
            new SharedParameterDefinition { Guid = new Guid("233888ac-7038-45f9-8e6c-ef1557b79f65"), Name = "RTS_Diameter", DataType = "LENGTH", GroupId = 1, Description = "AI: The nominal diameter of the element." },
            new SharedParameterDefinition { Guid = new Guid("dfc05318-6bd5-4b49-b883-fadc60f49929"), Name = "RTS_Discipline", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The engineering discipline associated with the element (e.g., Electrical, Fire)." },
            new SharedParameterDefinition { Guid = new Guid("f35bada0-303b-46d6-a9f1-b024383ff344"), Name = "RTS_Drawn", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The name or initials of the person who drafted the element." },
            new SharedParameterDefinition { Guid = new Guid("f7739405-4b26-43c5-b15c-c16740d97d1c"), Name = "RTS_Elevation_From_Level", DataType = "LENGTH", GroupId = 1, IsInstance = true, Description = "AI: The elevation of the element relative to its associated level." },
            new SharedParameterDefinition { Guid = new Guid("6dff7805-d967-41a8-89a7-dcd55db8f6ce"), Name = "RTS_Error", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Used for element error reporting" },
            new SharedParameterDefinition { Guid = new Guid("1b36de09-c6f2-4de8-bd30-a91f98c524a6"), Name = "RTS_Filled_Region_Visible", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: Controls the visibility of a filled region within the family." },
            new SharedParameterDefinition { Guid = new Guid("50f448fc-2e23-473d-befd-6eaa11051594"), Name = "RTS_Filter", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A filter value used for view visibility and scheduling." },
            new SharedParameterDefinition { Guid = new Guid("3f503834-26e8-46a8-8acf-e96e47ea58c4"), Name = "RTS_FRR", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: Fire Resistance Rating (FRR) for the element." },
            new SharedParameterDefinition { Guid = new Guid("01b9b283-8ddb-469c-b78b-a4fb4c40a34d"), Name = "RTS_Grid_Reference", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The structural grid reference for the element's location." },
            new SharedParameterDefinition { Guid = new Guid("73897540-fd50-4125-938e-e37914f4425e"), Name = "RTS_Height", DataType = "LENGTH", GroupId = 1, Description = "AI: The height of the element." },
            new SharedParameterDefinition { Guid = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b"), Name = "RTS_ID", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Generic ID data" },
            new SharedParameterDefinition { Guid = new Guid("f90d0c05-7e5c-4174-a851-299ab0df759b"), Name = "RTS_ITR", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Type: Inspection Test Record Code used for QA" },
            new SharedParameterDefinition { Guid = new Guid("4c99347c-af5b-4fa4-a6d3-1ae9de995dce"), Name = "RTS_Item_Number", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Used for listing electrical meter item numbers" },
            new SharedParameterDefinition { Guid = new Guid("785f2f84-a088-44ec-870f-b93f4f6e43d6"), Name = "RTS_Left_End_Detail", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A detail reference for the left end of the element." },
            new SharedParameterDefinition { Guid = new Guid("d9b0a125-4f3d-42dc-983d-42a1e478caa2"), Name = "RTS_Level", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The associated building level for the element." },
            new SharedParameterDefinition { Guid = new Guid("6610a2d5-a4de-403d-bb41-e3157a592a5b"), Name = "RTS_Location", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A description of the element's location." },
            new SharedParameterDefinition { Guid = new Guid("ef38e196-9583-46bb-9c98-ce03e5bc52c7"), Name = "RTS_Mass", DataType = "NUMBER", GroupId = 1, Description = "AI: The mass or weight of the element, in kilograms." },
            new SharedParameterDefinition { Guid = new Guid("8ec97a0b-f97b-4dae-ad1f-e0e588e27366"), Name = "RTS_Material", DataType = "MATERIAL", GroupId = 1, Description = "AI: The primary material of the element." },
            new SharedParameterDefinition { Guid = new Guid("22d4e01a-7948-43b6-96db-db31a1fa0e45"), Name = "RTS_MaterialSecondary", DataType = "MATERIAL", GroupId = 1, Description = "AI: The secondary material of the element." },
            new SharedParameterDefinition { Guid = new Guid("6165729a-8e8d-46dc-a108-5581f7020f93"), Name = "RTS_Meter_Brand", DataType = "TEXT", GroupId = 1, Description = "AI: The brand or manufacturer of the meter." },
            new SharedParameterDefinition { Guid = new Guid("fb2d1de3-a2d7-4371-9a03-1a30d145727e"), Name = "RTS_Meter_Description", DataType = "TEXT", GroupId = 1, Description = "AI: A description of the electrical meter." },
            new SharedParameterDefinition { Guid = new Guid("20cf18dc-f7c1-4670-8c33-4a40ad3fcf39"), Name = "RTS_Meter_Number", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The serial number of the electrical meter." },
            new SharedParameterDefinition { Guid = new Guid("e0501aba-d97a-4750-95f4-71c062472585"), Name = "RTS_Meter_Supplier", DataType = "TEXT", GroupId = 1, Description = "AI: The supplier of the electrical meter." },
            new SharedParameterDefinition { Guid = new Guid("562f3b8d-600b-4367-903c-74f46d4ae2c4"), Name = "RTS_Model_Manager", DataType = "TEXT", GroupId = 1, Description = "AI: The name of the model manager responsible for the project." },
            new SharedParameterDefinition { Guid = new Guid("4c191229-96c9-4ff4-a0a6-77a1a7e545cc"), Name = "RTS_Note", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Note shown on drawings" },
            new SharedParameterDefinition { Guid = new Guid("0b664d96-0503-4b76-9556-4b10525ba831"), Name = "RTS_Pipe_Loading_Cap_-_Drain", DataType = "NUMBER", GroupId = 1, Description = "AI: The loading capacity of a drainage pipe." },
            new SharedParameterDefinition { Guid = new Guid("e19e2c68-194e-4b02-8d0c-ccc32d47c8de"), Name = "RTS_QA_ARC", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: A quality assurance check for architectural coordination." },
            new SharedParameterDefinition { Guid = new Guid("561de3d7-d5e8-479d-b6ec-4a1e88a21b89"), Name = "RTS_QA_ASBUILT", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: A quality assurance check for as-built status." },
            new SharedParameterDefinition { Guid = new Guid("f816fea4-81e3-4100-a5cc-d623ad4d7e93"), Name = "RTS_QA_ELEC", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: A quality assurance check for electrical services." },
            new SharedParameterDefinition { Guid = new Guid("6c913130-7731-4d9c-945e-24f192c94397"), Name = "RTS_QA_STRUC", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: A quality assurance check for structural coordination." },
            new SharedParameterDefinition { Guid = new Guid("b1f7c709-b1fa-422a-982a-54a70ac6eac6"), Name = "RTS_QA_VIEW RANGE", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "AI: A quality assurance check for view range settings." },
            new SharedParameterDefinition { Guid = new Guid("e66f50fc-7162-40a3-9a3f-465cd0f15937"), Name = "RTS_Radius", DataType = "LENGTH", GroupId = 1, Description = "AI: The radius of the element." },
            new SharedParameterDefinition { Guid = new Guid("f2cc127c-75ea-498e-a007-b42e3c959e91"), Name = "RTS_Revision", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The drawing or model revision identifier." },
            new SharedParameterDefinition { Guid = new Guid("4ab43d43-ee11-4358-9b47-2efc3e98823f"), Name = "RTS_Sample_Number", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A sample number for testing or quality assurance." },
            new SharedParameterDefinition { Guid = new Guid("436e8731-4fb6-4c53-b44e-163d4bea84ff"), Name = "RTS_Service_Type", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Used for colour filtering. eg. Tray Switchroom A or B or HV" },
            new SharedParameterDefinition { Guid = new Guid("da52c87f-4f35-419a-b06d-1147467af8fa"), Name = "RTS_Sheet_Scale", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The scale of the sheet the view is placed on." },
            new SharedParameterDefinition { Guid = new Guid("6ab4d13f-4e30-4810-aadb-0d3ac79fa8b8"), Name = "RTS_Sheet_Series", DataType = "TEXT", GroupId = 1, Description = "AI: The drawing sheet series (e.g., E, S, A)." },
            new SharedParameterDefinition { Guid = new Guid("12e74f36-ac15-4c9a-acc2-7af1003127d4"), Name = "RTS_Status", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The current status of the element (e.g., For Construction)." },
            new SharedParameterDefinition { Guid = new Guid("1b0151ab-c4e1-4d86-8086-1f87441d3e0a"), Name = "RTS_Stock_Location", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The storage or stock location for the material." },
            new SharedParameterDefinition { Guid = new Guid("8df5f1a7-387a-49f1-8fc0-53a90fb3d0a0"), Name = "RTS_Sub_Discipline", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The sub-discipline (e.g., Power, Lighting)." },
            new SharedParameterDefinition { Guid = new Guid("43ebc6d6-bb07-42eb-9663-afa60ccb0297"), Name = "RTS_Survey Point", DataType = "YESNO", GroupId = 1, IsInstance = true, Description = "Controls Survey Point Visibility" },
            new SharedParameterDefinition { Guid = new Guid("fcd03c0f-9b2c-419b-b167-78141d4c1fa5"), Name = "RTS_Thickness", DataType = "LENGTH", GroupId = 1, Description = "AI: The thickness of the element." },
            new SharedParameterDefinition { Guid = new Guid("2c4d7eeb-5bc7-4c89-b9f4-73d99bb8bf1b"), Name = "RTS_Verified_By", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The name or initials of the person who verified the element." },
            new SharedParameterDefinition { Guid = new Guid("9042e640-2a5f-468a-a5b8-0d4f82f93c58"), Name = "RTS_Verifier", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: The name or initials of the person who verified the element." },
            new SharedParameterDefinition { Guid = new Guid("d6b26332-b677-48f1-ae63-705986439c6d"), Name = "RTS_Version_No.", DataType = "TEXT", GroupId = 1, Description = "AI: The version number of the element or design." },
            new SharedParameterDefinition { Guid = new Guid("19d55866-b017-40ab-92a8-735a8dcbc304"), Name = "RTS_View_Group", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A value used to group views in the Project Browser." },
            new SharedParameterDefinition { Guid = new Guid("7dc44a3f-2d00-49d9-8114-5cbd9330c9ae"), Name = "RTS_Watermark", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "AI: A watermark text to display on views or sheets." },
            new SharedParameterDefinition { Guid = new Guid("0ff67583-3b7f-430c-8764-57b5194a8cc4"), Name = "RTS_Width", DataType = "LENGTH", GroupId = 1, Description = "AI: The width of the element." },
            new SharedParameterDefinition { Guid = new Guid("0b0e5e92-a400-4e93-87d8-d49530e0039c"), Name = "RTS_Zone", DataType = "TEXT", GroupId = 1, IsInstance = true, Description = "Zone location" },

            // Group 2: RSGx (Cable Parameters)
            new SharedParameterDefinition { Guid = new Guid("c7430aff-c4ee-4354-9601-a060364b43d5"), Name = "Cables On Tray", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: A list or count of cables contained on the cable tray." },
            new SharedParameterDefinition { Guid = new Guid("d44ab6c4-aa8c-4abb-b2cb-16038885a7f9"), Name = "Conduit Number", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "Conduit parrallel run Number" },
            new SharedParameterDefinition { Guid = new Guid("4e8047d8-023b-4ae9-ae96-1f871cf51f4e"), Name = "Fire Saftey", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "FLS or ESS if fire rated or essential" },
            new SharedParameterDefinition { Guid = new Guid("51d670fa-0338-42e7-ac9e-f2c44a56ffcc"), Name = "RT_Cables Weight", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: The total weight of the cables on the support element." },
            new SharedParameterDefinition { Guid = new Guid("5ed6b64c-af5c-4425-ab69-85a7fa5fdffe"), Name = "RT_Tray Min Size", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: The minimum required size for the cable tray." },
            new SharedParameterDefinition { Guid = new Guid("a6f087c7-cecc-4335-831b-249cb9398abc"), Name = "RT_Tray Occupancy", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: The calculated occupancy percentage of the cable tray." },
            new SharedParameterDefinition { Guid = new Guid("3ea1a3bb-8416-45ed-b606-3f3a3f87d4be"), Name = "RTS_Branch Number", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "Unique Identifier for Cable Tray Branches" },
            new SharedParameterDefinition { Guid = new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), Name = "RTS_Cable_01", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), Name = "RTS_Cable_02", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"), Name = "RTS_Cable_03", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), Name = "RTS_Cable_04", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"), Name = "RTS_Cable_05", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), Name = "RTS_Cable_06", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"), Name = "RTS_Cable_07", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), Name = "RTS_Cable_08", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"), Name = "RTS_Cable_09", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), Name = "RTS_Cable_10", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"), Name = "RTS_Cable_11", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), Name = "RTS_Cable_12", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"), Name = "RTS_Cable_13", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), Name = "RTS_Cable_14", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"), Name = "RTS_Cable_15", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), Name = "RTS_Cable_16", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"), Name = "RTS_Cable_17", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), Name = "RTS_Cable_18", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"), Name = "RTS_Cable_19", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), Name = "RTS_Cable_20", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"), Name = "RTS_Cable_21", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), Name = "RTS_Cable_22", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"), Name = "RTS_Cable_23", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), Name = "RTS_Cable_24", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"), Name = "RTS_Cable_25", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), Name = "RTS_Cable_26", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"), Name = "RTS_Cable_27", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), Name = "RTS_Cable_28", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"), Name = "RTS_Cable_29", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b"), Name = "RTS_Cable_30", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: Generic data field for cable information." },
            new SharedParameterDefinition { Guid = new Guid("d81b6a0d-1e95-43b6-a752-1b0becaec861"), Name = "RTS_Variance", DataType = "NUMBER", GroupId = 2, IsInstance = true, Description = "Used to calculate the percentage variance between consultant and model values" },
            new SharedParameterDefinition { Guid = new Guid("bc3d8d0a-9ee3-43fa-bca5-0bc414306316"), Name = "Section Tag", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: A tag or identifier for a section view." },
            new SharedParameterDefinition { Guid = new Guid("ce390138-5f27-4ae5-a78d-a5fbd77dfb37"), Name = "Spare Capacity", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: The spare capacity percentage of a cable tray or conduit." },
            new SharedParameterDefinition { Guid = new Guid("b130acc1-d93d-43b6-925a-2a12b8d46dfb"), Name = "Spare Capacity on Circuit", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "AI: The spare capacity available on the electrical circuit." },
            new SharedParameterDefinition { Guid = new Guid("a9613056-877e-42bd-ad74-73707c1ad24e"), Name = "String Supply", DataType = "TEXT", GroupId = 2, IsInstance = true, Description = "Transformer Supply - Typicaly A or B" },
            
            // Group 3: PowerCAD
            new SharedParameterDefinition { Guid = new Guid("f5d583ec-f511-4653-828e-9b45281baa54"), Name = "PC_# of Phases", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The number of phases for the electrical circuit (e.g., 1 or 3)." },
            new SharedParameterDefinition { Guid = new Guid("62ba1846-99f5-4a7f-985a-9255d54c5b93"), Name = "PC_Active Conductor material", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The material of the active conductors (e.g., Copper, Aluminium)." },
            new SharedParameterDefinition { Guid = new Guid("91c8321f-1342-4efa-b648-7ba5e95c0085"), Name = "PC_Cable Size - Active conductors", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The size of the active conductors in the cable (e.g., 6mm²)." },
            new SharedParameterDefinition { Guid = new Guid("50b6ef99-8f5e-42e1-8645-cce97f6b94b6"), Name = "PC_Cable Size - Earthing conductor", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The size of the earthing conductor in the cable (e.g., 2.5mm²)." },
            new SharedParameterDefinition { Guid = new Guid("4dc884d6-07f4-4546-b3aa-98a13d5ae6f1"), Name = "PC_Clearance Time (sec)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The fault clearance time in seconds." },
            new SharedParameterDefinition { Guid = new Guid("214a1cbb-3f5d-48fd-b23b-7915d8e28c6f"), Name = "PC_Conductor Gap (mm)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The gap between conductors in millimeters." },
            new SharedParameterDefinition { Guid = new Guid("f3c6038d-7300-46a5-8787-14560678f531"), Name = "PC_Cores", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "Used for Cable Cores data" },
            new SharedParameterDefinition { Guid = new Guid("5eb8da2d-a094-4f80-8626-9960fb6b4aa3"), Name = "PC_Design Progress", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The progress status of the design (e.g., Preliminary, Detailed)." },
            new SharedParameterDefinition { Guid = new Guid("ee5c5f1a-7e6e-480b-9591-c1391cf0990b"), Name = "PC_Earth Conductor material", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "Earth material" },
            new SharedParameterDefinition { Guid = new Guid("1aff59c0-8add-4dae-851c-64ecd7543f75"), Name = "PC_Electrode Configuration", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The configuration of the earthing electrode." },
            new SharedParameterDefinition { Guid = new Guid("4bff1aed-3203-4c84-91aa-de7eba537b8d"), Name = "PC_Enclosure Depth (mm)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The depth of the enclosure in millimeters." },
            new SharedParameterDefinition { Guid = new Guid("a07e63e4-dafa-4da6-aaf8-f3dca5fea842"), Name = "PC_Enclosure Width (mm)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The width of the enclosure in millimeters." },
            new SharedParameterDefinition { Guid = new Guid("bc3f0ed3-e9f5-456f-83c6-d7f2dc0db169"), Name = "PC_Isolator Type", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The type of isolator used." },
            new SharedParameterDefinition { Guid = new Guid("4bf4ee16-1be4-47a4-bcd8-9fe79429e9f0"), Name = "PC_No. of Conduits", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The number of parallel conduits for the circuit." },
            new SharedParameterDefinition { Guid = new Guid("e1fe521c-673d-4d5a-a69a-407ef8f82518"), Name = "PC_Prospective Fault 3ø Isc (kA)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The prospective 3-phase short circuit current in kA." },
            new SharedParameterDefinition { Guid = new Guid("9a1a8f15-5316-4666-84ee-d84b8c84d8f1"), Name = "PC_Protective Device Breaking Capacity (kA)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The breaking capacity of the protective device in kA." },
            new SharedParameterDefinition { Guid = new Guid("9eb0f777-3bf0-4398-94f9-e7ab26d1fc19"), Name = "PC_Protective Device Description", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: A description of the protective device." },
            new SharedParameterDefinition { Guid = new Guid("285d845a-a1d9-4821-a3d5-09a9ee956c28"), Name = "PC_Protective Device Manufacturer", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The manufacturer of the protective device." },
            new SharedParameterDefinition { Guid = new Guid("c8cd91c1-3d04-46db-b103-700475ba47fd"), Name = "PC_Protective Device OCR/Trip Unit", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The type of overcurrent release or trip unit." },
            new SharedParameterDefinition { Guid = new Guid("50889a75-86ce-4640-9ee2-6495eee39ccf"), Name = "PC_Protective Device Settings", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The settings for the protective device." },
            new SharedParameterDefinition { Guid = new Guid("5dd52911-7bcd-4a06-869d-73fcef59951c"), Name = "PC_SWB From", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The source switchboard for the circuit." },
            new SharedParameterDefinition { Guid = new Guid("60f670e1-7d54-4ffc-b0f5-ed62c08d3b90"), Name = "PC_SWB Load", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The connected load at the switchboard." },
            new SharedParameterDefinition { Guid = new Guid("d325322e-9981-450f-8d41-11bc850b499f"), Name = "PC_SWB Load Scope", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The scope of the load at the switchboard." },
            new SharedParameterDefinition { Guid = new Guid("e142b0ed-d084-447a-991b-d9a3a3f67a8d"), Name = "PC_SWB To", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The destination switchboard for the circuit." },
            new SharedParameterDefinition { Guid = new Guid("9d5ab9c2-09e2-4d42-ae2f-2a5fba6f7131"), Name = "PC_SWB Type", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "AI: The type of switchboard." },
            new SharedParameterDefinition { Guid = new Guid("c3c68f41-8c3b-4f20-b355-bc8263ab557c"), Name = "Protective Device Rating (A)", DataType = "TEXT", GroupId = 3, IsInstance = true, Description = "Frame Size" },

            // Group 4: Lighting
            new SharedParameterDefinition { Guid = new Guid("0e7a4ec4-55b4-414a-b944-811e1aeb1568"), Name = "RTS_Target_Average_Illuminance", DataType = "ELECTRICAL_ILLUMINANCE", GroupId = 4, Description = "AI: The target average illuminance level for a space (in lux)." },
            new SharedParameterDefinition { Guid = new Guid("33fff47b-9400-45b1-8347-391d7c3d67d9"), Name = "RTS_Target_Max_Illuminance", DataType = "ELECTRICAL_ILLUMINANCE", GroupId = 4, Description = "AI: The target maximum illuminance level for a space (in lux)." },
            new SharedParameterDefinition { Guid = new Guid("1ee4c411-73ea-4b15-9ab1-44e88653c6c2"), Name = "RTS_Target_Min_Illuminance", DataType = "ELECTRICAL_ILLUMINANCE", GroupId = 4, Description = "AI: The target minimum illuminance level for a space (in lux)." },

            // Group 5: Manufacture
            new SharedParameterDefinition { Guid = new Guid("370590ac-139c-45b6-ac11-d8f1f96f8b55"), Name = "RTS_Meter_Type", DataType = "TEXT", GroupId = 5, Description = "\"Check Meter (CM), Authority Meter (PM), etc\"" },
            new SharedParameterDefinition { Guid = new Guid("f81cba40-664a-42ed-a049-07a8bb474a7f"), Name = "RTS_Supplier", DataType = "TEXT", GroupId = 5, Description = "AI: The supplier or manufacturer of the component." },
            new SharedParameterDefinition { Guid = new Guid("66707a17-3efc-4306-9e6b-3258b06058e2"), Name = "RTS_Type_Comment", DataType = "TEXT", GroupId = 5, Description = "AI: A comment specific to the family type." },

            // Group 8: Cable Tray
            new SharedParameterDefinition { Guid = new Guid("d9536eef-6f77-47ef-a019-d6620c57353e"), Name = "RTS_Rail_Thickness", DataType = "LENGTH", GroupId = 8, Description = "AI: The thickness of the cable tray side rail." },

            // Group 9: Fire
            new SharedParameterDefinition { Guid = new Guid("a1ff8ef1-1986-4470-9632-5a88e781f447"), Name = "RTS_Fire_Rated", DataType = "YESNO", GroupId = 9, Description = "AI: Specifies if the element has a fire rating." },
            new SharedParameterDefinition { Guid = new Guid("2a7fa3d2-6a9e-43c8-be1c-55af7449a52e"), Name = "RTS_Fire_Rating", DataType = "TEXT", GroupId = 9, Description = "AI: The specific fire rating of the element (e.g., -/60/60)." },

            // Group 10: Internal
            new SharedParameterDefinition { Guid = new Guid("c12c63c1-43c3-4b7d-9518-6cb89f24c421"), Name = "RTS_Calc", DataType = "YESNO", GroupId = 10, IsInstance = true, Visible = false, UserModifiable = false, Description = "Used internally to flag if created by script." },
            new SharedParameterDefinition { Guid = new Guid("f74948b1-4b6a-4859-bb21-62333dd04f21"), Name = "RTS_Flip", DataType = "YESNO", GroupId = 10, IsInstance = true, Visible = false, UserModifiable = false, Description = "Used for indicating when a vetor is to be treated as flipped." },
            new SharedParameterDefinition { Guid = new Guid("bff9351c-7be0-42a0-b685-01635dce858c"), Name = "RTS_Orientation", DataType = "TEXT", GroupId = 10, IsInstance = true, Visible = false, UserModifiable = false, Description = "Used for defining orientation of an element." },
            new SharedParameterDefinition { Guid = new Guid("f04fd842-499d-4567-82be-8d012e8776ac"), Name = "RTS_Original", DataType = "YESNO", GroupId = 10, IsInstance = true, Visible = false, UserModifiable = false, Description = "\"Used when multiple elements, but one is required to be original.\"" }
        };
    }
}
