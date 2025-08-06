using System;
using System.Collections.Generic;

namespace RTS.Utilities
{
    /// <summary>
    /// Static utility class containing GUID references for shared parameters used across the application.
    /// </summary>
    public static class SharedParameters
    {
        /// <summary>
        /// Cable-related shared parameter GUIDs
        /// </summary>
        public static class Cable
        {
            /// <summary>
            /// GUID for RTS_ID shared parameter
            /// </summary>
            public static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
            
            /// <summary>
            /// GUID for Branch Number shared parameter
            /// </summary>
            public static readonly Guid BRANCH_NUMBER_GUID = new Guid("3ea1a3bb-8416-45ed-b606-3f3a3f87d4be");

            /// <summary>
            /// Collection of GUIDs for RTS_Cable_01 through RTS_Cable_30 shared parameters
            /// </summary>
            public static readonly List<Guid> RTS_Cable_GUIDs = new List<Guid>
            {
                new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), // RTS_Cable_01
                new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), // RTS_Cable_02
                new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"), // RTS_Cable_03
                new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), // RTS_Cable_04
                new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"), // RTS_Cable_05
                new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), // RTS_Cable_06
                new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"), // RTS_Cable_07
                new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), // RTS_Cable_08
                new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"), // RTS_Cable_09
                new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), // RTS_Cable_10
                new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"), // RTS_Cable_11
                new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), // RTS_Cable_12
                new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"), // RTS_Cable_13
                new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), // RTS_Cable_14
                new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"), // RTS_Cable_15
                new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), // RTS_Cable_16
                new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"), // RTS_Cable_17
                new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), // RTS_Cable_18
                new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"), // RTS_Cable_19
                new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), // RTS_Cable_20
                new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"), // RTS_Cable_21
                new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), // RTS_Cable_22
                new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"), // RTS_Cable_23
                new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), // RTS_Cable_24
                new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"), // RTS_Cable_25
                new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), // RTS_Cable_26
                new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"), // RTS_Cable_27
                new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), // RTS_Cable_28
                new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"), // RTS_Cable_29
                new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b")  // RTS_Cable_30
            };
            
            /// <summary>
            /// Gets the RTS_Cable parameter GUID by its index (1-30)
            /// </summary>
            /// <param name="index">The 1-based index of the cable parameter (1-30)</param>
            /// <returns>The GUID for the specified RTS_Cable parameter</returns>
            /// <exception cref="ArgumentOutOfRangeException">Thrown when index is less than 1 or greater than 30</exception>
            public static Guid GetRTS_CableGUIDByIndex(int index)
            {
                if (index < 1 || index > 30)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and 30");
                
                return RTS_Cable_GUIDs[index - 1];
            }
        }
        
        // Add other shared parameter categories here as needed, for example:
        /*
        /// <summary>
        /// Fire-related shared parameter GUIDs
        /// </summary>
        public static class Fire
        {
            // Fire-related parameter GUIDs
        }
        
        /// <summary>
        /// Structural-related shared parameter GUIDs
        /// </summary>
        public static class Structural
        {
            // Structural-related parameter GUIDs
        }
        */
    }
}