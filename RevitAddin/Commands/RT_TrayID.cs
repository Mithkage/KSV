// File: RT_TrayID.cs
// Application: Revit 2022 External Command
// Description: This script identifies cable tray elements (horizontal, vertical, or sloped) that have values
//              in any of the RTS_Cable_XX shared parameters (RTS_Cable_01 to RTS_Cable_30).
//              For each identified cable tray, it generates a unique ID based on a prefix, a type suffix,
//              its reference level name, bottom elevation (in mm), and a sequential unique number.
//              This generated ID is then assigned to the 'RTS_ID' shared parameter.
//              This version also iterates through Detail Components, and where their RTS_ID matches
//              a cable tray's RTS_ID, it populates the Detail Item's RTS_Cable_XX parameters
//              with the corresponding Cable Tray's cable values, shifting values up to fill gaps.
//              This script is intended to be compiled as a Revit Add-in.
//              This version includes:
//              - Processing of cable trays regardless of their orientation (horizontal, vertical, or sloped).
//              - Sorting of cable trays to prioritize those occupying lower-indexed RTS_Cable_XX parameters first,
//                then by their values, and finally by all subsequent RTS_Cable_XX values to ensure continuity.
//              - Refined retrieval of bottom elevation using cable tray's absolute Z-coordinate, height parameter,
//                and associated level's elevation. The elevation component in the ID is always a positive value
//                representing the bottom elevation in millimeters relative to the project's internal origin.
//              - Improved and more robust retrieval of the reference level name.
//              - Application of the unique 4-digit number based on identical consecutive RTS_Cable_XX parameter values
//                or sequential numbering for trays without values.
//              - Addition of a prefix derived from the 5th character of the first available RTS_Cable_XX parameter.
//              - Validation of the prefix character: if not alphabetic, it defaults to 'T'.
//              - Addition of a 3-character type suffix based on the cable tray's type name ("FR" -> "FLS", "ESS" -> "ESS", else "DFT").
//              - Writes the unique 4-digit identifier to the 'Branch Number' shared parameter.
//              - Ensures the bottom elevation part of the RTS_ID is always 4 characters long, padded with leading zeros,
//                and represents a positive value.
//              - UPDATED: Handles cable trays without cable values by assigning an 'X' prefix and continuing Branch Numbering
//                         from the last used Branch Number + 1000, then incrementing sequentially for connected trays without values.
//              - NEW: Iterates through Detail Components, and now also Cable Tray Fittings, Conduits, and Conduit Fittings,
//                     matches them by RTS_ID to cable trays, and populates their RTS_Cable_XX parameters,
//                     shifting values to fill gaps. **RTS_ID is NOT generated for these elements.**
//
// RTS_ID Format: [Prefix]-[TypeSuffix]-[LevelAbbr]-[ElevationMM]-[UniqueSuffix]
// Example ID:    P-FLS-L01-4285-0001
//    - Prefix (P):        Derived from the 5th character of the first populated RTS_Cable_XX parameter (e.g., if RTS_Cable_01 = "CABLE-POWER-MAIN", 'P' is used). Defaults to 'X' if no cable values or non-alphabetic character.
//    - TypeSuffix (FLS):  3-character abbreviation based on Cable Tray Type Name (e.g., "FR" in type name becomes "FLS", "ESS" becomes "ESS", otherwise "DFT").
//    - LevelAbbr (L01):   First 3 characters of the associated Reference Level name (e.g., "Level 01" becomes "L01"). Padded with '?' if shorter than 3.
//    - ElevationMM (4285): Bottom elevation of the cable tray in millimeters, rounded to the nearest integer. This value is padded with leading zeros to 4 digits and can be negative.
//    - UniqueSuffix (0001): A sequential 4-digit number. Increments for each unique set of RTS_Cable_XX values. For trays without values, it starts at 1000 + last counter and increments.

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical; // Required for Conduit/ConduitFitting
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace RTS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_TrayIDClass : IExternalCommand
    {
        // Define GUIDs for the shared parameters
        private static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");
        private static readonly Guid BRANCH_NUMBER_GUID = new Guid("3ea1a3bb-8416-45ed-b606-3f3a3f87d4be"); // New: Branch Number GUID

        private static readonly List<Guid> RTS_Cable_GUIDs = new List<Guid>
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
        /// Helper method to find the index and value of the first non-empty RTS_Cable_XX parameter for an element.
        /// </summary>
        /// <param name="elem">The Revit element.</param>
        /// <returns>
        /// A Tuple where Item1 is the 0-indexed position of the first non-empty parameter (or RTS_Cable_GUIDs.Count if none are found),
        /// and Item2 is the string value of that parameter (or string.Empty).
        /// </returns>
        private Tuple<int, string> GetFirstCableParamValue(Element elem)
        {
            for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
            {
                Parameter cableParam = elem.get_Parameter(RTS_Cable_GUIDs[i]);
                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                {
                    // Return 0-indexed position and the value
                    return new Tuple<int, string>(i, cableParam.AsString());
                }
            }
            // If no cable parameters have values, return a large index and empty string
            return new Tuple<int, string>(RTS_Cable_GUIDs.Count, string.Empty);
        }

        /// <summary>
        /// Helper method to get all cable parameter values for sorting or populating.
        /// </summary>
        /// <param name="elem">The Revit element.</param>
        /// <returns>A concatenated string of all cable parameter values, separated by a delimiter.</returns>
        private string GetAllCableParamValues(Element elem)
        {
            StringBuilder sb = new StringBuilder();
            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
            {
                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                sb.Append(cableParam?.AsString() ?? string.Empty);
                sb.Append("|"); // Use a delimiter to prevent accidental matches (e.g., "A" + "B" vs "AB")
            }
            return sb.ToString();
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the active Revit document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Dictionary to store generated RTS_ID and their associated cable values for quick lookup
            // Key: RTS_ID, Value: List of non-empty RTS_Cable_XX values from the Cable Tray
            Dictionary<string, List<string>> cableTrayRtsIdToCableValuesMap = new Dictionary<string, List<string>>();

            // Start a transaction to modify the document
            using (Transaction t = new Transaction(doc, "Update RTS_IDs and Cable Parameters")) // Changed transaction name
            {
                try
                {
                    t.Start();

                    // --- 1. Process Cable Tray elements (generate RTS_ID and collect cable data) ---
                    FilteredElementCollector cableTrayCollector = new FilteredElementCollector(doc);
                    List<Element> allCableTrays = cableTrayCollector
                        .OfCategory(BuiltInCategory.OST_CableTray)
                        .WhereElementIsNotElementType()
                        .ToList();

                    List<Element> cableTraysWithValues = new List<Element>();
                    List<Element> cableTraysWithoutValues = new List<Element>();

                    // Separate cable trays into two groups: with cable values and without.
                    foreach (Element elem in allCableTrays)
                    {
                        bool hasCableValue = false;
                        foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                        {
                            Parameter cableParam = elem.get_Parameter(cableParamGuid);
                            if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                            {
                                hasCableValue = true;
                                break;
                            }
                        }

                        if (hasCableValue)
                        {
                            cableTraysWithValues.Add(elem);
                        }
                        else
                        {
                            cableTraysWithoutValues.Add(elem);
                        }
                    }

                    // *** Process cable trays WITH cable values first ***
                    IOrderedEnumerable<Element> sortedCableTraysWithValues = null;

                    if (cableTraysWithValues.Any())
                    {
                        // Sort primarily by the index of the first non-empty RTS_Cable_XX parameter (lower index first),
                        // then by the value of that first non-empty parameter.
                        sortedCableTraysWithValues = cableTraysWithValues
                            .OrderBy(e => GetFirstCableParamValue(e).Item1) // Sort by the index (e.g., 0 for RTS_Cable_01)
                            .ThenBy(e => GetFirstCableParamValue(e).Item2); // Then by the value of that parameter

                        // Apply subsequent sort criteria (ThenBy) for all RTS_Cable_GUIDs to ensure full consistency
                        // for the unique ID generation (currentCableKey).
                        for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
                        {
                            Guid currentGuid = RTS_Cable_GUIDs[i];
                            sortedCableTraysWithValues = sortedCableTraysWithValues
                                .ThenBy(e => e.get_Parameter(currentGuid)?.AsString() ?? string.Empty);
                        }
                    }
                    else
                    {
#if REVIT2024_OR_GREATER
                        sortedCableTraysWithValues = Enumerable.Empty<Element>().OrderBy(e => e.Id.Value);
#else
                        sortedCableTraysWithValues = Enumerable.Empty<Element>().OrderBy(e => e.Id.IntegerValue);
#endif
                    }

                    int updatedCableTrayCount = 0;
                    int uniqueIdCounter = 0; // Initialize counter for cable trays with values
                    string previousCableKey = null; // Stores the concatenated values of RTS_Cable_XX for the previous element

                    foreach (Element elem in sortedCableTraysWithValues) // Iterate through the sorted list
                    {
                        CableTray cableTray = elem as CableTray;
                        if (cableTray == null) continue;

                        Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                        Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                        if (rtsIdParam != null && !rtsIdParam.IsReadOnly) // Ensure parameter is writeable
                        {
                            // 1. Determine the prefix from the 5th character of the first RTS_Cable_XX with a value
                            string prefixChar = "X"; // Default prefix character, will be overridden if cable values exist
                            List<string> currentCableTrayValues = new List<string>(); // Store non-empty cable values for the map

                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    currentCableTrayValues.Add(paramValue); // Add to list for map

                                    if (paramValue.Length >= 5)
                                    {
                                        char charAtIndex4 = paramValue[4]; // Get the 5th character (index 4)
                                        if (char.IsLetter(charAtIndex4))
                                        {
                                            prefixChar = charAtIndex4.ToString().ToUpper(); // Ensure uppercase
                                        }
                                        else
                                        {
                                            prefixChar = "T"; // Use "T" if not alphabetic
                                        }
                                    }
                                    // Found the first non-empty, determine prefix, no need to break here as we need all values for currentCableTrayValues
                                }
                            }
                            // If no cable values were found, prefixChar remains 'X'

                            // 2. Determine the 3-character type suffix
                            string trayTypeSuffix = "DFT"; // Default suffix
                            ElementType cableTrayType = doc.GetElement(elem.GetTypeId()) as ElementType;
                            if (cableTrayType != null && !string.IsNullOrEmpty(cableTrayType.Name))
                            {
                                string typeName = cableTrayType.Name.ToUpper();
                                if (typeName.Contains("FR"))
                                {
                                    trayTypeSuffix = "FLS";
                                }
                                else if (typeName.Contains("ESS"))
                                {
                                    trayTypeSuffix = "ESS";
                                }
                                // Else, remains "DFT"
                            }

                            // 3. Get the first 3 characters of the Reference Level
                            string levelAbbreviation = "???"; // Default to "???"
                            Level associatedLevel = null;
                            ElementId levelElementId = ElementId.InvalidElementId;

                            // Try to get LevelId from RBS_START_LEVEL_PARAM first
                            Parameter startLevelParam = elem.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                            if (startLevelParam != null && startLevelParam.HasValue)
                            {
                                levelElementId = startLevelParam.AsElementId();
                            }
                            // Fallback to Element.LevelId if RBS_START_LEVEL_PARAM didn't provide a valid ID
                            if (levelElementId == ElementId.InvalidElementId && elem.LevelId != ElementId.InvalidElementId)
                            {
                                levelElementId = elem.LevelId;
                            }

                            if (levelElementId != ElementId.InvalidElementId)
                            {
                                associatedLevel = doc.GetElement(levelElementId) as Level;
                                if (associatedLevel != null && !string.IsNullOrEmpty(associatedLevel.Name))
                                {
                                    string levelName = associatedLevel.Name;
                                    levelAbbreviation = levelName.Length >= 3 ? levelName.Substring(0, 3).ToUpper() : levelName.ToUpper().PadRight(3, '?');
                                }
                            }

                            // --- Calculate tray bottom elevation relative to tray level ---
                            double offsetFromLevelFeet = 0.0;
                            Parameter offsetParam = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                            if (offsetParam != null && offsetParam.HasValue)
                            {
                                offsetFromLevelFeet = offsetParam.AsDouble();
                            }

                            double trayHeightFeet = 0.0;
                            Parameter heightParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                            if (heightParam != null && heightParam.HasValue)
                            {
                                trayHeightFeet = heightParam.AsDouble();
                            }

                            // Bottom elevation relative to level (in feet)
                            double bottomElevationFromLevelFeet = offsetFromLevelFeet - trayHeightFeet / 2.0;

                            // Convert to millimeters (allow negative values)
                            long bottomElevationFromLevelMM = (long)Math.Round(bottomElevationFromLevelFeet * 304.8);

                            // Format as 4 digits, preserving sign if negative
                            string formattedBottomElevation = bottomElevationFromLevelMM.ToString("D4");

                            // 5. Generate a unique 4-digit number based on the sorted order and parameter values
                            string currentCableKey = GetAllCableParamValues(elem);

                            if (uniqueIdCounter == 0 || currentCableKey != previousCableKey)
                            {
                                uniqueIdCounter++;
                            }
                            previousCableKey = currentCableKey; // Update the previous key for the next iteration

                            string uniqueSuffix = uniqueIdCounter.ToString("D4"); // Formats to 0001, 0002, etc.

                            string newRtsId = $"{prefixChar}-{trayTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
                            rtsIdParam.Set(newRtsId);

                            if (branchNumberParam != null && !branchNumberParam.IsReadOnly)
                            {
                                branchNumberParam.Set(uniqueSuffix);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have a writeable 'Branch Number' shared parameter or it's missing.");
                            }
                            updatedCableTrayCount++;

                            // Store the generated RTS_ID and the collected cable values in the map
                            if (!cableTrayRtsIdToCableValuesMap.ContainsKey(newRtsId))
                            {
                                cableTrayRtsIdToCableValuesMap.Add(newRtsId, currentCableTrayValues);
                            }
                            else
                            {
                                // This should ideally not happen if RTS_ID is truly unique
                                System.Diagnostics.Debug.WriteLine($"Warning: Duplicate RTS_ID '{newRtsId}' generated for Cable Tray {elem.Id}. Overwriting cable values in map.");
                                cableTrayRtsIdToCableValuesMap[newRtsId] = currentCableTrayValues;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have a writeable 'RTS_ID' shared parameter or it's missing.");
                        }
                    }

                    // *** Process cable trays WITHOUT cable values ***
                    if (cableTraysWithoutValues.Any())
                    {
                        // Start the branch number sequence for trays without values.
                        // Increment by 1000 from the last used counter for trays with values, or start at 1000 if none.
                        int noCableUniqueIdCounter = uniqueIdCounter == 0 ? 1000 : (uniqueIdCounter / 1000 + 1) * 1000;

                        List<Element> sortedCableTraysWithoutValues = cableTraysWithoutValues
                            .OrderBy(e =>
                            {
                                // Attempt to get level for sorting
                                ElementId levelId = ElementId.InvalidElementId;
                                Parameter startLevelParam = e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                                if (startLevelParam != null && startLevelParam.HasValue)
                                {
                                    levelId = startLevelParam.AsElementId();
                                }
                                if (levelId == ElementId.InvalidElementId && e.LevelId != ElementId.InvalidElementId)
                                {
                                    levelId = e.LevelId;
                                }
#if REVIT2024_OR_GREATER
                                return levelId.Value;
#else
                                return levelId.IntegerValue;
#endif
                            })
                            .ThenBy(e =>
                            {
                                // Sort by Z-coordinate, then X, then Y for approximate spatial grouping
                                if (e.Location is LocationCurve lc) return lc.Curve.GetEndPoint(0).Z;
                                if (e.Location is LocationPoint lp) return lp.Point.Z;
                                return 0.0; // Default if no location
                            })
                            .ThenBy(e =>
                            {
                                if (e.Location is LocationCurve lc) return lc.Curve.GetEndPoint(0).X;
                                if (e.Location is LocationPoint lp) return lp.Point.X;
                                return 0.0;
                            })
                            .ThenBy(e =>
                            {
                                if (e.Location is LocationCurve lc) return lc.Curve.GetEndPoint(0).Y;
                                if (e.Location is LocationPoint lp) return lp.Point.Y;
                                return 0.0;
                            })
                            .ToList();

                        foreach (Element elem in sortedCableTraysWithoutValues)
                        {
                            CableTray cableTray = elem as CableTray;
                            if (cableTray == null) continue;

                            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                            if (rtsIdParam != null && !rtsIdParam.IsReadOnly)
                            {
                                // Prefix for trays without cable values defaults to 'X'
                                string prefixChar = "X";

                                // Determine the 3-character type suffix (same logic as before)
                                string trayTypeSuffix = "DFT";
                                ElementType cableTrayType = doc.GetElement(elem.GetTypeId()) as ElementType;
                                if (cableTrayType != null && !string.IsNullOrEmpty(cableTrayType.Name))
                                {
                                    string typeName = cableTrayType.Name.ToUpper();
                                    if (typeName.Contains("FR"))
                                    {
                                        trayTypeSuffix = "FLS";
                                    }
                                    else if (typeName.Contains("ESS"))
                                    {
                                        trayTypeSuffix = "ESS";
                                    }
                                }

                                // Get the first 3 characters of the Reference Level (same logic as before)
                                string levelAbbreviation = "???";
                                Level associatedLevel = null;
                                ElementId levelElementId = ElementId.InvalidElementId;
                                Parameter startLevelParam = elem.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                                if (startLevelParam != null && startLevelParam.HasValue)
                                {
                                    levelElementId = startLevelParam.AsElementId();
                                }
                                if (levelElementId == ElementId.InvalidElementId && elem.LevelId != ElementId.InvalidElementId)
                                {
                                    levelElementId = elem.LevelId;
                                }

                                if (levelElementId != ElementId.InvalidElementId)
                                {
                                    associatedLevel = doc.GetElement(levelElementId) as Level;
                                    if (associatedLevel != null && !string.IsNullOrEmpty(associatedLevel.Name))
                                    {
                                        string levelName = associatedLevel.Name;
                                        levelAbbreviation = levelName.Length >= 3 ? levelName.Substring(0, 3).ToUpper() : levelName.ToUpper().PadRight(3, '?');
                                    }
                                }

                                // --- Calculate tray bottom elevation relative to tray level ---
                                double offsetFromLevelFeet = 0.0;
                                Parameter offsetParam = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (offsetParam != null && offsetParam.HasValue)
                                {
                                    offsetFromLevelFeet = offsetParam.AsDouble();
                                }

                                double trayHeightFeet = 0.0;
                                Parameter heightParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                                if (heightParam != null && heightParam.HasValue)
                                {
                                    trayHeightFeet = heightParam.AsDouble();
                                }

                                // Bottom elevation relative to level (in feet)
                                double bottomElevationFromLevelFeet = offsetFromLevelFeet - trayHeightFeet / 2.0;

                                // Convert to millimeters (allow negative values)
                                long bottomElevationFromLevelMM = (long)Math.Round(bottomElevationFromLevelFeet * 304.8);

                                // Format as 4 digits, preserving sign if negative
                                string formattedBottomElevation = bottomElevationFromLevelMM.ToString("D4");

                                // Increment the counter for trays without values
                                string uniqueSuffix = noCableUniqueIdCounter.ToString("D4");
                                noCableUniqueIdCounter++;

                                string newRtsId = $"{prefixChar}-{trayTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
                                rtsIdParam.Set(newRtsId);

                                if (branchNumberParam != null && !branchNumberParam.IsReadOnly)
                                {
                                    branchNumberParam.Set(uniqueSuffix);
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have a writeable 'Branch Number' shared parameter or it's missing.");
                                }
                                updatedCableTrayCount++;

                                // For trays without cable values, their cable values list will be empty
                                if (!cableTrayRtsIdToCableValuesMap.ContainsKey(newRtsId))
                                {
                                    cableTrayRtsIdToCableValuesMap.Add(newRtsId, new List<string>());
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Warning: Duplicate RTS_ID '{newRtsId}' generated for Cable Tray {elem.Id} (no values). Overwriting cable values in map.");
                                    cableTrayRtsIdToCableValuesMap[newRtsId] = new List<string>(); // Ensure it's empty
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have a writeable 'RTS_ID' shared parameter or it's missing.");
                            }
                        }
                    }

                    // --- 2. Process Detail Components, Cable Tray Fittings, Conduits, and Conduit Fittings ---
                    // These elements will have their RTS_Cable_XX parameters ordered/compacted
                    // based on a matching RTS_ID found on a Cable Tray. No new RTS_ID will be generated.
                    var categoriesToUpdateCableParams = new List<BuiltInCategory>
                    {
                        BuiltInCategory.OST_DetailComponents,
                        BuiltInCategory.OST_CableTrayFitting,
                        BuiltInCategory.OST_Conduit,
                        BuiltInCategory.OST_ConduitFitting
                    };
                    var filterForCableParamUpdate = new ElementMulticategoryFilter(categoriesToUpdateCableParams);

                    var otherElementsToUpdate = new FilteredElementCollector(doc)
                        .WherePasses(filterForCableParamUpdate)
                        .WhereElementIsNotElementType()
                        .ToList();

                    int otherElementsUpdatedCount = 0;

                    foreach (Element currentElement in otherElementsToUpdate)
                    {
                        Parameter elementRtsIdParam = currentElement.get_Parameter(RTS_ID_GUID);

                        if (elementRtsIdParam != null && !string.IsNullOrEmpty(elementRtsIdParam.AsString()))
                        {
                            string elementRtsId = elementRtsIdParam.AsString();

                            // Try to find a matching cable tray's RTS_ID
                            if (cableTrayRtsIdToCableValuesMap.TryGetValue(elementRtsId, out List<string> cableTrayValues))
                            {
                                System.Diagnostics.Debug.WriteLine($"Processing {currentElement.Category.Name} {currentElement.Id} with RTS_ID '{elementRtsId}'. Matching cable tray found.");

                                // Sort the cable values obtained from the cable tray (already done by the TrayID part)
                                // and apply them to the current element's RTS_Cable_XX parameters, filling gaps.
                                // The cableTrayValues List is already guaranteed to be non-null and correctly ordered from how it was built.

                                int cableValueIndex = 0;
                                for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
                                {
                                    Parameter elementCableParam = currentElement.get_Parameter(RTS_Cable_GUIDs[i]);
                                    if (elementCableParam != null && !elementCableParam.IsReadOnly) // Ensure parameter is writeable
                                    {
                                        if (cableValueIndex < cableTrayValues.Count)
                                        {
                                            // Set the parameter with the next available cable tray value
                                            // Only set if the value is different to avoid unnecessary modifications
                                            if (elementCableParam.AsString() != cableTrayValues[cableValueIndex])
                                            {
                                                elementCableParam.Set(cableTrayValues[cableValueIndex]);
                                                System.Diagnostics.Debug.WriteLine($"  - Set RTS_Cable_{i + 1:D2} on {currentElement.Category.Name} {currentElement.Id} to: '{cableTrayValues[cableValueIndex]}'");
                                            }
                                            cableValueIndex++;
                                        }
                                        else
                                        {
                                            // If no more values from cable tray, clear the remaining parameters on the current element
                                            if (!string.IsNullOrEmpty(elementCableParam.AsString()))
                                            {
                                                elementCableParam.Set(string.Empty); // Clear existing value
                                                System.Diagnostics.Debug.WriteLine($"  - Cleared RTS_Cable_{i + 1:D2} on {currentElement.Category.Name} {currentElement.Id}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine($"  - Warning: {currentElement.Category.Name} {currentElement.Id} missing writeable RTS_Cable_{i + 1:D2} parameter.");
                                    }
                                }
                                otherElementsUpdatedCount++;
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"{currentElement.Category.Name} {currentElement.Id} has RTS_ID '{elementRtsId}' but no matching Cable Tray found. Not updating cable parameters.");
                                // Option: You might want to clear cable parameters of elements that don't match any tray.
                                // For now, they are left as is if no match.
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Warning: {currentElement.Category.Name} {currentElement.Id} does not have a valid 'RTS_ID' shared parameter value. Skipping cable parameter update.");
                        }
                    }


                    t.Commit();
                    TaskDialog.Show("RT_TrayID", $"{updatedCableTrayCount} Cable Trays and {otherElementsUpdatedCount} other elements (Detail Items, Fittings, Conduits) updated successfully.");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = $"Error: {ex.Message}";
                    TaskDialog.Show("RT_TrayID Error", message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
    }
}
