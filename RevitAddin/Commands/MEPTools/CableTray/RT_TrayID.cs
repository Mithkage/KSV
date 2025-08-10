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
//              - Addition of a two-character prefix based on the cable tray's Type Name (e.g., "LV_SUB-A" -> "LA").
//              - Validation of the prefix: defaults to 'XX' if no Type Name matches defined criteria.
//              - Addition of a 3-character type suffix based on the cable tray's type name ("FR" -> "FLS", "ESS" -> "ESS", else "DFT").
//              - Writes the unique 4-digit identifier to the 'Branch Number' shared parameter.
//              - Ensures the bottom elevation part of the RTS_ID is always 4 characters long, padded with leading zeros,
//                and represents a positive value.
//              - UPDATED: Handles cable trays without cable values by continuing Branch Numbering
//                         from the last used Branch Number + 1000, then incrementing sequentially for connected trays without values.
//              - NEW: Iterates through Detail Components, and now also Cable Tray Fittings, Conduits, and Conduit Fittings,
//                     matches them by RTS_ID to cable trays, and populates their RTS_Cable_XX parameters,
//                     shifting values to fill gaps. **RTS_ID is NOT generated for these elements.**
//
// RTS_ID Format: [Prefix]-[TypeSuffix]-[LevelAbbr]-[ElevationMM]-[UniqueSuffix]
// Example ID:    LA-FLS-L01-4285-0001
//   - Prefix (LA):         2-character prefix based on Type Name (e.g., "LV_SUB-A" -> "LA"). Defaults to "XX" if no criteria are met.
//   - TypeSuffix (FLS):  3-character abbreviation based on Cable Tray Type Name (e.g., "FR" in type name becomes "FLS", "ESS" becomes "ESS", otherwise "DFT").
//   - LevelAbbr (L01):   First 3 characters of the associated Reference Level name (e.g., "Level 01" becomes "L01"). Padded with '?' if shorter than 3.
//   - ElevationMM (4285): Bottom elevation of the cable tray in millimeters, rounded to the nearest integer. This value is padded with leading zeros to 4 digits and can be negative.
//   - UniqueSuffix (0001): A sequential 4-digit number. Increments for each unique set of RTS_Cable_XX values. For trays without values, it starts at 1000 + last counter and increments.

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical; // Required for Conduit/ConduitFitting
using Autodesk.Revit.UI;
using RTS.Utilities; // Add reference to the utilities namespace for shared parameters
#endregion

namespace RTS.Commands.MEPTools.CableTray
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_TrayIDClass : IExternalCommand
    {
        // Use the shared parameter GUIDs from the utility class
        private static readonly Guid RTS_ID_GUID = SharedParameters.Cable.RTS_ID_GUID;
        private static readonly Guid BRANCH_NUMBER_GUID = SharedParameters.Cable.BRANCH_NUMBER_GUID;
        private static readonly List<Guid> RTS_Cable_GUIDs = SharedParameters.Cable.RTS_Cable_GUIDs;

        // Define the mapping from Type Name to RTS_ID prefix
        private static readonly Dictionary<string, string> TypeNameToPrefixMapping = new Dictionary<string, string>
        {
            { "LV_SUB-A", "LA" },
            { "LV_SUB-B", "LB" },
            { "_LV_SPARE", "LX" },
            { "LV_UPS-L", "UL" },
            { "LV_UPS-P", "UP" },
            { "_SHARED_LV", "LS" },
            { "LV_GEN-A", "GA" },
            { "LV_GEN-B", "GB" },
            { "LV_EARTHING", "LE" },
            { "_HV", "HV" }
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

        /// <summary>
        /// Helper method to determine the two-character prefix based on the Type Name of the cable tray.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="elem">The cable tray element.</param>
        /// <returns>A two-character prefix string based on the mapping.</returns>
        private string GetPrefixFromTypeName(Document doc, Element elem)
        {
            // Default prefix if no criteria are met
            string prefixChars = "XX";

            ElementType cableTrayType = doc.GetElement(elem.GetTypeId()) as ElementType;
            if (cableTrayType != null && !string.IsNullOrEmpty(cableTrayType.Name))
            {
                string typeName = cableTrayType.Name.ToUpper();

                // Check if the type name contains any of the keys in our mapping
                foreach (var mapping in TypeNameToPrefixMapping)
                {
                    if (typeName.Contains(mapping.Key))
                    {
                        return mapping.Value;
                    }
                }
            }

            return prefixChars;
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
                        // FIX: Use fully qualified name to avoid namespace collision
                        Autodesk.Revit.DB.Electrical.CableTray cableTray = elem as Autodesk.Revit.DB.Electrical.CableTray;
                        if (cableTray == null) continue;

                        Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                        Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                        if (rtsIdParam != null && !rtsIdParam.IsReadOnly) // Ensure parameter is writeable
                        {
                            // 1. Determine the prefix based on the Type Name of the cable tray
                            string prefixChars = GetPrefixFromTypeName(doc, elem);

                            List<string> currentCableTrayValues = new List<string>(); // Store non-empty cable values for the map

                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    currentCableTrayValues.Add(paramValue); // Add to list for map
                                }
                            }

                            // 2. Determine the 3-character type suffix
                            string trayTypeSuffix = "DFT"; // Default suffix
                            ElementType cableTrayType = doc.GetElement(elem.GetTypeId()) as ElementType;
                            if (cableTrayType != null && !string.IsNullOrEmpty(cableTrayType.Name))
                            {
                                string typeName = cableTrayType.Name.ToUpper();
                                if (typeName.Contains("FR") || typeName.Contains("FIRE"))
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

                            string newRtsId = $"{prefixChars}-{trayTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
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
                    // Declare the counter at a higher scope so it's accessible later
                    int noCableUniqueIdCounter = uniqueIdCounter == 0 ? 1000 : (uniqueIdCounter / 1000 + 1) * 1000;

                    if (cableTraysWithoutValues.Any())
                    {
                        // Start the branch number sequence for trays without values.
                        // Increment by 1000 from the last used counter for trays with values, or start at 1000 if none.

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
                            // FIX: Use fully qualified name to avoid namespace collision
                            Autodesk.Revit.DB.Electrical.CableTray cableTray = elem as Autodesk.Revit.DB.Electrical.CableTray;
                            if (cableTray == null) continue;

                            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                            if (rtsIdParam != null && !rtsIdParam.IsReadOnly)
                            {
                                // Use the same method for determining the prefix from Type Name
                                string prefixChars = GetPrefixFromTypeName(doc, elem);

                                // Determine the 3-character type suffix (same logic as before)
                                string trayTypeSuffix = "DFT";
                                ElementType cableTrayType = doc.GetElement(elem.GetTypeId()) as ElementType;
                                if (cableTrayType != null && !string.IsNullOrEmpty(cableTrayType.Name))
                                {
                                    string typeName = cableTrayType.Name.ToUpper();
                                    if (typeName.Contains("FR") || typeName.Contains("FIRE"))
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

                                string newRtsId = $"{prefixChars}-{trayTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
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

                    // --- 3. Process Conduit elements (generate RTS_ID using same pattern as cable trays) ---
                    FilteredElementCollector conduitCollector = new FilteredElementCollector(doc);
                    List<Element> allConduits = conduitCollector
                        .OfCategory(BuiltInCategory.OST_Conduit)
                        .WhereElementIsNotElementType()
                        .ToList();

                    List<Element> conduitsWithValues = new List<Element>();
                    List<Element> conduitsWithoutValues = new List<Element>();

                    // Separate conduits into two groups: with cable values and without.
                    foreach (Element elem in allConduits)
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
                            conduitsWithValues.Add(elem);
                        }
                        else
                        {
                            conduitsWithoutValues.Add(elem);
                        }
                    }

                    // *** Process conduits WITH cable values first ***
                    IOrderedEnumerable<Element> sortedConduitsWithValues = null;

                    if (conduitsWithValues.Any())
                    {
                        // Sort primarily by the index of the first non-empty RTS_Cable_XX parameter (lower index first),
                        // then by the value of that first non-empty parameter.
                        sortedConduitsWithValues = conduitsWithValues
                            .OrderBy(e => GetFirstCableParamValue(e).Item1) // Sort by the index (e.g., 0 for RTS_Cable_01)
                            .ThenBy(e => GetFirstCableParamValue(e).Item2); // Then by the value of that parameter

                        // Apply subsequent sort criteria (ThenBy) for all RTS_Cable_GUIDs to ensure full consistency
                        // for the unique ID generation (currentCableKey).
                        for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
                        {
                            Guid currentGuid = RTS_Cable_GUIDs[i];
                            sortedConduitsWithValues = sortedConduitsWithValues
                                .ThenBy(e => e.get_Parameter(currentGuid)?.AsString() ?? string.Empty);
                        }
                    }
                    else
                    {
#if REVIT2024_OR_GREATER
                        sortedConduitsWithValues = Enumerable.Empty<Element>().OrderBy(e => e.Id.Value);
#else
                        sortedConduitsWithValues = Enumerable.Empty<Element>().OrderBy(e => e.Id.IntegerValue);
#endif
                    }

                    int updatedConduitCount = 0;
                    string previousConduitCableKey = null; // Stores the concatenated values of RTS_Cable_XX for the previous element

                    foreach (Element elem in sortedConduitsWithValues) // Iterate through the sorted list
                    {
                        Conduit conduit = elem as Conduit;
                        if (conduit == null) continue;

                        Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                        Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                        if (rtsIdParam != null && !rtsIdParam.IsReadOnly) // Ensure parameter is writeable
                        {
                            // 1. Determine the prefix based on the Type Name of the conduit (same logic as cable tray)
                            string prefixChars = GetPrefixFromTypeName(doc, elem);

                            List<string> currentConduitValues = new List<string>(); // Store non-empty cable values for the map

                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
                                    currentConduitValues.Add(paramValue); // Add to list for map
                                }
                            }

                            // 2. Determine the 3-character type suffix
                            string conduitTypeSuffix = "DFT"; // Default suffix
                            ElementType conduitType = doc.GetElement(elem.GetTypeId()) as ElementType;
                            if (conduitType != null && !string.IsNullOrEmpty(conduitType.Name))
                            {
                                string typeName = conduitType.Name.ToUpper();
                                if (typeName.Contains("FR") || typeName.Contains("FIRE"))
                                {
                                    conduitTypeSuffix = "FLS";
                                }
                                else if (typeName.Contains("ESS"))
                                {
                                    conduitTypeSuffix = "ESS";
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

                            // --- Calculate conduit bottom elevation ---
                            double offsetFromLevelFeet = 0.0;
                            Parameter offsetParam = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                            if (offsetParam != null && offsetParam.HasValue)
                            {
                                offsetFromLevelFeet = offsetParam.AsDouble();
                            }

                            double conduitDiameterFeet = 0.0;
                            Parameter diameterParam = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                            if (diameterParam != null && diameterParam.HasValue)
                            {
                                conduitDiameterFeet = diameterParam.AsDouble();
                            }

                            // Bottom elevation for a conduit is center point minus radius
                            double bottomElevationFromLevelFeet = offsetFromLevelFeet - conduitDiameterFeet / 2.0;

                            // Convert to millimeters
                            long bottomElevationFromLevelMM = (long)Math.Round(bottomElevationFromLevelFeet * 304.8);

                            // Format as 4 digits, preserving sign if negative
                            string formattedBottomElevation = bottomElevationFromLevelMM.ToString("D4");

                            // 5. Generate a unique 4-digit number based on the sorted order and parameter values
                            // Use the same counter that was used for cable trays to ensure uniqueness across both types
                            string currentCableKey = GetAllCableParamValues(elem);

                            if (uniqueIdCounter == 0 || currentCableKey != previousConduitCableKey)
                            {
                                uniqueIdCounter++;
                            }
                            previousConduitCableKey = currentCableKey; // Update the previous key for the next iteration

                            string uniqueSuffix = uniqueIdCounter.ToString("D4"); // Formats to 0001, 0002, etc.

                            string newRtsId = $"{prefixChars}-{conduitTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
                            rtsIdParam.Set(newRtsId);

                            if (branchNumberParam != null && !branchNumberParam.IsReadOnly)
                            {
                                branchNumberParam.Set(uniqueSuffix);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Warning: Conduit {elem.Id} does not have a writeable 'Branch Number' shared parameter or it's missing.");
                            }
                            updatedConduitCount++;

                            // Store the generated RTS_ID and the collected cable values in the map
                            if (!cableTrayRtsIdToCableValuesMap.ContainsKey(newRtsId))
                            {
                                cableTrayRtsIdToCableValuesMap.Add(newRtsId, currentConduitValues);
                            }
                            else
                            {
                                // This should ideally not happen if RTS_ID is truly unique
                                System.Diagnostics.Debug.WriteLine($"Warning: Duplicate RTS_ID '{newRtsId}' generated for Conduit {elem.Id}. Overwriting cable values in map.");
                                cableTrayRtsIdToCableValuesMap[newRtsId] = currentConduitValues;
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Warning: Conduit {elem.Id} does not have a writeable 'RTS_ID' shared parameter or it's missing.");
                        }
                    }

                    // *** Process conduits WITHOUT cable values ***
                    if (conduitsWithoutValues.Any())
                    {
                        // Continuation of the uniqueIdCounter used by cable trays
                        int noConduitValuesUniqueIdCounter = noCableUniqueIdCounter;

                        List<Element> sortedConduitsWithoutValues = conduitsWithoutValues
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

                        foreach (Element elem in sortedConduitsWithoutValues)
                        {
                            Conduit conduit = elem as Conduit;
                            if (conduit == null) continue;

                            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

                            if (rtsIdParam != null && !rtsIdParam.IsReadOnly)
                            {
                                // Use the same method for determining the prefix from Type Name
                                string prefixChars = GetPrefixFromTypeName(doc, elem);

                                // Determine the 3-character type suffix (same logic as before)
                                string conduitTypeSuffix = "DFT";
                                ElementType conduitType = doc.GetElement(elem.GetTypeId()) as ElementType;
                                if (conduitType != null && !string.IsNullOrEmpty(conduitType.Name))
                                {
                                    string typeName = conduitType.Name.ToUpper();
                                    if (typeName.Contains("FR") || typeName.Contains("FIRE"))
                                    {
                                        conduitTypeSuffix = "FLS";
                                    }
                                    else if (typeName.Contains("ESS"))
                                    {
                                        conduitTypeSuffix = "ESS";
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

                                // --- Calculate conduit bottom elevation ---
                                double offsetFromLevelFeet = 0.0;
                                Parameter offsetParam = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (offsetParam != null && offsetParam.HasValue)
                                {
                                    offsetFromLevelFeet = offsetParam.AsDouble();
                                }

                                double conduitDiameterFeet = 0.0;
                                Parameter diameterParam = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                                if (diameterParam != null && diameterParam.HasValue)
                                {
                                    conduitDiameterFeet = diameterParam.AsDouble();
                                }

                                // Bottom elevation relative to level (in feet)
                                double bottomElevationFromLevelFeet = offsetFromLevelFeet - conduitDiameterFeet / 2.0;

                                // Convert to millimeters (allow negative values)
                                long bottomElevationFromLevelMM = (long)Math.Round(bottomElevationFromLevelFeet * 304.8);

                                // Format as 4 digits, preserving sign if negative
                                string formattedBottomElevation = bottomElevationFromLevelMM.ToString("D4");

                                // Increment the counter for conduits without values
                                string uniqueSuffix = noConduitValuesUniqueIdCounter.ToString("D4");
                                noConduitValuesUniqueIdCounter++;

                                string newRtsId = $"{prefixChars}-{conduitTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";
                                rtsIdParam.Set(newRtsId);

                                if (branchNumberParam != null && !branchNumberParam.IsReadOnly)
                                {
                                    branchNumberParam.Set(uniqueSuffix);
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Warning: Conduit {elem.Id} does not have a writeable 'Branch Number' shared parameter or it's missing.");
                                }
                                updatedConduitCount++;

                                // For conduits without cable values, their cable values list will be empty
                                if (!cableTrayRtsIdToCableValuesMap.ContainsKey(newRtsId))
                                {
                                    cableTrayRtsIdToCableValuesMap.Add(newRtsId, new List<string>());
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Warning: Duplicate RTS_ID '{newRtsId}' generated for Conduit {elem.Id} (no values). Overwriting cable values in map.");
                                    cableTrayRtsIdToCableValuesMap[newRtsId] = new List<string>(); // Ensure it's empty
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Warning: Conduit {elem.Id} does not have a writeable 'RTS_ID' shared parameter or it's missing.");
                            }
                        }

                        // Update noCableUniqueIdCounter for potential future use
                        noCableUniqueIdCounter = noConduitValuesUniqueIdCounter;
                    }

                    t.Commit();
                    TaskDialog.Show("RT_TrayID", $"{updatedCableTrayCount} Cable Trays, {updatedConduitCount} Conduits, and {otherElementsUpdatedCount} other elements (Detail Items, Fittings) updated successfully.");
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
