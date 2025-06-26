// File: RT_TrayID.cs
// Application: Revit 2022 External Command
// Description: This script identifies cable tray elements (horizontal, vertical, or sloped) that have values
//              in any of the RTS_Cable_XX shared parameters (RTS_Cable_01 to RTS_Cable_30).
//              For each identified cable tray, it generates a unique ID based on a prefix, a type suffix,
//              its reference level name, bottom elevation (in mm), and a sequential unique number.
//              This generated ID is then assigned to the 'RTS_ID' shared parameter.
//              This script is intended to be compiled as a Revit Add-in.
//              This version includes:
//              - Processing of cable trays regardless of their orientation (horizontal, vertical, or sloped).
//              - NEW: Sorting of cable trays to prioritize those occupying lower-indexed RTS_Cable_XX parameters first,
//                     then by their values, and finally by all subsequent RTS_Cable_XX values to ensure continuity.
//              - Refined retrieval of bottom elevation using cable tray's Z-coordinate, height parameter,
//                and associated level's elevation.
//              - Improved and more robust retrieval of the reference level name.
//              - Application of the unique 4-digit number based on identical consecutive RTS_Cable_XX parameter values.
//              - Addition of a prefix derived from the 5th character of the first available RTS_Cable_XX parameter.
//              - Validation of the prefix character: if not alphabetic, it defaults to 'T'.
//              - Addition of a 3-character type suffix based on the cable tray's type name ("FR" -> "FLS", "ESS" -> "ESS", else "DFT").
//              - Writes the unique 4-digit identifier to the 'Branch Number' shared parameter.
//              - Ensures the bottom elevation part of the RTS_ID is always 4 characters long, padded with leading zeros.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical; // Required for CableTray class
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RT_TrayID
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
        /// <param name="elem">The Revit element (Cable Tray).</param>
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

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the active Revit document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Start a transaction to modify the document
            using (Transaction t = new Transaction(doc, "Update Cable Tray IDs"))
            {
                try
                {
                    t.Start();

                    // Filter for Cable Tray elements
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    List<Element> cableTraysToProcess = collector
                        .OfCategory(BuiltInCategory.OST_CableTray)
                        .WhereElementIsNotElementType() // Exclude element types (definitions)
                        .ToList(); // Convert to List to allow filtering and sorting

                    List<Element> filteredCableTrays = new List<Element>();
                    // Filter: Check if any of the RTS_Cable_XX parameters have a value
                    foreach (Element elem in cableTraysToProcess)
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
                            // Include all cable trays with cable values, regardless of orientation
                            filteredCableTrays.Add(elem);
                        }
                    }

                    // *** UPDATED: Implement the hierarchical sorting logic to prioritize lower-indexed cable parameters ***
                    IOrderedEnumerable<Element> sortedCableTrays = null;

                    if (filteredCableTrays.Any())
                    {
                        // Sort primarily by the index of the first non-empty RTS_Cable_XX parameter (lower index first),
                        // then by the value of that first non-empty parameter.
                        sortedCableTrays = filteredCableTrays
                            .OrderBy(e => GetFirstCableParamValue(e).Item1) // Sort by the index (e.g., 0 for RTS_Cable_01)
                            .ThenBy(e => GetFirstCableParamValue(e).Item2); // Then by the value of that parameter

                        // Apply subsequent sort criteria (ThenBy) for all RTS_Cable_GUIDs to ensure full consistency
                        // for the unique ID generation (currentCableKey).
                        for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
                        {
                            Guid currentGuid = RTS_Cable_GUIDs[i];
                            sortedCableTrays = sortedCableTrays
                                .ThenBy(e => e.get_Parameter(currentGuid)?.AsString() ?? string.Empty);
                        }
                    }
                    else
                    {
                        // If no filtered cable trays, or no GUIDs, proceed with an empty or default sort
                        sortedCableTrays = filteredCableTrays.OrderBy(e => e.Id.IntegerValue); // Fallback sort
                    }
                    // End of UPDATED sorting logic

                    int updatedCount = 0;
                    int uniqueIdCounter = 0; // Initialize counter
                    string previousCableKey = null; // Stores the concatenated values of RTS_Cable_XX for the previous element

                    foreach (Element elem in sortedCableTrays) // Iterate through the sorted list
                    {
                        CableTray cableTray = elem as CableTray;
                        if (cableTray == null) continue;

                        // Get RTS_ID parameter
                        Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
                        // New: Get Branch Number parameter
                        Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);


                        if (rtsIdParam != null)
                        {
                            // 1. Determine the prefix from the 5th character of the first RTS_Cable_XX with a value
                            string prefixChar = "X"; // Default prefix character
                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                                {
                                    string paramValue = cableParam.AsString();
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
                                    // Take the first available valid character and break
                                    break;
                                }
                            }

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

                            // 4. Get the bottom elevation in millimeters (rounded)
                            double bottomElevationFeet = 0.0; // This will be the final relative bottom elevation in feet
                            double absoluteMiddleElevationFeet = 0.0; // Absolute Z-coordinate of the cable tray's middle
                            double trayHeightFeet = 0.0;

                            // Primary attempt to get absolute middle elevation from LocationCurve.
                            // For a linear element like CableTray, the Z of the start point is a good reference.
                            if (elem.Location is LocationCurve locationCurve)
                            {
                                absoluteMiddleElevationFeet = locationCurve.Curve.GetEndPoint(0).Z;
                            }
                            else if (elem.Location is LocationPoint locationPoint)
                            {
                                // Fallback for point-based elements (though less common for cable trays)
                                absoluteMiddleElevationFeet = locationPoint.Point.Z;
                            }

                            // Get tray height from BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM
                            Parameter heightParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                            if (heightParam != null && heightParam.HasValue)
                            {
                                trayHeightFeet = heightParam.AsDouble();
                            }

                            // Calculate bottom elevation relative to the associated level
                            if (associatedLevel != null)
                            {
                                // Subtract the level's absolute elevation to get elevation relative to the level
                                double relativeMiddleElevation = absoluteMiddleElevationFeet - associatedLevel.Elevation;

                                // Calculate bottom elevation from middle elevation and half height
                                if (trayHeightFeet > 0)
                                {
                                    bottomElevationFeet = relativeMiddleElevation - (trayHeightFeet / 2.0);
                                }
                                else
                                {
                                    // If height is not available or zero, use middle elevation relative to level as fallback for "bottom"
                                    bottomElevationFeet = relativeMiddleElevation;
                                }
                            }
                            else // Fallback if no associated level could be determined
                            {
                                // In this case, we use the absolute middle elevation from LocationCurve as the "bottom"
                                // This value might still be "too high" if the level cannot be determined,
                                // but it's the best available approximation.
                                bottomElevationFeet = absoluteMiddleElevationFeet;
                                System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} could not determine associated level. Using absolute Z for elevation.");
                            }

                            // Convert feet to millimeters and round to no decimal places
                            long bottomElevationMM = (long)Math.Round(bottomElevationFeet * 304.8);

                            // Format bottomElevationMM to be a 4-character string, padded with leading zeros
                            string formattedBottomElevation = bottomElevationMM.ToString("D4");


                            // 5. Generate a unique 4-digit number based on the sorted order and parameter values
                            StringBuilder currentCableKeyBuilder = new StringBuilder();
                            foreach (Guid cableParamGuid in RTS_Cable_GUIDs)
                            {
                                Parameter cableParam = elem.get_Parameter(cableParamGuid);
                                currentCableKeyBuilder.Append(cableParam?.AsString() ?? string.Empty);
                                currentCableKeyBuilder.Append("|"); // Use a delimiter to prevent accidental matches (e.g., "A" + "B" vs "AB")
                            }
                            string currentCableKey = currentCableKeyBuilder.ToString();

                            // If this is the first element, or the current cable key is different from the previous one, increment the counter
                            if (uniqueIdCounter == 0 || currentCableKey != previousCableKey)
                            {
                                uniqueIdCounter++;
                            }
                            previousCableKey = currentCableKey; // Update the previous key for the next iteration

                            string uniqueSuffix = uniqueIdCounter.ToString("D4"); // Formats to 0001, 0002, etc.

                            // Construct the new RTS_ID with the new prefix and type suffix
                            string newRtsId = $"{prefixChar}-{trayTypeSuffix}-{levelAbbreviation}-{formattedBottomElevation}-{uniqueSuffix}";

                            // Set the new value to the RTS_ID parameter
                            rtsIdParam.Set(newRtsId);

                            // Set the uniqueSuffix to the Branch Number parameter
                            if (branchNumberParam != null)
                            {
                                branchNumberParam.Set(uniqueSuffix);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have the 'Branch Number' shared parameter.");
                            }
                            updatedCount++;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Warning: Cable Tray {elem.Id} does not have the 'RTS_ID' shared parameter.");
                        }
                    }

                    t.Commit();
                    TaskDialog.Show("RT_TrayID", $"{updatedCount} Cable Trays updated successfully (grouped by cable values).");
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
