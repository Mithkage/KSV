// File: RT_CableLengths.cs
// C# Revit 2022 and 2024 Add-in for finding cable lengths via cable tray and conduit elements

// ### PREPROCESSOR DIRECTIVES MUST BE AT THE VERY TOP ###
#if REVIT2024 || REVIT2023 || REVIT2022 // Use ForgeTypeId for 2022 and newer
#define USE_FORGE_TYPE_ID
#elif YOUR_SYMBOL_FOR_VERSIONS_OLDER_THAN_2022 // e.g., REVIT2021 - if you support them
    // Define symbols for versions older than 2022 if you need to differentiate further
#else
    // This error will trigger if a recognized Revit version symbol isn't defined
#error "Revit compilation symbol (e.g., REVIT2024, REVIT2023, REVIT2022) not defined in project build settings."
#endif

// Standard using statements will go below this block
#region Namespaces
// Standard system namespaces
using System;
using System.Collections.Generic;
using System.Linq;

// Autodesk Revit API namespaces
using Autodesk.Revit.ApplicationServices; // For Application and Document
using Autodesk.Revit.Attributes;          // For Transaction and RegenerationOption
using Autodesk.Revit.DB;                 // For Revit database elements and classes
using Autodesk.Revit.UI;                 // For UI elements like TaskDialog


// Conditionally use ForgeTypeId alias if USE_FORGE_TYPE_ID is defined
#if USE_FORGE_TYPE_ID
using ForgeTypeId = Autodesk.Revit.DB.ForgeTypeId;
#endif
#endregion

namespace RTS.Commands // Namespace to organize the code, reflecting the file name
{
    /// <summary>
    /// Revit external command to calculate and update cable lengths.
    /// </summary>
    [Transaction(TransactionMode.Manual)] // Indicates that the command will manage its own transactions
    [Regeneration(RegenerationOption.Manual)] // Indicates that regeneration will be handled manually
    public class RTCableLengthsCommand : IExternalCommand
    {
        // Define the GUID for the cable length shared parameter ("PC_Cable_Length").
        private readonly Guid _pcCableLengthGuid = new Guid("09007343-dd0a-44c3-b04d-145118778ac3");

        // Define the specific GUID for the "PC_SWB To" parameter on Detail Items.
        private readonly Guid _pcSwbToGuidForDetailItems = new Guid("e142b0ed-d084-447a-991b-d9a3a3f67a8d");

        // Define the list of GUIDs for the "Cable_XX" shared parameters on length elements.
        private readonly List<Guid> _cableXxGuidsForLengthElements = new List<Guid>
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
        /// Retrieves the value of the first matching "Cable_XX" parameter found on a length element.
        /// </summary>
        /// <param name="element">The Revit length element to check.</param>
        /// <param name="guidsToSearch">The list of GUIDs representing the parameters to search.</param>
        /// <returns>The string value of the first found parameter, or null if none are found or have a value.</returns>
        private string GetCableXxValueFromLengthElement(Element element, List<Guid> guidsToSearch)
        {
            foreach (Guid guid in guidsToSearch)
            {
                Parameter param = element.get_Parameter(guid);
                if (param?.HasValue ?? false)
                {
                    string value = param.AsString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        return value; // Return the first non-empty value found
                    }
                }
            }
            return null; // No matching or valid parameter found
        }

        /// <summary>
        /// The main entry point for the external command.
        /// </summary>
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Define categories for elements involved in length calculation
            List<BuiltInCategory> lengthCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_CableTray
            };
            // Define categories for items where lengths will be assigned
            // NOTE: Only Detail Items are currently handled based on the original code logic.
            // If Generic Annotations need similar handling, their category and logic
            // (including identifying their specific 'lookup' parameter) would need to be added.
            List<BuiltInCategory> targetCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_DetailComponents,
                // BuiltInCategory.OST_GenericAnnotation // Add this if needed
            };

            try
            {
                // Step 1: Collect Detail Items (and potentially Generic Annotations)
                ElementMulticategoryFilter targetFilter = new ElementMulticategoryFilter(targetCategories);
                FilteredElementCollector targetCollector = new FilteredElementCollector(doc)
                    .WherePasses(targetFilter)
                    .WhereElementIsNotElementType();
                IList<Element> targetItems = targetCollector.ToElements();

                if (!targetItems.Any())
                {
                    TaskDialog.Show("Info", "No Detail Items or Generic Annotations found in the project.");
                    return Result.Succeeded;
                }

                // Step 2: Clear existing values
                using (Transaction clearTransaction = new Transaction(doc, "Clear Cable Lengths"))
                {
                    clearTransaction.Start();
                    foreach (Element item in targetItems)
                    {
                        Parameter cableLengthParam = item.get_Parameter(_pcCableLengthGuid);
                        if (cableLengthParam != null && !cableLengthParam.IsReadOnly)
                        {
                            cableLengthParam.Set(""); // Clear by setting to empty string
                        }
                    }
                    clearTransaction.Commit();
                }

                // Step 3: Map Target Items (Detail Items/Generic Annotations) using relevant "To" parameter
                Dictionary<string, List<ElementId>> targetItemMapByLookupValue = new Dictionary<string, List<ElementId>>();
                foreach (Element item in targetItems)
                {
                    string lookupValue = null;
                    // Check if it's a Detail Item and get its specific "PC_SWB To" value
#if REVIT2024_OR_GREATER
                    if (item.Category.Id.Value == (int)BuiltInCategory.OST_DetailComponents)
#else
                    if (item.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DetailComponents)
#endif
                    {
                        Parameter pcSwbToParam = item.get_Parameter(_pcSwbToGuidForDetailItems);
                        if (pcSwbToParam?.HasValue ?? false)
                        {
                            lookupValue = pcSwbToParam.AsString();
                        }
                    }
                    // *** ADD LOGIC HERE FOR GENERIC ANNOTATIONS IF NEEDED ***
                    // #if REVIT2024_OR_GREATER
                    // else if (item.Category.Id.Value == (int)BuiltInCategory.OST_GenericAnnotation)
                    // #else
                    // else if (item.Category.Id.IntegerValue == (int)BuiltInCategory.OST_GenericAnnotation)
                    // #endif
                    // {
                    //      // Get the relevant lookup parameter for Generic Annotations
                    //      // Parameter genericLookupParam = item.get_Parameter(your_generic_annotation_guid);
                    //      // if (genericLookupParam?.HasValue ?? false)
                    //      // {
                    //      //       lookupValue = genericLookupParam.AsString();
                    //      // }
                    // }

                    if (!string.IsNullOrEmpty(lookupValue))
                    {
                        if (!targetItemMapByLookupValue.ContainsKey(lookupValue))
                        {
                            targetItemMapByLookupValue[lookupValue] = new List<ElementId>();
                        }
                        targetItemMapByLookupValue[lookupValue].Add(item.Id);
                    }
                }

                if (!targetItemMapByLookupValue.Any())
                {
                    TaskDialog.Show("Info", "No target items found with a value in their lookup parameter.");
                    return Result.Succeeded;
                }

                // Step 4: Sum lengths of Conduits/Cable Trays using *any* "Cable_XX" parameter
                Dictionary<string, double> summedLengthsByLookupValue = new Dictionary<string, double>();
                ElementMulticategoryFilter lengthCategoryFilter = new ElementMulticategoryFilter(lengthCategories);
                FilteredElementCollector lengthElementCollector = new FilteredElementCollector(doc)
                    .WherePasses(lengthCategoryFilter)
                    .WhereElementIsNotElementType();

                foreach (Element elem in lengthElementCollector)
                {
                    string lookupValueFromLengthElement = GetCableXxValueFromLengthElement(elem, _cableXxGuidsForLengthElements);

                    if (!string.IsNullOrEmpty(lookupValueFromLengthElement) && targetItemMapByLookupValue.ContainsKey(lookupValueFromLengthElement))
                    {
                        Parameter lengthParam = elem.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                        if (lengthParam?.HasValue ?? false)
                        {
                            double lengthInInternalUnits = lengthParam.AsDouble();
                            if (!summedLengthsByLookupValue.ContainsKey(lookupValueFromLengthElement))
                            {
                                summedLengthsByLookupValue[lookupValueFromLengthElement] = 0.0;
                            }
                            summedLengthsByLookupValue[lookupValueFromLengthElement] += lengthInInternalUnits;
                        }
                    }
                }

                // Step 5: Update "PC_Cable_Length" parameter on Target Items
                bool step5UpdateSuccessful = false;
                using (Transaction updateTransaction = new Transaction(doc, "Update Cable Lengths"))
                {
                    updateTransaction.Start();

#if USE_FORGE_TYPE_ID
                    ForgeTypeId internalUnits = UnitTypeId.Feet; // Revit's internal unit for length
                    ForgeTypeId targetUnits = UnitTypeId.Meters; // Desired output unit

                    foreach (var kvp in targetItemMapByLookupValue)
                    {
                        string lookupValueKey = kvp.Key;
                        List<ElementId> associatedItemIds = kvp.Value;

                        double totalLengthInInternalUnits = summedLengthsByLookupValue.ContainsKey(lookupValueKey) ? summedLengthsByLookupValue[lookupValueKey] : 0.0;
                        double totalLengthInMeters = UnitUtils.Convert(totalLengthInInternalUnits, internalUnits, targetUnits);

                        // ### START: NEW LOGIC ###
                        // Round up to the nearest integer
                        double processedLength = Math.Ceiling(totalLengthInMeters);

                        // If the rounded value is > 0, add 5
                        if (processedLength > 0)
                        {
                            processedLength += 5;
                        }

                        // Convert to an integer string (F0 format)
                        string lengthText = processedLength.ToString("F0");
                        // ### END: NEW LOGIC ###

                        foreach (ElementId itemId in associatedItemIds)
                        {
                            Element targetItem = doc.GetElement(itemId);
                            Parameter cableLengthParam = targetItem?.get_Parameter(_pcCableLengthGuid);
                            if (cableLengthParam != null && !cableLengthParam.IsReadOnly)
                            {
                                cableLengthParam.Set(lengthText); // Set the processed length
                            }
                        }
                    }
                    updateTransaction.Commit();
                    step5UpdateSuccessful = true;

#elif REVIT2022 // Note: Original code had DUT_FEET_FRACTIONAL_INCHES, ensure this is correct if < 2022 support is critical
                    DisplayUnitType internalUnitsDUT = DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES;
                    DisplayUnitType targetUnitsDUT = DisplayUnitType.DUT_METERS;

                    foreach (var kvp in targetItemMapByLookupValue)
                    {
                        string lookupValueKey = kvp.Key;
                        List<ElementId> associatedItemIds = kvp.Value;
                        double totalLengthInInternalUnits = summedLengthsByLookupValue.ContainsKey(lookupValueKey) ? summedLengthsByLookupValue[lookupValueKey] : 0.0;
                        double totalLengthInMeters = UnitUtils.Convert(totalLengthInInternalUnits, internalUnitsDUT, targetUnitsDUT);

                        // ### START: NEW LOGIC ###
                        // Round up to the nearest integer
                        double processedLength = Math.Ceiling(totalLengthInMeters);

                        // If the rounded value is > 0, add 5
                        if (processedLength > 0)
                        {
                            processedLength += 5;
                        }

                        // Convert to an integer string (F0 format)
                        string lengthText = processedLength.ToString("F0");
                        // ### END: NEW LOGIC ###


                        foreach (ElementId itemId in associatedItemIds)
                        {
                            Element targetItem = doc.GetElement(itemId);
                            Parameter cableLengthParam = targetItem?.get_Parameter(_pcCableLengthGuid);
                            if (cableLengthParam != null && !cableLengthParam.IsReadOnly)
                            {
                                cableLengthParam.Set(lengthText); // Set the processed length
                            }
                        }
                    }
                    updateTransaction.Commit();
                    step5UpdateSuccessful = true;
#else
                    message = "Unsupported Revit version or build configuration issue. Cannot perform unit conversion for updating lengths.";
                    TaskDialog.Show("Error", message);
                    updateTransaction.RollBack();
                    step5UpdateSuccessful = false;
#endif
                } // End using transaction

                // Step 6: Provide feedback to the user
                if (step5UpdateSuccessful)
                {
                    if (summedLengthsByLookupValue.Any(kvp => kvp.Value > 0))
                    {
                        TaskDialog.Show("Success", "Cable lengths updated successfully.");
                    }
                    else if (targetItemMapByLookupValue.Any())
                    {
                        TaskDialog.Show("Info", "No matching Conduits/Cable Trays found with lengths, or their 'Cable_XX' parameter was empty/did not match. 'PC_Cable_Length' parameters remain cleared or as they were.");
                    }
                    else
                    {
                        TaskDialog.Show("Info", "No operations performed. Check if target items and relevant Conduits/Cable Trays exist with matching lookup values.");
                    }
                    return Result.Succeeded;
                }
                else
                {
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "Failed to update cable lengths due to an unspecified issue during the update transaction.";
                    }
                    TaskDialog.Show("Failed", message);
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                message = "An unexpected error occurred: " + ex.Message + "\n\nStack Trace:\n" + ex.StackTrace;
                TaskDialog.Show("Critical Error", message);
                return Result.Failed;
            }
        }
    }
}
