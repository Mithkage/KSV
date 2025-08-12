// File: RT_TrayID.cs
// Application: Revit 2022/2024 External Command
// Description: This script identifies cable tray and conduit elements that have values
//              in any of the RTS_Cable_XX shared parameters. For each, it generates a unique ID
//              and assigns it to the 'RTS_ID' shared parameter. It then finds matching Detail Components,
//              Fittings, etc., by their RTS_ID and populates their cable parameters from the corresponding tray/conduit.
//
// RTS_ID Format: [Prefix]-[TypeSuffix]-[LevelAbbr]-[ElevationMM]-[UniqueSuffix]
// Example ID:    LA-FLS-L01-4285-0001
//   - Prefix (LA):         2-character prefix based on Type Name (e.g., "LV_SUB-A" -> "LA"). Defaults to "XX".
//   - TypeSuffix (FLS):  3-character abbreviation based on Type Name ("FR" -> "FLS", "ESS" -> "ESS", else "DFT").
//   - LevelAbbr (L01):   First 3 characters of the Reference Level name.
//   - ElevationMM (4285): Bottom elevation in millimeters, zero-padded to 4 digits.
//   - UniqueSuffix (0001): A sequential 4-digit number.

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using RTS.Utilities;
#endregion

namespace RTS.Commands.MEPTools.CableTray
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_TrayIDClass : IExternalCommand
    {
        // Use the shared parameter GUIDs from the utility class's compatibility layer
        private static readonly Guid RTS_ID_GUID = SharedParameters.General.RTS_ID;
        private static readonly Guid BRANCH_NUMBER_GUID = SharedParameters.Cable.RTS_Branch_Number;
        private static readonly List<Guid> RTS_Cable_GUIDs = SharedParameters.Cable.AllCableGuids;

        // Define the mapping from Type Name to RTS_ID prefix
        private static readonly Dictionary<string, string> TypeNameToPrefixMapping = new Dictionary<string, string>
        {
            { "LV_SUB-A", "LA" }, { "LV_SUB-B", "LB" }, { "_LV_SPARE", "LX" },
            { "LV_UPS-L", "UL" }, { "LV_UPS-P", "UP" }, { "_SHARED_LV", "LS" },
            { "LV_GEN-A", "GA" }, { "LV_GEN-B", "GB" }, { "LV_EARTHING", "LE" },
            { "_HV", "HV" }
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var idToCableValuesMap = new Dictionary<string, List<string>>();
            int updatedTrayCount = 0, updatedConduitCount = 0, updatedOtherCount = 0;

            using (var t = new Transaction(doc, "Generate and Apply RTS_IDs"))
            {
                try
                {
                    t.Start();

                    // --- 1. Process Cable Trays ---
                    var allCableTrays = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_CableTray)
                        .WhereElementIsNotElementType().ToList();
                    ProcessElements(doc, allCableTrays, idToCableValuesMap, ref updatedTrayCount, 0);

                    // --- 2. Process Conduits ---
                    var allConduits = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Conduit)
                        .WhereElementIsNotElementType().ToList();
                    ProcessElements(doc, allConduits, idToCableValuesMap, ref updatedConduitCount, updatedTrayCount);

                    // --- 3. Process Dependent Elements ---
                    var categoriesToUpdate = new List<BuiltInCategory>
                    {
                        BuiltInCategory.OST_DetailComponents,
                        BuiltInCategory.OST_CableTrayFitting,
                        BuiltInCategory.OST_ConduitFitting
                    };
                    var dependentElements = new FilteredElementCollector(doc)
                        .WherePasses(new ElementMulticategoryFilter(categoriesToUpdate))
                        .WhereElementIsNotElementType().ToList();

                    UpdateDependentElements(dependentElements, idToCableValuesMap, ref updatedOtherCount);

                    t.Commit();
                    TaskDialog.Show("RT_TrayID", $"{updatedTrayCount} Cable Trays, {updatedConduitCount} Conduits, and {updatedOtherCount} other elements updated successfully.");
                }
                catch (Exception ex)
                {
                    if (t.HasStarted()) t.RollBack();
                    message = $"Error: {ex.Message}";
                    TaskDialog.Show("RT_TrayID Error", message);
                    return Result.Failed;
                }
            }
            return Result.Succeeded;
        }

        private void ProcessElements(Document doc, List<Element> elements, Dictionary<string, List<string>> idMap, ref int updatedCount, int initialCounter)
        {
            var elementsWithValues = new List<Element>();
            var elementsWithoutValues = new List<Element>();

            foreach (var elem in elements)
            {
                if (RTS_Cable_GUIDs.Any(g => !string.IsNullOrEmpty(elem.get_Parameter(g)?.AsString())))
                    elementsWithValues.Add(elem);
                else
                    elementsWithoutValues.Add(elem);
            }

            var sortedElementsWithValues = elementsWithValues
                .OrderBy(e => GetFirstCableParamValue(e).Item1)
                .ThenBy(e => GetFirstCableParamValue(e).Item2);

            foreach (var guid in RTS_Cable_GUIDs)
            {
                sortedElementsWithValues = sortedElementsWithValues.ThenBy(e => e.get_Parameter(guid)?.AsString() ?? "");
            }

            int uniqueIdCounter = initialCounter;
            string previousCableKey = null;

            foreach (var elem in sortedElementsWithValues)
            {
                string currentCableKey = GetAllCableParamValues(elem);
                if (uniqueIdCounter == initialCounter || currentCableKey != previousCableKey)
                {
                    uniqueIdCounter++;
                }
                GenerateAndApplyId(doc, elem, uniqueIdCounter, idMap);
                previousCableKey = currentCableKey;
                updatedCount++;
            }

            int noValueCounter = (uniqueIdCounter == initialCounter ? 1000 : (uniqueIdCounter / 1000 + 1) * 1000);
            foreach (var elem in elementsWithoutValues)
            {
                GenerateAndApplyId(doc, elem, noValueCounter, idMap);
                noValueCounter++;
                updatedCount++;
            }
        }

        private void GenerateAndApplyId(Document doc, Element elem, int counter, Dictionary<string, List<string>> idMap)
        {
            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

            if (rtsIdParam == null || rtsIdParam.IsReadOnly) return;

            string prefix = GetPrefixFromTypeName(doc, elem);
            string suffix = "DFT";
            ElementType type = doc.GetElement(elem.GetTypeId()) as ElementType; // FIX: Added 'as ElementType' cast
            if (type != null)
            {
                string typeName = type.Name.ToUpper();
                if (typeName.Contains("FR") || typeName.Contains("FIRE")) suffix = "FLS";
                else if (typeName.Contains("ESS")) suffix = "ESS";
            }

            string levelAbbr = "???";
            ElementId levelId = elem.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)?.AsElementId() ?? elem.LevelId;
            if (levelId != ElementId.InvalidElementId && doc.GetElement(levelId) is Level level)
            {
                levelAbbr = level.Name.Length >= 3 ? level.Name.Substring(0, 3).ToUpper() : level.Name.ToUpper().PadRight(3, '?');
            }

            double offset = elem.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.AsDouble() ?? 0;
            double height = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? // For Cable Tray
                            elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)?.AsDouble() ?? 0; // For Conduit

            double bottomElevation = offset - (height / 2.0);
            long elevationMm = (long)Math.Round(bottomElevation * 304.8);
            string formattedElevation = elevationMm.ToString("D4");

            string uniqueSuffix = counter.ToString("D4");
            string newRtsId = $"{prefix}-{suffix}-{levelAbbr}-{formattedElevation}-{uniqueSuffix}";

            rtsIdParam.Set(newRtsId);
            branchNumberParam?.Set(uniqueSuffix);

            var cableValues = RTS_Cable_GUIDs
                .Select(g => elem.get_Parameter(g)?.AsString())
                .Where(v => !string.IsNullOrEmpty(v)).ToList();

            if (!idMap.ContainsKey(newRtsId))
            {
                idMap.Add(newRtsId, cableValues);
            }
        }

        private void UpdateDependentElements(List<Element> elements, Dictionary<string, List<string>> idMap, ref int updatedCount)
        {
            foreach (var elem in elements)
            {
                string rtsId = elem.get_Parameter(RTS_ID_GUID)?.AsString();
                if (string.IsNullOrEmpty(rtsId) || !idMap.TryGetValue(rtsId, out var cableValues)) continue;

                int cableIndex = 0;
                foreach (var guid in RTS_Cable_GUIDs)
                {
                    Parameter p = elem.get_Parameter(guid);
                    if (p != null && !p.IsReadOnly)
                    {
                        string newValue = (cableIndex < cableValues.Count) ? cableValues[cableIndex] : string.Empty;
                        if (p.AsString() != newValue)
                        {
                            p.Set(newValue);
                        }
                        cableIndex++;
                    }
                }
                updatedCount++;
            }
        }

        private Tuple<int, string> GetFirstCableParamValue(Element elem)
        {
            for (int i = 0; i < RTS_Cable_GUIDs.Count; i++)
            {
                Parameter cableParam = elem.get_Parameter(RTS_Cable_GUIDs[i]);
                if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))
                {
                    return new Tuple<int, string>(i, cableParam.AsString());
                }
            }
            return new Tuple<int, string>(RTS_Cable_GUIDs.Count, string.Empty);
        }

        private string GetAllCableParamValues(Element elem)
        {
            var sb = new StringBuilder();
            foreach (var guid in RTS_Cable_GUIDs)
            {
                sb.Append(elem.get_Parameter(guid)?.AsString() ?? "").Append("|");
            }
            return sb.ToString();
        }

        private string GetPrefixFromTypeName(Document doc, Element elem)
        {
            ElementType type = doc.GetElement(elem.GetTypeId()) as ElementType; // FIX: Added 'as ElementType' cast
            if (type != null)
            {
                string typeName = type.Name.ToUpper();
                foreach (var mapping in TypeNameToPrefixMapping)
                {
                    if (typeName.Contains(mapping.Key)) return mapping.Value;
                }
            }
            return "XX";
        }
    }
}
