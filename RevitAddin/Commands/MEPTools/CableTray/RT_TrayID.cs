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
            int renamedTrayCount = 0, unchangedTrayCount = 0;
            int renamedConduitCount = 0, unchangedConduitCount = 0;
            int renamedFittingCount = 0, unchangedFittingCount = 0;
            int renamedDetailCount = 0, unchangedDetailCount = 0;

            using (var t = new Transaction(doc, "Generate and Apply RTS_IDs"))
            {
                try
                {
                    t.Start();

                    // --- 1. Collect all processable elements ---
                    var allTrays = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToList();
                    var allConduits = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToList();
                    var allFittings = new FilteredElementCollector(doc).WherePasses(new ElementMulticategoryFilter(new[] { BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_ConduitFitting })).WhereElementIsNotElementType().ToList();

                    var allProcessableElements = new List<Element>();
                    allProcessableElements.AddRange(allTrays);
                    allProcessableElements.AddRange(allConduits);
                    allProcessableElements.AddRange(allFittings);

                    // --- 2. Separate elements into categories for processing ---
                    var elementsWithValues = new List<Element>();
                    var fittingsWithoutValues = new List<Element>();
                    var traysAndConduitsWithoutValues = new List<Element>();

                    foreach (var elem in allProcessableElements)
                    {
                        if (RTS_Cable_GUIDs.Any(g => !string.IsNullOrEmpty(elem.get_Parameter(g)?.AsString())))
                        {
                            elementsWithValues.Add(elem);
                        }
                        else
                        {
                            if (elem is FamilyInstance) fittingsWithoutValues.Add(elem);
                            else traysAndConduitsWithoutValues.Add(elem);
                        }
                    }

                    // --- 3. Globally determine branch numbers based on unique cable sets ---
                    var cableKeyToBranchMap = new Dictionary<string, int>();
                    int branchCounter = 0;
                    var groupedByCableKey = elementsWithValues.GroupBy(e => GetAllCableParamValues(e)).OrderBy(g => g.Key);

                    foreach (var group in groupedByCableKey)
                    {
                        branchCounter++;
                        cableKeyToBranchMap[group.Key] = branchCounter;
                    }

                    // --- 4. Apply IDs to elements with cable data ---
                    foreach (var elem in elementsWithValues)
                    {
                        string cableKey = GetAllCableParamValues(elem);
                        int branchNumber = cableKeyToBranchMap[cableKey];
                        bool changed = GenerateAndApplyId(doc, elem, branchNumber, idToCableValuesMap);

                        if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray) { if (changed) renamedTrayCount++; else unchangedTrayCount++; }
                        else if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit) { if (changed) renamedConduitCount++; else unchangedConduitCount++; }
                        else { if (changed) renamedFittingCount++; else unchangedFittingCount++; }
                    }

                    // --- 5. Handle trays and conduits without cable data ---
                    int noValueCounter = (branchCounter == 0 ? 1000 : (branchCounter / 1000 + 1) * 1000);
                    foreach (var elem in traysAndConduitsWithoutValues)
                    {
                        bool changed = GenerateAndApplyId(doc, elem, noValueCounter, idToCableValuesMap);
                        if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray) { if (changed) renamedTrayCount++; else unchangedTrayCount++; }
                        else { if (changed) renamedConduitCount++; else unchangedConduitCount++; }
                        noValueCounter++;
                    }

                    // --- 6. Handle fittings without cable data ---
                    foreach (var elem in fittingsWithoutValues)
                    {
                        if (ClearIdParameters(elem)) renamedFittingCount++;
                        else unchangedFittingCount++;
                    }

                    // --- 7. Process Dependent Detail Components ---
                    var detailComponents = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToList();
                    UpdateDependentElements(detailComponents, idToCableValuesMap, ref renamedDetailCount, ref unchangedDetailCount);

                    t.Commit();

                    var summary = new StringBuilder();
                    summary.AppendLine("RTS_ID Generation Complete.");
                    summary.AppendLine();
                    summary.AppendLine($"Cable Trays: {renamedTrayCount} renamed, {unchangedTrayCount} unchanged.");
                    summary.AppendLine($"Conduits: {renamedConduitCount} renamed, {unchangedConduitCount} unchanged.");
                    summary.AppendLine($"Fittings: {renamedFittingCount} updated or cleared, {unchangedFittingCount} unchanged.");
                    summary.AppendLine($"Detail Components: {renamedDetailCount} updated, {unchangedDetailCount} unchanged.");

                    TaskDialog.Show("RT_TrayID", summary.ToString());
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

        private bool GenerateAndApplyId(Document doc, Element elem, int counter, Dictionary<string, List<string>> idMap)
        {
            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);

            if (rtsIdParam == null || rtsIdParam.IsReadOnly) return false;

            string oldRtsId = rtsIdParam.AsString();
            string oldBranchNumber = branchNumberParam?.AsString();

            string prefix = GetPrefixFromTypeName(doc, elem);
            string suffix = "DFT";
            ElementType type = doc.GetElement(elem.GetTypeId()) as ElementType;
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

            long elevationMm = GetElevationInMm(elem, doc);
            string formattedElevation = elevationMm.ToString("D4");

            string uniqueSuffix = counter.ToString("D4");
            string newRtsId = $"{prefix}-{suffix}-{levelAbbr}-{formattedElevation}-{uniqueSuffix}";

            bool idChanged = oldRtsId != newRtsId;
            bool branchChanged = branchNumberParam != null && oldBranchNumber != uniqueSuffix;

            if (idChanged) rtsIdParam.Set(newRtsId);
            if (branchChanged) branchNumberParam.Set(uniqueSuffix);

            var cableValues = RTS_Cable_GUIDs
                .Select(g => elem.get_Parameter(g)?.AsString())
                .Where(v => !string.IsNullOrEmpty(v)).ToList();

            if (idChanged && !string.IsNullOrEmpty(newRtsId) && !idMap.ContainsKey(newRtsId))
            {
                idMap.Add(newRtsId, cableValues);
            }

            return idChanged || branchChanged;
        }

        private bool ClearIdParameters(Element elem)
        {
            bool changed = false;
            Parameter rtsIdParam = elem.get_Parameter(RTS_ID_GUID);
            if (rtsIdParam != null && !rtsIdParam.IsReadOnly && !string.IsNullOrEmpty(rtsIdParam.AsString()))
            {
                rtsIdParam.Set(string.Empty);
                changed = true;
            }

            Parameter branchNumberParam = elem.get_Parameter(BRANCH_NUMBER_GUID);
            if (branchNumberParam != null && !branchNumberParam.IsReadOnly && !string.IsNullOrEmpty(branchNumberParam.AsString()))
            {
                branchNumberParam.Set(string.Empty);
                changed = true;
            }
            return changed;
        }

        private long GetElevationInMm(Element elem, Document doc)
        {
            // Handle MEPCurves (Trays/Conduits)
            if (elem is MEPCurve curve)
            {
                double offset = curve.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)?.AsDouble() ?? 0;
                double height = curve.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? // For Cable Tray
                                curve.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)?.AsDouble() ?? 0; // For Conduit
                double bottomElevation = offset - (height / 2.0);
                return (long)Math.Round(bottomElevation * 304.8);
            }

            // Handle FamilyInstances (Fittings)
            if (elem is FamilyInstance fitting)
            {
                var mepModel = fitting.MEPModel;
                if (mepModel?.ConnectorManager?.Connectors != null && mepModel.ConnectorManager.Connectors.Size > 0)
                {
                    // Method 3: Try to inherit from a connected element
                    foreach (Connector connector in mepModel.ConnectorManager.Connectors)
                    {
                        if (connector.IsConnected)
                        {
                            foreach (Connector connectedRef in connector.AllRefs)
                            {
                                Element connectedElem = connectedRef.Owner;
                                if (connectedElem is MEPCurve)
                                {
                                    Parameter connectedRtsIdParam = connectedElem.get_Parameter(RTS_ID_GUID);
                                    if (connectedRtsIdParam != null && connectedRtsIdParam.HasValue)
                                    {
                                        string connectedRtsId = connectedRtsIdParam.AsString();
                                        string[] parts = connectedRtsId.Split('-');
                                        if (parts.Length == 5 && long.TryParse(parts[3], out long inheritedElevation))
                                        {
                                            return inheritedElevation; // Success!
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Method 2 (Fallback): Average elevation of connectors
                    double totalZ = 0;
                    int connectorCount = 0;
                    foreach (Connector connector in mepModel.ConnectorManager.Connectors)
                    {
                        totalZ += connector.Origin.Z;
                        connectorCount++;
                    }
                    if (connectorCount > 0)
                    {
                        double avgElevationInFeet = totalZ / connectorCount;
                        return (long)Math.Round(avgElevationInFeet * 304.8);
                    }
                }
            }

            // Default fallback for any other element type or if fitting has no connectors
            return 0;
        }

        private void UpdateDependentElements(List<Element> elements, Dictionary<string, List<string>> idMap, ref int renamedCount, ref int unchangedCount)
        {
            foreach (var elem in elements)
            {
                bool elementChanged = false;
                string rtsId = elem.get_Parameter(RTS_ID_GUID)?.AsString();
                if (string.IsNullOrEmpty(rtsId) || !idMap.TryGetValue(rtsId, out var cableValues))
                {
                    unchangedCount++;
                    continue;
                }

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
                            elementChanged = true;
                        }
                        cableIndex++;
                    }
                }

                if (elementChanged)
                    renamedCount++;
                else
                    unchangedCount++;
            }
        }

        private string GetAllCableParamValues(Element elem)
        {
            var cableValues = RTS_Cable_GUIDs
                .Select(g => elem.get_Parameter(g)?.AsString())
                .Where(v => !string.IsNullOrEmpty(v))
                .OrderBy(v => v) // Sort alphabetically to create a canonical key
                .ToList();

            return string.Join("|", cableValues);
        }

        private string GetPrefixFromTypeName(Document doc, Element elem)
        {
            ElementType type = doc.GetElement(elem.GetTypeId()) as ElementType;
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