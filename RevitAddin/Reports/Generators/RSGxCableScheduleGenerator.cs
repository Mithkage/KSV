//-----------------------------------------------------------------------------
// <copyright file="RSGxCableScheduleGenerator.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Generates the RSGx Cable Schedule report
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Reports.Base;
using RTS.Reports.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Generator for RSGx Cable Schedule reports
    /// </summary>
    public class RSGxCableScheduleGenerator : ReportGeneratorBase
    {
        public RSGxCableScheduleGenerator(
            Document doc,
            ExternalCommandData commandData,
            PC_ExtensibleClass pcExtensible) : base(doc, commandData, pcExtensible)
        {
        }

        public override void GenerateReport()
        {
            string filePath = GetOutputFilePath("RSGx Cable Schedule.csv", "Save RSGx Cable Schedule Report");
            if (string.IsNullOrEmpty(filePath))
            {
                ShowInfo("Export Cancelled", "Output file not selected. RSGx Cable Schedule export cancelled.");
                return;
            }

            try
            {
                List<PC_ExtensibleClass.CableData> primaryData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    Document, PC_ExtensibleClass.PrimarySchemaGuid, PC_ExtensibleClass.PrimarySchemaName,
                    PC_ExtensibleClass.PrimaryFieldName, PC_ExtensibleClass.PrimaryDataStorageElementName
                );

                List<PC_ExtensibleClass.CableData> consultantData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.CableData>(
                    Document, PC_ExtensibleClass.ConsultantSchemaGuid, PC_ExtensibleClass.ConsultantSchemaName,
                    PC_ExtensibleClass.ConsultantFieldName, PC_ExtensibleClass.ConsultantDataStorageElementName
                );

                List<PC_ExtensibleClass.ModelGeneratedData> modelGeneratedData = PcExtensible.RecallDataFromExtensibleStorage<PC_ExtensibleClass.ModelGeneratedData>(
                    Document, PC_ExtensibleClass.ModelGeneratedSchemaGuid, PC_ExtensibleClass.ModelGeneratedSchemaName,
                    PC_ExtensibleClass.ModelGeneratedFieldName, PC_ExtensibleClass.ModelGeneratedDataStorageElementName
                );

                var primaryDict = primaryData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var consultantDict = consultantData.ToDictionary(c => c.CableReference, c => c, StringComparer.OrdinalIgnoreCase);
                var modelGeneratedDict = modelGeneratedData.ToDictionary(m => m.CableReference, m => m, StringComparer.OrdinalIgnoreCase);

                var finalReportCableRefs = new List<string>();
                var processedCableRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var primaryCable in primaryData)
                {
                    if (!string.IsNullOrEmpty(primaryCable.CableReference) && processedCableRefs.Add(primaryCable.CableReference))
                    {
                        finalReportCableRefs.Add(primaryCable.CableReference);
                    }
                }

                var consultantOnlyRefs = consultantData
                    .Where(c => !string.IsNullOrEmpty(c.CableReference) && !processedCableRefs.Contains(c.CableReference))
                    .Select(c => c.CableReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(cr => cr, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var conRef in consultantOnlyRefs)
                {
                    if (processedCableRefs.Add(conRef))
                    {
                        finalReportCableRefs.Add(conRef);
                    }
                }

                var modelGeneratedOnlyRefs = modelGeneratedData
                    .Where(m => !string.IsNullOrEmpty(m.CableReference) && !processedCableRefs.Contains(m.CableReference))
                    .Select(m => m.CableReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(cr => cr, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var modelRef in modelGeneratedOnlyRefs)
                {
                    if (processedCableRefs.Add(modelRef))
                    {
                        finalReportCableRefs.Add(modelRef);
                    }
                }

                var reportData = new List<RSGxCableData>();
                int rowNum = 1;

                foreach (string cableRef in finalReportCableRefs)
                {
                    primaryDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData primaryInfo);
                    consultantDict.TryGetValue(cableRef, out PC_ExtensibleClass.CableData consultantInfo);
                    modelGeneratedDict.TryGetValue(cableRef, out PC_ExtensibleClass.ModelGeneratedData modelInfo);

                    string originDeviceID = primaryInfo?.From ?? consultantInfo?.From ?? "N/A";
                    string destinationDeviceID = primaryInfo?.To ?? consultantInfo?.To ?? "N/A";
                    string rsgxRouteLength = primaryInfo?.CableLength ?? "N/A";
                    string djvDesignLength = consultantInfo?.CableLength ?? "N/A";

                    string cableLengthDifference = "N/A";
                    if (double.TryParse(rsgxRouteLength, out double rsgxLen) && double.TryParse(djvDesignLength, out double djvLen))
                    {
                        cableLengthDifference = (rsgxLen - djvLen).ToString("F1");
                    }

                    string cableSizeChangeYN;
                    string primaryActiveSize = primaryInfo?.ActiveCableSize;
                    string consultantActiveSize = consultantInfo?.ActiveCableSize;

                    if (string.IsNullOrEmpty(primaryActiveSize) || primaryActiveSize == "N/A" ||
                        string.IsNullOrEmpty(consultantActiveSize) || consultantActiveSize == "N/A")
                    {
                        cableSizeChangeYN = "N/A";
                    }
                    else
                    {
                        cableSizeChangeYN = (string.Equals(primaryActiveSize, consultantActiveSize, StringComparison.OrdinalIgnoreCase)) ? "N" : "Y";
                    }

                    string earthIncludedRaw = primaryInfo?.SeparateEarthForMulticore;
                    string earthIncludedFormatted = string.Equals(earthIncludedRaw, "No", StringComparison.OrdinalIgnoreCase) ? "Y" : "N";

                    string fireRatingValue = "";
                    if (primaryInfo?.CableType != null && primaryInfo.CableType.IndexOf("Fire", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        fireRatingValue = "WS52W";
                    }

                    var summaryParts = new List<string>();
                    if (double.TryParse(cableLengthDifference, out double diff))
                    {
                        if (diff > 0) summaryParts.Add("Length increased.");
                        else if (diff < 0) summaryParts.Add("Length decreased.");
                        else summaryParts.Add("Length consistent.");
                    }

                    string primaryLoadStr = primaryInfo?.LoadA;
                    string consultantLoadStr = consultantInfo?.LoadA;
                    if (double.TryParse(primaryLoadStr, out double primaryLoad) && double.TryParse(consultantLoadStr, out double consultantLoad))
                    {
                        if (primaryLoad < consultantLoad) summaryParts.Add("MD decreased");
                        else if (primaryLoad > consultantLoad) summaryParts.Add("MD increased");
                        else summaryParts.Add("MD un-changed");
                    }
                    else
                    {
                        summaryParts.Add("N/A");
                    }
                    string updateSummaryValue = string.Join(" | ", summaryParts.Where(s => !string.IsNullOrEmpty(s)));

                    string maxLengthPermissible = primaryInfo?.CableMaxLengthM ?? "N/A";
                    if (double.TryParse(maxLengthPermissible, out double maxLen))
                    {
                        maxLengthPermissible = Math.Floor(maxLen).ToString();
                    }
                    else
                    {
                        maxLengthPermissible = "N/A";
                    }

                    string voltageRating;
                    string cableDescriptionValue;
                    string cores = primaryInfo?.Cores ?? "N/A";

                    string activeCableSize = primaryInfo?.ActiveCableSize ?? "N/A";
                    string type = primaryInfo?.CableType ?? "N/A";
                    string sheath = primaryInfo?.Sheath ?? "N/A";
                    string insulation = primaryInfo?.Insulation ?? "N/A";

                    if (cores == "N/A")
                    {
                        voltageRating = "N/A";
                        cableDescriptionValue = "N/A";
                    }
                    else
                    {
                        voltageRating = "0.6/1kV";
                        if (destinationDeviceID == "N/A")
                        {
                            voltageRating = "";
                        }

                        var descriptionParts = new List<string>();

                        if (!string.IsNullOrEmpty(activeCableSize) && activeCableSize != "N/A") descriptionParts.Add($"{activeCableSize}mm²");
                        if (!string.IsNullOrEmpty(cores) && cores != "N/A") descriptionParts.Add(cores);
                        if (!string.IsNullOrEmpty(voltageRating) && voltageRating != "N/A") descriptionParts.Add(voltageRating);
                        if (!string.IsNullOrEmpty(type) && type != "N/A") descriptionParts.Add(type);
                        if (!string.IsNullOrEmpty(sheath) && sheath != "N/A") descriptionParts.Add(sheath);
                        if (!string.IsNullOrEmpty(insulation) && insulation != "N/A") descriptionParts.Add(insulation);
                        if (!string.IsNullOrEmpty(fireRatingValue) && fireRatingValue != "N/A") descriptionParts.Add(fireRatingValue);
                        cableDescriptionValue = string.Join(" | ", descriptionParts);
                    }

                    if (!string.IsNullOrEmpty(type) && !string.Equals(type, "N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        type += ", LSZH";
                    }

                    reportData.Add(new RSGxCableData
                    {
                        RowNumber = (rowNum++).ToString(),
                        CableTag = cableRef,
                        OriginDeviceID = originDeviceID,
                        DestinationDeviceID = destinationDeviceID,
                        RSGxRouteLengthM = rsgxRouteLength,
                        DJVDesignLengthM = djvDesignLength,
                        CableLengthDifferenceM = cableLengthDifference,
                        MaxLengthPermissibleForCableSizeM = maxLengthPermissible,
                        ActiveCableSizeMM2 = activeCableSize,
                        NoOfSets = primaryInfo?.NumberOfActiveCables ?? "N/A",
                        NeutralCableSizeMM2 = primaryInfo?.NeutralCableSize ?? "N/A",
                        Cores = cores,
                        ConductorType = primaryInfo?.ConductorActive ?? "N/A",
                        CableSizeChangeFromDesignYN = cableSizeChangeYN,
                        PreviousDesignSize = consultantInfo?.ActiveCableSize ?? "N/A",
                        EarthIncludedYesNo = earthIncludedFormatted,
                        NumberOfEarthCables = primaryInfo?.NumberOfEarthCables ?? "N/A",
                        EarthSizeMM2 = primaryInfo?.EarthCableSize ?? "N/A",
                        Voltage = primaryInfo?.VoltageVac ?? "N/A",
                        VoltageRating = voltageRating,
                        Type = type,
                        SheathConstruction = sheath,
                        InsulationConstruction = insulation,
                        FireRating = fireRatingValue,
                        LoadA = primaryInfo?.LoadA ?? "N/A",
                        CableDescription = cableDescriptionValue,
                        Comments = "",
                        UpdateSummary = updateSummaryValue
                    });
                }

                if (!reportData.Any())
                {
                    ShowInfo("No Data", "No unique cable references found to generate the RSGx Cable Schedule.");
                    return;
                }

                ExportDataToCsvGeneric(reportData, filePath, "RSGx Cable Schedule");
                ShowSuccess("Export Complete", $"RSGx Cable Schedule successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                ShowError($"Failed to export RSGx Cable Schedule: {ex.Message}");
            }
        }
    }
}