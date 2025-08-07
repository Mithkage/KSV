//-----------------------------------------------------------------------------
// <copyright file="RSGxCableData.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Model class for RSGx Cable Schedule data
// </summary>
//-----------------------------------------------------------------------------

namespace RTS.Reports.Models
{
    /// <summary>
    /// Data model for RSGx Cable Schedule reports
    /// </summary>
    public class RSGxCableData
    {
        public string RowNumber { get; set; } = "";
        public string CableTag { get; set; } = "";
        public string OriginDeviceID { get; set; } = "";
        public string DestinationDeviceID { get; set; } = "";
        public string RSGxRouteLengthM { get; set; } = "";
        public string DJVDesignLengthM { get; set; } = "";
        public string CableLengthDifferenceM { get; set; } = "";
        public string MaxLengthPermissibleForCableSizeM { get; set; } = "";
        public string ActiveCableSizeMM2 { get; set; } = "";
        public string NoOfSets { get; set; } = "";
        public string NeutralCableSizeMM2 { get; set; } = "";
        public string Cores { get; set; } = "";
        public string ConductorType { get; set; } = "";
        public string CableSizeChangeFromDesignYN { get; set; } = "";
        public string PreviousDesignSize { get; set; } = "";
        public string EarthIncludedYesNo { get; set; } = "";
        public string NumberOfEarthCables { get; set; } = "";
        public string EarthSizeMM2 { get; set; } = "";
        public string Voltage { get; set; } = "";
        public string VoltageRating { get; set; } = "";
        public string Type { get; set; } = "";
        public string SheathConstruction { get; set; } = "";
        public string InsulationConstruction { get; set; } = "";
        public string FireRating { get; set; } = "";
        public string LoadA { get; set; } = "";
        public string CableDescription { get; set; } = "";
        public string Comments { get; set; } = "";
        public string UpdateSummary { get; set; } = "";
    }
}