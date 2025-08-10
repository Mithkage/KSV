//
// --- FILE: CableDataModels.cs ---
//
// Description:
// Contains data models related to cable data that are shared across multiple commands.
// Centralizing these models ensures consistency across different parts of the application.
//
// Change Log:
// - August 16, 2025: Initial creation. Extracted data models from PC_Extensible.cs and other files.
//

namespace RTS.Models
{
    /// <summary>
    /// Data model for cable information used across multiple commands
    /// </summary>
    public class CableData
    {
        public string FileName { get; set; }
        public string ImportDate { get; set; }
        public string CableReference { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string CableType { get; set; }
        public string CableCode { get; set; }
        public string CableConfiguration { get; set; }
        public string Cores { get; set; }
        public string ConductorActive { get; set; }
        public string Insulation { get; set; }
        public string ConductorEarth { get; set; }
        public string SeparateEarthForMulticore { get; set; }
        public string CableLength { get; set; }
        public string TotalCableRunWeight { get; set; }
        public string NominalOverallDiameter { get; set; }
        public string AsNsz3008CableDeratingFactor { get; set; }
        public string NumberOfActiveCables { get; set; }
        public string ActiveCableSize { get; set; }
        public string NumberOfNeutralCables { get; set; }
        public string NeutralCableSize { get; set; }
        public string NumberOfEarthCables { get; set; }
        public string EarthCableSize { get; set; }
        public string CablesKgPerM { get; set; }
        public string Sheath { get; set; }
        public string CableMaxLengthM { get; set; }
        public string VoltageVac { get; set; }
        public string LoadA { get; set; }
    }

    /// <summary>
    /// Data model for model-generated cable information
    /// </summary>
    public class ModelGeneratedData
    {
        public string CableReference { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string CableLengthM { get; set; }
        public string Variance { get; set; }
        public string Comment { get; set; }
    }

    /// <summary>
    /// Data model for cable tray occupancy calculations
    /// </summary>
    public class TrayOccupancyData
    {
        public string RtsId { get; set; }
        public string CableReference { get; set; }
        public string TrayOccupancy { get; set; }
        public string CablesWeight { get; set; }
        public string TrayMinSize { get; set; }
        public string Comment { get; set; }
    }
}