//-----------------------------------------------------------------------------
// <copyright file="RSGxCableSummaryGenerator.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Generates the RSGx Cable Summary report in Excel format
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands;
using RTS.Reports.Base;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Generator for RSGx Cable Summary reports in Excel format
    /// </summary>
    public class RSGxCableSummaryGenerator : ReportGeneratorBase
    {
        public RSGxCableSummaryGenerator(Document doc, ExternalCommandData commandData, PC_Extensible.PC_ExtensibleClass pcExtensible) 
            : base(doc, commandData, pcExtensible)
        {
        }

        public override void GenerateReport()
        {
            string message = "";
            ElementSet elements = new ElementSet();

            PC_Generate_MDClass excelGenerator = new PC_Generate_MDClass();
            Result result = excelGenerator.Execute(CommandData, ref message, elements);

            if (result == Result.Failed)
            {
                ShowError($"Failed to generate RSGx Cable Summary (XLSX): {message}");
            }
        }
    }
}