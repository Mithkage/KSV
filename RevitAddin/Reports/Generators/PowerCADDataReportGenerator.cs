//-----------------------------------------------------------------------------
// <copyright file="PowerCADDataReportGenerator.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Generates PowerCAD data reports from extensible storage
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.DataExchange.DataManagement;
using RTS.Reports.Base;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators
{
    /// <summary>
    /// Generator for PowerCAD data reports
    /// </summary>
    public class PowerCADDataReportGenerator : ReportGeneratorBase
    {
        private readonly Guid _schemaGuid;
        private readonly string _schemaName;
        private readonly string _fieldName;
        private readonly string _dataStorageElementName;
        private readonly string _defaultFileName;
        private readonly string _dataTypeName;

        public PowerCADDataReportGenerator(
            Document doc,
            ExternalCommandData commandData,
            PC_ExtensibleClass pcExtensible,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName,
            string defaultFileName,
            string dataTypeName) : base(doc, commandData, pcExtensible)
        {
            _schemaGuid = schemaGuid;
            _schemaName = schemaName;
            _fieldName = fieldName;
            _dataStorageElementName = dataStorageElementName;
            _defaultFileName = defaultFileName;
            _dataTypeName = dataTypeName;
        }

        public override void GenerateReport()
        {
            List<object> dataToExport = PcExtensible.RecallDataFromExtensibleStorage<object>(
                Document, _schemaGuid, _schemaName, _fieldName, _dataStorageElementName);

            if (dataToExport == null || !dataToExport.Any())
            {
                ShowInfo("No Data", $"No {_dataTypeName} found to export.");
                return;
            }

            string filePath = GetOutputFilePath(_defaultFileName, $"Save {_dataTypeName} Report");
            if (string.IsNullOrEmpty(filePath))
            {
                ShowInfo("Export Cancelled", "Output file not selected.");
                return;
            }

            try
            {
                ExportDataToCsvGeneric(dataToExport, filePath, _dataTypeName);
                ShowSuccess("Export Complete", $"{_dataTypeName} successfully exported to:\n{filePath}");
            }
            catch (Exception ex)
            {
                ShowError($"Failed to export {_dataTypeName}: {ex.Message}");
            }
        }
    }
}