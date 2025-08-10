//-----------------------------------------------------------------------------
// <copyright file="ReportGeneratorBase.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//     Base class for report generation, containing common functionality.
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RTS.Commands.DataExchange.DataManagement;
using System;
using System.Collections.Generic;
using System.Linq; // Added System.Linq for LINQ extension methods
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RTS.Reports.Base
{
    /// <summary>
    /// Base class with common report generation functionality
    /// </summary>
    public abstract class ReportGeneratorBase
    {
        protected Document Document { get; private set; }
        protected ExternalCommandData CommandData { get; private set; }
        protected PC_ExtensibleClass PcExtensible { get; private set; }

        public ReportGeneratorBase(Document doc, ExternalCommandData commandData, PC_ExtensibleClass pcExtensible)
        {
            Document = doc;
            CommandData = commandData;
            PcExtensible = pcExtensible;
        }

        /// <summary>
        /// Shows an error message dialog
        /// </summary>
        protected void ShowError(string message)
        {
            TaskDialog.Show("Error", message);
        }

        /// <summary>
        /// Shows a success message dialog
        /// </summary>
        protected void ShowSuccess(string title, string message)
        {
            TaskDialog.Show(title, message);
        }

        /// <summary>
        /// Shows an information message dialog
        /// </summary>
        protected void ShowInfo(string title, string message)
        {
            TaskDialog.Show(title, message);
        }

        /// <summary>
        /// Gets a file path for saving the report
        /// </summary>
        protected string GetOutputFilePath(string defaultFileName, string dialogTitle)
        {
            return RTS.Utilities.FileDialogs.PromptForSaveFile(
                defaultFileName,
                dialogTitle,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                null,
                "csv");
        }

        /// <summary>
        /// Exports data to a CSV file
        /// </summary>
        protected void ExportDataToCsvGeneric<T>(List<T> data, string filePath, string dataTypeName) where T : class
        {
            var sb = new StringBuilder();
            PropertyInfo[] allTypeProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            List<string> headers = new List<string>();
            List<PropertyInfo> orderedPropertiesToExport = new List<PropertyInfo>();

            if (typeof(T) == typeof(Models.RSGxCableData))
            {
                headers.AddRange(new string[] {
                    "Row Number", "Cable Tag", "Origin Device (ID)", "Destination Device (ID)",
                    "RSGx Route Length (m)", "DJV Design Length (m)", "Cable Length Difference (m)",
                    "Maximum Length permissible for Cable Size (m)",
                    "RSGx Active Cable Size (mm²)",
                    "DJV Active Cable Size (mm²)",
                    "Cable size Change from Design (Y/N)",
                    "No. of Sets",
                    "Neutral Cable Size (mm²)", "Cores", "Conductor Type",
                    "Earth Included (Y/N)",
                    "Number of Earth Cables", "Earth Size (mm2)", "Voltage", "VoltageRating", "Type", "Sheath Construction",
                    "Insulation Construction", "Fire Rating", "Load (A)",
                    "Cable Description", "Comments", "Update Summary"
                });

                var rsgxPropertyNamesInOrder = new List<string> {
                    "RowNumber", "CableTag", "OriginDeviceID", "DestinationDeviceID",
                    "RSGxRouteLengthM", "DJVDesignLengthM", "CableLengthDifferenceM",
                    "MaxLengthPermissibleForCableSizeM",
                    "ActiveCableSizeMM2",
                    "PreviousDesignSize",
                    "CableSizeChangeFromDesignYN",
                    "NoOfSets",
                    "NeutralCableSizeMM2", "Cores", "ConductorType",
                    "EarthIncludedYesNo",
                    "NumberOfEarthCables", "EarthSizeMM2", "Voltage", "VoltageRating", "Type", "SheathConstruction",
                    "InsulationConstruction", "FireRating", "LoadA",
                    "CableDescription", "Comments", "UpdateSummary"
                };

                foreach (var propName in rsgxPropertyNamesInOrder)
                {
                    var propInfo = typeof(T).GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                    if (propInfo != null)
                    {
                        orderedPropertiesToExport.Add(propInfo);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Property '{propName}' not found in RSGxCableData class.");
                    }
                }
            }
            else
            {
                headers = allTypeProperties.Select(p => Regex.Replace(p.Name, "([a-z])([A-Z])", "$1 $2")).ToList();
                orderedPropertiesToExport = allTypeProperties.ToList();
            }

            sb.AppendLine(string.Join(",", headers));

            foreach (var item in data)
            {
                string[] values = orderedPropertiesToExport.Select(p =>
                {
                    object value = p.GetValue(item);
                    string stringValue = value?.ToString() ?? string.Empty;
                    if (stringValue.Contains(",") || stringValue.Contains("\"") || stringValue.Contains("\n") || stringValue.Contains("\r"))
                    {
                        return $"\"{stringValue.Replace("\"", "\"\"")}\"";
                    }
                    return stringValue.Trim();
                }).ToArray();
                sb.AppendLine(string.Join(",", values));
            }

            System.IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Main method to generate the report
        /// </summary>
        public abstract void GenerateReport();
    }
}