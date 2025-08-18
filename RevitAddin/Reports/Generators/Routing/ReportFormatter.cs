//-----------------------------------------------------------------------------
// <copyright file="ReportFormatter.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Handles formatting of the final report strings.
// </summary>
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// CHANGE LOG:
// 2024-08-13:
// - [APPLIED FIX]: Added FormatEquipment method to get a display name for equipment.
// - [APPLIED FIX]: Replaced non-standard OrderedSet with List and HashSet to resolve compilation error.
// - [APPLIED FIX]: Updated report header to include Total, Supported, and Unsupported length columns.
// - [APPLIED FIX]: Updated FormatRtsIdList to include containment type and length.
//
// Author: Kyle Vorster
//
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Handles formatting of the final report strings.
    /// </summary>
    public static class ReportFormatter
    {
        /// <summary>
        /// Creates the main header for the CSV report.
        /// </summary>
        public static StringBuilder CreateReportHeader(ProjectInfo projectInfo, Document doc)
        {
            string projectName = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NAME)?.AsString() ?? "N/A";
            string projectNumber = projectInfo?.get_Parameter(BuiltInParameter.PROJECT_NUMBER)?.AsString() ?? "N/A";
            var sb = new StringBuilder();
            sb.AppendLine("Cable Routing Sequence");
            sb.AppendLine($"Project:,{projectName}");
            sb.AppendLine($"Project No:,{projectNumber}");
            sb.AppendLine($"Date:,{DateTime.Now:dd/MM/yyyy}");
            sb.AppendLine();
            sb.AppendLine("Separator Legend:,,,,Tray/Conduit ID Legend:");
            sb.AppendLine(",, (Comma) = Direct connection between same containment types,,,,Format: [Prefix]-[TypeSuffix]-[Level]-[ElevMM]-[BranchNum]");
            sb.AppendLine(",, + (Plus) = Transition between different containment types,,,,Example: LA-FLS-L01-4285-0001");
            sb.AppendLine(",, || (Double Pipe) = Jump between non-connected containment of the same type,,,,Prefix: Service Type (LA=LV Sub-A, LB=LV Sub-B, GA=Gen-A)");
            sb.AppendLine(",, >> (Double Arrow) = Separates disconnected routing segments,,,,TypeSuffix: FLS=Fire-rated, ESS=Essential, DFT=Default");
            sb.AppendLine();
            sb.AppendLine("Cable Reference,From,To,From Status,To Status,Status,Total Length (m),Supported Length (m),Unsupported Length (m),Branch Sequencing,Routing Sequence,Assigned Containment,Graphed Containment,Island Count");
            return sb;
        }

        /// <summary>
        /// Builds a dictionary mapping RTS_ID strings to their corresponding elements.
        /// </summary>
        public static Dictionary<string, Element> BuildRtsIdMap(List<Element> elements, Guid rtsIdGuid)
        {
            var map = new Dictionary<string, Element>();
            foreach (var elem in elements)
            {
                string rtsId = elem.get_Parameter(rtsIdGuid)?.AsString();
                if (!string.IsNullOrWhiteSpace(rtsId) && !map.ContainsKey(rtsId))
                {
                    map.Add(rtsId, elem);
                }
            }
            return map;
        }

        /// <summary>
        /// Formats a list of element IDs into a human-readable, comma-separated string of RTS_IDs.
        /// </summary>
        public static string FormatPath(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap, Document doc, Guid rtsIdGuid)
        {
            if (path == null || !path.Any()) return "Path not found";

            var sequence = new List<string>();
            for (int i = 0; i < path.Count; i++)
            {
                var currentId = path[i];
                var currentElem = doc.GetElement(currentId);
                if (currentElem == null) continue;

                string rtsId = rtsIdToElementMap.FirstOrDefault(kvp => kvp.Value.Id == currentId).Key ??
                               currentElem.get_Parameter(rtsIdGuid)?.AsString() ??
                               "[MISSING_RTS_ID]";

                if (i > 0)
                {
                    var prevElem = doc.GetElement(path[i - 1]);
                    if (prevElem != null && IsContainmentTypeChange(currentElem, prevElem))
                    {
                        sequence.Add("+");
                    }
                    else
                    {
                        sequence.Add(",");
                    }
                }
                sequence.Add(rtsId);
            }
            return string.Join(" ", sequence).Replace(" , ", ", ").Replace(" + ", " + ");
        }

        /// <summary>
        /// Formats multiple disconnected paths into a single string separated by ">>".
        /// </summary>
        public static string FormatDisconnectedPaths(List<List<ElementId>> allPaths, Dictionary<string, Element> rtsIdMap, Document doc, Guid rtsIdGuid)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < allPaths.Count; i++)
            {
                if (i > 0) sb.Append(" >> ");
                sb.Append(FormatPath(allPaths[i], rtsIdMap, doc, rtsIdGuid));
            }
            return sb.ToString();
        }

        /// <summary>
        /// Extracts and formats the unique branch numbers from a path into a comma-separated string.
        /// </summary>
        public static string FormatBranchSequence(List<ElementId> path, Dictionary<string, Element> rtsIdToElementMap, Guid rtsIdGuid)
        {
            if (path == null || !path.Any()) return "N/A";

            var uniqueBranches = new List<string>();
            var seenBranches = new HashSet<string>();
            foreach (var elementId in path)
            {
                string rtsId = rtsIdToElementMap.FirstOrDefault(kvp => kvp.Value.Id == elementId).Key;
                if (!string.IsNullOrEmpty(rtsId) && rtsId.Length >= 4)
                {
                    string branchNumber = rtsId.Substring(rtsId.Length - 4);
                    if (seenBranches.Add(branchNumber)) // .Add() returns true if the item was new
                    {
                        uniqueBranches.Add(branchNumber);
                    }
                }
            }
            return string.Join(", ", uniqueBranches);
        }

        /// <summary>
        /// Formats a list of elements into a comma-separated string of their RTS_IDs, including type and length.
        /// </summary>
        public static string FormatRtsIdList(List<Element> elements, Guid rtsIdGuid, Dictionary<ElementId, double> lengthMap)
        {
            if (elements == null || !elements.Any()) return "N/A";

            var formattedParts = elements.Select(e =>
            {
                string rtsId = e.get_Parameter(rtsIdGuid)?.AsString() ?? "[MISSING_RTS_ID]";
                string suffix = GetContainmentSuffix(e);
                double lengthInFeet = lengthMap.ContainsKey(e.Id) ? lengthMap[e.Id] : 0.0;
                double lengthInMeters = lengthInFeet * 0.3048;
                int roundedLength = (int)Math.Ceiling(lengthInMeters);
                string lengthString = $"[{roundedLength}m]";

                return $"{rtsId}{suffix}{lengthString}";
            });

            return string.Join(", ", formattedParts);
        }

        /// <summary>
        /// Gets a display name for an equipment element, preferring RTS_ID over the panel name.
        /// </summary>
        public static string FormatEquipment(Element equipment, Document doc, Guid rtsIdGuid)
        {
            if (equipment == null) return "Unknown Equipment";
            return equipment.get_Parameter(rtsIdGuid)?.AsString()
                ?? equipment.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)?.AsString()
                ?? "Unnamed Equipment";
        }

        /// <summary>
        /// Checks if two elements represent a change in containment type (e.g., tray to conduit).
        /// </summary>
        private static bool IsContainmentTypeChange(Element elemA, Element elemB)
        {
            var catA = elemA.Category.Id.IntegerValue;
            var catB = elemB.Category.Id.IntegerValue;

            bool isTrayA = catA == (int)BuiltInCategory.OST_CableTray || catA == (int)BuiltInCategory.OST_CableTrayFitting;
            bool isConduitA = catA == (int)BuiltInCategory.OST_Conduit || catA == (int)BuiltInCategory.OST_ConduitFitting;

            bool isTrayB = catB == (int)BuiltInCategory.OST_CableTray || catB == (int)BuiltInCategory.OST_CableTrayFitting;
            bool isConduitB = catB == (int)BuiltInCategory.OST_Conduit || catB == (int)BuiltInCategory.OST_ConduitFitting;

            return (isTrayA && isConduitB) || (isConduitA && isTrayB);
        }

        /// <summary>
        /// Gets the containment type suffix for a given element.
        /// </summary>
        private static string GetContainmentSuffix(Element e)
        {
            if (e.Category == null) return "[?]";
            var catId = e.Category.Id.IntegerValue;

            if (catId == (int)BuiltInCategory.OST_CableTray) return "[T]";
            if (catId == (int)BuiltInCategory.OST_CableTrayFitting) return "[TF]";
            if (catId == (int)BuiltInCategory.OST_Conduit) return "[C]";
            if (catId == (int)BuiltInCategory.OST_ConduitFitting) return "[CF]";

            return "[?]";
        }
    }
}