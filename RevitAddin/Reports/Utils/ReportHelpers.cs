//-----------------------------------------------------------------------------
// <copyright file="ReportHelpers.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Common utility methods for report generation
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RTS.Reports.Utils
{
    /// <summary>
    /// Helper methods for report generation
    /// </summary>
    public static class ReportHelpers
    {
        /// <summary>
        /// Gets the location of a Revit element
        /// </summary>
        public static XYZ GetElementLocation(Element element)
        {
            if (element == null) return null;
            try
            {
                if (element.Location is LocationPoint locPoint) return locPoint.Point;
                if (element.Location is LocationCurve locCurve) return locCurve.Curve.GetEndPoint(0);
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Cleans up a cable reference from various formats
        /// </summary>
        public static string CleanCableReference(string cableReference)
        {
            if (string.IsNullOrEmpty(cableReference)) return cableReference;
            string cleaned = cableReference.Trim();
            int openParenIndex = cleaned.IndexOf('(');
            if (openParenIndex != -1)
            {
                cleaned = cleaned.Substring(0, openParenIndex).Trim();
            }
            int firstSlashIndex = cleaned.IndexOf('/');
            if (firstSlashIndex != -1)
            {
                cleaned = cleaned.Substring(0, firstSlashIndex).Trim();
            }
            string[] parts = cleaned.Split('-');
            Regex prefixPattern = new Regex(@"^[A-Za-z]{2}\d{2}", RegexOptions.IgnoreCase);
            if (parts.Length >= 3 && prefixPattern.IsMatch(parts[0]))
            {
                if (!int.TryParse(parts[2], out _)) return $"{parts[0]}-{parts[1]}";
            }
            if (parts.Length >= 4 && prefixPattern.IsMatch(parts[0]))
            {
                if (int.TryParse(parts[2], out _) && !int.TryParse(parts[3], out _))
                    return $"{parts[0]}-{parts[1]}-{parts[2]}";
            }
            return cleaned;
        }

        /// <summary>
        /// A Win32Window adapter for WPF interop
        /// </summary>
        public class Win32Window : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public Win32Window(IntPtr handle)
            {
                Handle = handle;
            }
        }
    }
}