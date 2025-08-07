//-----------------------------------------------------------------------------
// <copyright file="ReportLinkViewModel.cs" company="RTS Reports">
//     Copyright (c) RTS Reports. All rights reserved.
// </copyright>
// <summary>
//     Link view model for reports
// </summary>
//-----------------------------------------------------------------------------

namespace RTS.Reports.Models
{
    /// <summary>
    /// View model for Revit link items in reports
    /// </summary>
    public class ReportLinkViewModel
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsLoaded { get; set; }
        public bool IsPlaceholder { get; set; }
    }
}