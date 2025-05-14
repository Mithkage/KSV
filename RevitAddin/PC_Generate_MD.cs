using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel; // Added for ClosedXML
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms; // For SaveFileDialog and MessageBox
// System.Drawing.Color will be used with XLColor.FromColor() or XLColor.FromArgb()

namespace PC_Generate_MD
{

    // Helper class to store submain data entries (remains the same)
    public class SubmainEntry
    {
        public string PC_From { get; set; }
        public string PC_SWB_to { get; set; }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_Generate_MDClass : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string projectName = doc.ProjectInformation.Name;
            string projectNumber = doc.ProjectInformation.Number;
            string projectAddress = doc.ProjectInformation.Address?.Replace("\r\n", ", ");

            if (string.IsNullOrEmpty(projectName)) projectName = "N/A";
            if (string.IsNullOrEmpty(projectNumber)) projectNumber = "N/A";
            if (string.IsNullOrEmpty(projectAddress)) projectAddress = "N/A";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Save Excel Report";
            saveFileDialog.FileName = $"ProjectReport_{doc.Title.Replace(".rvt", "")}.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(saveFileDialog.FileName))
                {
                    try
                    {
                        GenerateExcelReport_ClosedXML(saveFileDialog.FileName, projectName, projectNumber, projectAddress, doc);
                        MessageBox.Show($"Report generated successfully to:\n{saveFileDialog.FileName}", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        message = $"Error generating report: {ex.Message}\n{ex.StackTrace}";
                        MessageBox.Show(message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return Result.Failed;
                    }
                }
                else
                {
                    message = "File name not provided. Export cancelled.";
                    MessageBox.Show(message, "Export Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return Result.Cancelled;
                }
            }
            else
            {
                return Result.Cancelled;
            }
        }

        public void GenerateExcelReport_ClosedXML(string filePath, string revitProjectName, string revitProjectNumber, string revitProjectAddress, Document revitDoc)
        {
            using (var workbook = new XLWorkbook())
            {
                // --- 1. Set Default Workbook Font ("Normal" Style) ---
                workbook.Style.Font.FontName = "Calibri";
                workbook.Style.Font.FontSize = 10;

                // --- Create data-specific sheets first ---
                CreateSubmainsSheet_ClosedXML(workbook, revitDoc);

                // --- Create sheets that might have TOCs including the above sheets ---
                var coverSheet = workbook.Worksheets.Add("Cover Page");
                CreateSheetLayout_ClosedXML(coverSheet, "Maximum Demand Calculation", workbook, revitProjectName, revitProjectNumber, revitProjectAddress);

                var authoritySheet = workbook.Worksheets.Add("Authority");
                CreateSheetLayout_ClosedXML(authoritySheet, "Maximum Demand Calculation", workbook, revitProjectName, revitProjectNumber, revitProjectAddress);

                workbook.SaveAs(filePath);
            }
        }

        private List<SubmainEntry> GetSubmainsData(Document revitDoc) // Remains the same
        {
            var collectedData = new Dictionary<string, string>();
            FilteredElementCollector collector = new FilteredElementCollector(revitDoc);
            IList<Element> detailItems = collector.OfClass(typeof(FamilyInstance))
                                                  .OfCategory(BuiltInCategory.OST_DetailComponents)
                                                  .WhereElementIsNotElementType()
                                                  .ToList();
            foreach (Element elem in detailItems)
            {
                Parameter powerCadParam = elem.LookupParameter("PC_PowerCAD");
                if (powerCadParam != null && powerCadParam.HasValue && powerCadParam.StorageType == StorageType.Integer)
                {
                    if (powerCadParam.AsInteger() == 1)
                    {
                        Parameter fromParam = elem.LookupParameter("PC_From");
                        Parameter swbToParam = elem.LookupParameter("PC_SWB to");
                        string pcFrom = fromParam?.AsString();
                        string pcSwbTo = swbToParam?.AsString();
                        if (!string.IsNullOrEmpty(pcSwbTo) && pcFrom != null)
                        {
                            collectedData[pcSwbTo] = pcFrom;
                        }
                    }
                }
            }
            return collectedData
                .Select(kvp => new SubmainEntry { PC_SWB_to = kvp.Key, PC_From = kvp.Value })
                .OrderBy(entry => entry.PC_SWB_to)
                .ToList();
        }

        private void CreateSubmainsSheet_ClosedXML(XLWorkbook workbook, Document revitDoc)
        {
            var worksheet = workbook.Worksheets.Add("Submains");
            List<SubmainEntry> submainsData = GetSubmainsData(revitDoc);

            // Add Headers
            var headerFromCell = worksheet.Cell(1, 1);
            headerFromCell.Value = "PC_From";
            ApplyHeaderStyle_ClosedXML(headerFromCell);

            var headerToCell = worksheet.Cell(1, 2);
            headerToCell.Value = "PC_SWB to";
            ApplyHeaderStyle_ClosedXML(headerToCell);


            // Add Data and Create Table
            if (submainsData.Any())
            {
                var table = worksheet.Cell(2, 1).InsertTable(submainsData, "TB_Submains", true);
                table.Theme = XLTableTheme.TableStyleMedium9;
            }
            else
            {
                worksheet.Cell(2, 1).Value = "[No Submain data found with PC_PowerCAD = Yes]";
                var rangeForEmptyTable = worksheet.Range(1, 1, 1, 2);
                if (worksheet.Tables.FirstOrDefault(t => t.Name == "TB_Submains") == null)
                {
                    var table = rangeForEmptyTable.CreateTable("TB_Submains");
                    table.Theme = XLTableTheme.TableStyleMedium9;
                }
            }

            worksheet.Columns().AdjustToContents();

            // Page Setup
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.FitToPages(1, 0);
            worksheet.PageSetup.Margins.Top = 0.75;
            worksheet.PageSetup.Margins.Bottom = 0.75;
            worksheet.PageSetup.Margins.Left = 0.7;
            worksheet.PageSetup.Margins.Right = 0.7;
            worksheet.PageSetup.Margins.Header = 0.3;
            worksheet.PageSetup.Margins.Footer = 0.3;
            worksheet.PageSetup.CenterHorizontally = true;

            // Footer Setup - Corrected for older ClosedXML API pattern
            var hfRight = worksheet.PageSetup.Footer.Right; // IXLHFItem
            hfRight.Clear(); // Clear previous content
            // AddText returns IXLRichText which can be styled
            hfRight.AddText("Page &P of &N")
                   .SetFontName("Calibri")
                   .SetFontSize(9)
                   .SetFontColor(XLColor.Gray);
        }

        private void CreateSheetLayout_ClosedXML(IXLWorksheet worksheet, string mainPageTitle, XLWorkbook workbook,
                                              string revitProjectName, string revitProjectNumber, string revitProjectAddress)
        {
            #region Page Layout and Column Widths
            worksheet.Column(1).Width = 5;
            worksheet.Column(2).Width = 25;
            worksheet.Column(3).Width = 40;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 5;
            worksheet.ShowGridLines = false;
            #endregion

            #region Sheet Content
            int currentRow = 2;
            int mainContentColumn = 3;
            int labelColumn = 2;
            int valueColumn = 3;

            var titleCell = worksheet.Cell(currentRow, mainContentColumn);
            titleCell.Value = mainPageTitle;
            ApplyTitleStyle_ClosedXML(titleCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(16, 1.5, 15);
            currentRow++;

            var sheetNameCell = worksheet.Cell(currentRow, mainContentColumn);
            sheetNameCell.Value = worksheet.Name;
            ApplyHeading1Style_ClosedXML(sheetNameCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(14, 1.5, 15);
            currentRow += 3;

            var projectDetailsHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
            projectDetailsHeaderCell.Value = "Project Details";
            ApplyHeading2Style_ClosedXML(projectDetailsHeaderCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
            currentRow++;

            Action<string, string> addProjectDetailRow = (label, value) =>
            {
                var labelCell = worksheet.Cell(currentRow, labelColumn);
                labelCell.Value = label;
                ApplyInfoTextStyle_ClosedXML(labelCell);

                var valueCell = worksheet.Cell(currentRow, valueColumn);
                valueCell.Value = value;
                ApplyInfoTextStyle_ClosedXML(valueCell);
                currentRow++;
            };

            addProjectDetailRow("Document Number:", "E-MD-001");
            addProjectDetailRow("Revision:", "1");
            addProjectDetailRow("Project Name:", revitProjectName);
            addProjectDetailRow("Project Number:", revitProjectNumber);
            addProjectDetailRow("Project Address:", revitProjectAddress);
            currentRow += 2;

            addProjectDetailRow("Engineer:", "Kyle Vorster");
            addProjectDetailRow("Date:", "13.05.2025");
            addProjectDetailRow("Issue Purpose:", "For Review");
            addProjectDetailRow("Reference Schematic:", "");
            currentRow += 3;

            var tocHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
            tocHeaderCell.Value = "Table of Contents";
            ApplyHeading2Style_ClosedXML(tocHeaderCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
            currentRow++;

            int pageCounter = 1;
            foreach (var sheet in workbook.Worksheets)
            {
                var tocItemCell = worksheet.Cell(currentRow, labelColumn);
                tocItemCell.Value = sheet.Name;
                ApplyGenericContentStyle_ClosedXML(tocItemCell);

                var tocPageCell = worksheet.Cell(currentRow, valueColumn);
                tocPageCell.Value = $"Page {pageCounter:D2}";
                ApplyGenericContentStyle_ClosedXML(tocPageCell);
                currentRow++;
                pageCounter++;
            }
            currentRow += 3;

            var assumptionsHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
            assumptionsHeaderCell.Value = "Assumptions";
            ApplyHeading2Style_ClosedXML(assumptionsHeaderCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
            currentRow++;
            var assumptionsTextCell = worksheet.Cell(currentRow, labelColumn);
            assumptionsTextCell.Value = "[Specify assumptions here...]";
            ApplyGenericContentStyle_ClosedXML(assumptionsTextCell);
            currentRow += 3;

            var notesHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
            notesHeaderCell.Value = "General Notes";
            ApplyHeading2Style_ClosedXML(notesHeaderCell);
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
            currentRow++;

            var note1Cell = worksheet.Cell(currentRow, labelColumn);
            note1Cell.Value = "1. Where possible, AS3000 Max Demand methods have been used.";
            ApplyGenericContentStyle_ClosedXML(note1Cell);
            currentRow++;
            var note2Cell = worksheet.Cell(currentRow, labelColumn);
            note2Cell.Value = "2. Where deviations or engineering assessment have been applied, it is noted within the page and justified.";
            ApplyGenericContentStyle_ClosedXML(note2Cell);
            currentRow++;
            #endregion

            #region Page Setup
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.FitToPages(1, 1);
            worksheet.PageSetup.Margins.Top = 0.75;
            worksheet.PageSetup.Margins.Bottom = 0.75;
            worksheet.PageSetup.Margins.Left = 0.7;
            worksheet.PageSetup.Margins.Right = 0.7;
            worksheet.PageSetup.Margins.Header = 0.3;
            worksheet.PageSetup.Margins.Footer = 0.3;
            worksheet.PageSetup.CenterHorizontally = true;
            worksheet.PageSetup.CenterVertically = true;
            #endregion

            #region Footer Setup
            // Corrected Footer Setup for older ClosedXML API pattern
            var hfRightLayout = worksheet.PageSetup.Footer.Right; // IXLHFItem
            hfRightLayout.Clear(); // Clear previous content
            // AddText returns IXLRichText which can be styled
            hfRightLayout.AddText("Page &P of &N")
                   .SetFontName("Calibri")
                   .SetFontSize(9)
                   .SetFontColor(XLColor.Gray);
            #endregion
        }

        #region ClosedXML Style Helper Methods
        private double GetRowHeightForStyle(double fontSize, double multiplier, double minHeight)
        {
            return Math.Max(minHeight, fontSize * multiplier);
        }

        private void ApplyHeaderStyle_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 10;
            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyTitleStyle_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 16;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
        }

        private void ApplyHeading1Style_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 14;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyHeading2Style_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 12;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.FromArgb(0, 154, 147); // For #009A93: R=0, G=154, B=147
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyHeading3Style_ClosedXML(IXLCell cell) // Kept for consistency, though not used in current content
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 11;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyInfoTextStyle_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontSize = 12;
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyGenericContentStyle_ClosedXML(IXLCell cell)
        {
            cell.Style.Font.FontName = "Calibri";
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            cell.Style.Alignment.WrapText = true;
        }
        #endregion
    }
}
