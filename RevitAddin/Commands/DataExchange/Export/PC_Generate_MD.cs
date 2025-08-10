// File name: PC_Generate_MD.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB; // Required for Revit API types like Document, Parameter, Element, FilteredElementCollector, SpecTypeId, etc.
using Autodesk.Revit.UI; // Required for Revit UI types like UIApplication, UIDocument, IExternalCommand, etc.
using ClosedXML.Excel; // For Excel manipulation
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms; // For SaveFileDialog and MessageBox

// System.Drawing.Color can be used with XLColor.FromColor() or XLColor.FromArgb().
// For example, XLColor.FromArgb(System.Drawing.Color.Red.ToArgb()) or XLColor.FromArgb(R,G,B).
// using System.Drawing; // Uncomment if needed, but XLColor.FromArgb(R,G,B) is used directly.

// Corrected namespace to match the expected reference in ReportSelectionWindow.xaml.cs
namespace RTS.Commands.DataExchange.Export
{
    public class SubmainEntry
    {
        public string PC_From { get; set; }
        public string PC_SWB_to { get; set; }
        public string PC_SWB_Type { get; set; }
        public string MaxDemandA { get; set; }
        public string SubmainDiversityFactor { get; set; } // Was DiversityFactor
        public string PC_SWB_Load { get; set; }
        public string BusDiversity { get; set; } // Was PC_ProtectiveDeviceTripSettingA
        public string BusLoad { get; set; }
        public string Notes { get; set; }

        public SubmainEntry()
        {
            PC_From = string.Empty;
            PC_SWB_to = string.Empty;
            PC_SWB_Type = string.Empty;
            MaxDemandA = string.Empty;
            SubmainDiversityFactor = string.Empty;
            PC_SWB_Load = string.Empty;
            BusDiversity = string.Empty;
            BusLoad = string.Empty;
            Notes = string.Empty;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PC_Generate_MDClass : IExternalCommand
    {
        // GUIDs for the shared parameters.
        private static readonly Guid PC_POWERCAD_GUID = new Guid("8f31d68f-60c9-4ec6-a7ff-78a6e3bdaab6");
        private static readonly Guid PC_FROM_GUID = new Guid("5dd52911-7bcd-4a06-869d-73fcef59951c");
        private static readonly Guid PC_SWB_TO_GUID = new Guid("e142b0ed-d084-447a-991b-d9a3a3f67a8d");
        private static readonly Guid PC_SWB_TYPE_GUID = new Guid("9d5ab9c2-09e2-4d42-ae2f-2a5fba6f7131");

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

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                Title = "Save Excel Report",
                FileName = $"ProjectReport_{doc.Title.Replace(".rvt", "")}.xlsx"
            };

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
                        string fullErrorMessage = $"Error generating report: {ex.Message}\n\nStack Trace: {ex.StackTrace}";
                        if (ex.InnerException != null)
                        {
                            fullErrorMessage += $"\n\nInner Exception: {ex.InnerException.Message}\n\nInner Stack Trace: {ex.InnerException.StackTrace}";
                        }
                        message = fullErrorMessage;
                        MessageBox.Show(fullErrorMessage, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                workbook.Style.Font.FontName = "Calibri";
                workbook.Style.Font.FontSize = 10;

                var coverSheet = workbook.Worksheets.Add("Cover Page");
                CreateSheetLayout_ClosedXML(coverSheet, "Maximum Demand Calculation", workbook, revitProjectName, revitProjectNumber, revitProjectAddress, false);
                coverSheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;

                string tempTemplateSheetName = "TempTemplateSheetForCopying";
                var templateSheet = coverSheet.CopyTo(tempTemplateSheetName);
                int lastRowOnTemplate = templateSheet.LastRowUsed()?.RowNumber() ?? 4;
                if (lastRowOnTemplate >= 5)
                {
                    templateSheet.Rows(5, lastRowOnTemplate).Clear(XLClearOptions.Contents);
                }
                templateSheet.Visibility = XLWorksheetVisibility.Hidden;

                var submainsSheet = templateSheet.CopyTo("Submains");
                submainsSheet.Visibility = XLWorksheetVisibility.Visible;
                submainsSheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
                List<SubmainEntry> submainsData = GetSubmainsData(revitDoc);
                CreateSubmainsSheetFromCopy_ClosedXML(submainsSheet, submainsData);

                bool liftSheetCreated = false;
                List<string> exclusionKeywordsForSwbTo = new List<string> { "MSB", "FDCIE", "EWCIE", "PUMP", "BUS ", "TDB", "TOB" };
                List<string> exclusionKeywordsForSwbType = new List<string> { "TOB", "METER PANEL", "BUS", "EQUIPMENT", "TDB" };

                char[] invalidSheetNameChars = new char[] { ':', '\\', '/', '?', '*', '[', ']' };

                foreach (var entry in submainsData)
                {
                    if (!string.IsNullOrEmpty(entry.PC_SWB_Type) &&
                        exclusionKeywordsForSwbType.Any(keyword => entry.PC_SWB_Type.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping sheet creation for '{entry.PC_SWB_to}' because its type '{entry.PC_SWB_Type}' is in the exclusion list.");
                        continue;
                    }

                    string rawSheetName = entry.PC_SWB_to;
                    if (exclusionKeywordsForSwbTo.Any(keyword => rawSheetName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping sheet creation for '{entry.PC_SWB_to}' because its name contains an exclusion keyword (from PC_SWB_to list).");
                        continue;
                    }

                    string actualSheetNameToCreate;
                    bool isLiftRelated = rawSheetName.IndexOf("LIFT", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isLiftRelated)
                    {
                        if (liftSheetCreated) continue;
                        actualSheetNameToCreate = "LIFT";
                    }
                    else
                    {
                        actualSheetNameToCreate = rawSheetName;
                    }

                    string sanitizedSheetName = new string(actualSheetNameToCreate.Where(ch => !invalidSheetNameChars.Contains(ch)).ToArray()).Trim();
                    if (string.IsNullOrWhiteSpace(sanitizedSheetName))
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping sheet creation for '{rawSheetName}' due to invalid sanitized name.");
                        continue;
                    }
                    if (sanitizedSheetName.Length > 31) sanitizedSheetName = sanitizedSheetName.Substring(0, 31);

                    if (workbook.Worksheets.Any(ws => ws.Name.Equals(sanitizedSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        if (isLiftRelated && sanitizedSheetName.Equals("LIFT", StringComparison.OrdinalIgnoreCase)) liftSheetCreated = true;
                        continue;
                    }

                    IXLWorksheet newSheet = templateSheet.CopyTo(sanitizedSheetName);
                    newSheet.Visibility = XLWorksheetVisibility.Visible;
                    newSheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;

                    if (isLiftRelated) liftSheetCreated = true;
                }

                foreach (var worksheet in workbook.Worksheets.ToList())
                {
                    if (worksheet.Name.IndexOf("HDB", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        AddHDBContentStructure(worksheet);
                    }
                    else if (worksheet.Name.IndexOf("MP", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        AddMPContentStructure(worksheet);
                    }
                }

                workbook.Worksheet(tempTemplateSheetName).Delete();

                var coverSheetForSort = workbook.Worksheet("Cover Page");
                var submainsSheetForSort = workbook.Worksheet("Submains");

                var otherSheets = workbook.Worksheets
                                                .Where(ws => ws.Name != "Cover Page" && ws.Name != "Submains")
                                                .OrderBy(ws => ws.Name, StringComparer.OrdinalIgnoreCase)
                                                .ToList();

                if (coverSheetForSort != null) coverSheetForSort.Position = 1;
                if (submainsSheetForSort != null) submainsSheetForSort.Position = 2;

                int currentPosition = 3;
                foreach (var sheet in otherSheets)
                {
                    sheet.Position = currentPosition;
                    currentPosition++;
                }

                PopulateTableOfContents(coverSheetForSort, workbook);
                workbook.SaveAs(filePath);
            }
        }

        private void PopulateTableOfContents(IXLWorksheet coverSheet, XLWorkbook workbook)
        {
            if (coverSheet == null)
            {
                System.Diagnostics.Debug.WriteLine("Cover Page sheet not found. Cannot populate TOC.");
                return;
            }

            var tocHeaderCell = coverSheet.Search("Table of Contents").FirstOrDefault();
            if (tocHeaderCell == null)
            {
                System.Diagnostics.Debug.WriteLine("TOC Header cell not found on Cover Page. Cannot populate TOC.");
                return;
            }

            int currentRow = tocHeaderCell.Address.RowNumber + 1;
            int tocLabelColumn = 4;
            int tocValueColumn = 5;

            var assumptionsHeader = coverSheet.Search("Assumptions").FirstOrDefault();
            int endClearRowLimit = assumptionsHeader != null ? assumptionsHeader.Address.RowNumber - 1 : currentRow + 50;
            if (endClearRowLimit < currentRow) endClearRowLimit = currentRow + 50;

            coverSheet.Range(currentRow, tocLabelColumn, endClearRowLimit, tocValueColumn).Clear(XLClearOptions.Contents);


            int pageCounter = 1;
            foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
            {
                var tocItemCell = coverSheet.Cell(currentRow, tocLabelColumn);
                tocItemCell.Value = sheet.Name;
                ApplyGenericContentStyle_ClosedXML(tocItemCell);

                var tocPageCell = coverSheet.Cell(currentRow, tocValueColumn);
                tocPageCell.Value = $"Page {pageCounter:D2}";
                ApplyGenericContentStyle_ClosedXML(tocPageCell);
                currentRow++;
                pageCounter++;
            }
        }

        private List<SubmainEntry> GetSubmainsData(Document revitDoc)
        {
            List<SubmainEntry> allQualifyingEntries = new List<SubmainEntry>();
            FilteredElementCollector collector = new FilteredElementCollector(revitDoc);
            IList<Element> detailItems = collector.OfClass(typeof(FamilyInstance))
                                                   .OfCategory(BuiltInCategory.OST_DetailComponents)
                                                   .WhereElementIsNotElementType()
                                                   .ToList();

            foreach (Element elem in detailItems)
            {
                Parameter powerCadParam = elem.get_Parameter(PC_POWERCAD_GUID);
                if (powerCadParam != null && powerCadParam.HasValue && powerCadParam.AsInteger() == 1)
                {
                    Parameter fromParam = elem.get_Parameter(PC_FROM_GUID);
                    Parameter swbToParam = elem.get_Parameter(PC_SWB_TO_GUID);

                    string pcFromValue = fromParam?.AsString() ?? string.Empty;
                    string pcSwbToValue = swbToParam?.AsString() ?? string.Empty;

                    string pcSwbTypeValue = string.Empty;
                    ElementType elementType = revitDoc.GetElement(elem.GetTypeId()) as ElementType;
                    if (elementType != null)
                    {
                        Parameter swbTypeParam = elementType.get_Parameter(PC_SWB_TYPE_GUID);
                        pcSwbTypeValue = swbTypeParam?.AsString() ?? string.Empty;
                    }

                    if (!string.IsNullOrEmpty(pcSwbToValue))
                    {
                        allQualifyingEntries.Add(new SubmainEntry
                        {
                            PC_From = pcFromValue,
                            PC_SWB_to = pcSwbToValue,
                            PC_SWB_Type = pcSwbTypeValue,
                            MaxDemandA = "",
                            SubmainDiversityFactor = "",
                            PC_SWB_Load = "",
                            BusDiversity = "",
                            BusLoad = "",
                            Notes = ""
                        });
                    }
                }
            }

            var groupedEntries = allQualifyingEntries.GroupBy(entry => entry.PC_SWB_to);
            List<SubmainEntry> finalUniqueEntries = new List<SubmainEntry>();
            foreach (var group in groupedEntries)
            {
                string keySwbTo = group.Key;
                string propagatedPcFrom = group.Select(e => e.PC_From)
                                                   .FirstOrDefault(pf => !string.IsNullOrEmpty(pf)) ?? string.Empty;

                var firstEntryInGroup = group.First();

                finalUniqueEntries.Add(new SubmainEntry
                {
                    PC_SWB_to = keySwbTo,
                    PC_From = propagatedPcFrom,
                    PC_SWB_Type = firstEntryInGroup.PC_SWB_Type,
                    MaxDemandA = firstEntryInGroup.MaxDemandA,
                    SubmainDiversityFactor = firstEntryInGroup.SubmainDiversityFactor,
                    PC_SWB_Load = firstEntryInGroup.PC_SWB_Load,
                    BusDiversity = firstEntryInGroup.BusDiversity,
                    BusLoad = firstEntryInGroup.BusLoad,
                    Notes = firstEntryInGroup.Notes
                });
            }

            return finalUniqueEntries
                .OrderBy(entry => entry.PC_From, StringComparer.OrdinalIgnoreCase)
                .ThenBy(entry => entry.PC_SWB_to, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void CreateSubmainsSheetFromCopy_ClosedXML(IXLWorksheet worksheet, List<SubmainEntry> submainsData)
        {
            worksheet.Range("D5:L150").Clear(XLClearOptions.All);

            int startHeaderRow = 7;
            int startDataRow = 8;
            int startColumn = 4;

            worksheet.Cell(startHeaderRow, startColumn).Value = "PC_From";
            worksheet.Cell(startHeaderRow, startColumn + 1).Value = "PC_SWB to";
            worksheet.Cell(startHeaderRow, startColumn + 2).Value = "Switchboard Type";
            worksheet.Cell(startHeaderRow, startColumn + 3).Value = "Max Demand (A)";
            worksheet.Cell(startHeaderRow, startColumn + 4).Value = "Submain Diversity Factor";
            worksheet.Cell(startHeaderRow, startColumn + 5).Value = "PC_SWB Load";
            worksheet.Cell(startHeaderRow, startColumn + 6).Value = "Bus Diversity";
            worksheet.Cell(startHeaderRow, startColumn + 7).Value = "Bus Load";
            worksheet.Cell(startHeaderRow, startColumn + 8).Value = "Notes";

            for (int i = 0; i < 9; i++)
            {
                ApplyHeaderStyle_ClosedXML(worksheet.Cell(startHeaderRow, startColumn + i));
            }

            if (submainsData.Any())
            {
                worksheet.Cell(startDataRow, startColumn).InsertData(submainsData);

                if (submainsData.Count > 0)
                {
                    var tableRange = worksheet.Range(startHeaderRow, startColumn, startHeaderRow + submainsData.Count, startColumn + 8);
                    var excelTable = tableRange.CreateTable("TB_Submains");
                    excelTable.Theme = XLTableTheme.TableStyleMedium9;
                    excelTable.ShowHeaderRow = true;

                    worksheet.Column(startColumn + 2).Style.Alignment.WrapText = true;
                    worksheet.Column(startColumn + 3).Style.NumberFormat.Format = "0.0";
                    worksheet.Column(startColumn + 4).Style.NumberFormat.Format = "0%";
                    worksheet.Column(startColumn + 5).Style.NumberFormat.Format = "0.0";
                    worksheet.Column(startColumn + 6).Style.NumberFormat.Format = "0%";
                    worksheet.Column(startColumn + 7).Style.NumberFormat.Format = "0";
                    worksheet.Column(startColumn + 8).Style.Alignment.WrapText = false;
                }
            }
            else
            {
                worksheet.Cell(startDataRow, startColumn).Value = "[No Submain data found with PC_PowerCAD = Yes]";
            }

            for (int i = 0; i < 9; i++)
            {
                worksheet.Column(startColumn + i).AdjustToContents();
            }

            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.FitToPages(1, 0);
            worksheet.PageSetup.CenterHorizontally = true;
        }

        private void AddHDBContentStructure(IXLWorksheet worksheet)
        {
            int lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 5;
            worksheet.Range(5, 4, Math.Max(5, lastRowUsed), 12).Clear(XLClearOptions.All);

            int currentRow = 5;
            int startColD = 4;

            var cellSwitchboardName = worksheet.Cell(currentRow, startColD);
            cellSwitchboardName.Value = "Switchboard Name: " + worksheet.Name;
            cellSwitchboardName.Style.Font.Bold = true;
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 8).Merge(); // REMOVED MERGE
            currentRow++;

            var cellAsNzsNote = worksheet.Cell(currentRow, startColD);
            cellAsNzsNote.Value = "AS/NZS 3000 Table C2 Column 3: Factories, Shops, Offices (Note, adjust diversities for application of Table C2 column 2)";
            cellAsNzsNote.Style.Font.FontSize = 9;
            cellAsNzsNote.Style.Alignment.WrapText = true; // Wrap text is important if not merging
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 8).Merge(); // REMOVED MERGE
            worksheet.Row(currentRow).AdjustToContents();
            currentRow++;

            string[] headers = {
                "Load Group", "Description", "Quantity", "Unit Load VA", "Total Load VA",
                "AS3000 Diversity", "Additional Diversity", "Total Load VA (with Diversity)", "Totals in VA"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                var headerCell = worksheet.Cell(currentRow, startColD + i);
                headerCell.Value = headers[i];
                ApplyHeaderStyle_ClosedXML(headerCell);
            }
            currentRow++;

            var cellLoadGroupA = worksheet.Cell(currentRow, startColD);
            cellLoadGroupA.Value = "A: Lighting other than in load group F(100% connected load)";
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 1).Merge(); // REMOVED MERGE
            cellLoadGroupA.Style.Font.Bold = true;
            cellLoadGroupA.Style.Alignment.WrapText = true;
            // If you want the text to span visually, you might need to write to subsequent cells or adjust column widths.
            // For now, it will be in the first cell (startColD).
            worksheet.Row(currentRow).AdjustToContents();
            currentRow += 24;

            var cellLoadGroupB1 = worksheet.Cell(currentRow, startColD);
            cellLoadGroupB1.Value = "B(i): Socket outlets not exceeding 10A other than those in B (1000W + 750W each) Buildings without heating";
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 1).Merge(); // REMOVED MERGE
            cellLoadGroupB1.Style.Font.Bold = true;
            cellLoadGroupB1.Style.Alignment.WrapText = true;
            worksheet.Row(currentRow).AdjustToContents();
            currentRow += 5;

            var cellLoadGroupB2 = worksheet.Cell(currentRow, startColD);
            cellLoadGroupB2.Value = "B(ii): Socket outlets not exceeding 10A in buildings or portions of buildings provided with permantently installed heating or cooling equipment or both (1000W + 100W each)";
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 1).Merge(); // REMOVED MERGE
            cellLoadGroupB2.Style.Font.Bold = true;
            cellLoadGroupB2.Style.Alignment.WrapText = true;
            worksheet.Row(currentRow).AdjustToContents();
            currentRow += 3;

            var cellLoadGroupB3 = worksheet.Cell(currentRow, startColD);
            cellLoadGroupB3.Value = "B(iii): Socket outlets exceeding 10A (Full current rating of highest Rated socket outlet plus 75% of full current current rating of remainder)";
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 1).Merge(); // REMOVED MERGE
            cellLoadGroupB3.Style.Font.Bold = true;
            cellLoadGroupB3.Style.Alignment.WrapText = true;
            worksheet.Row(currentRow).AdjustToContents();
            currentRow += 10;

            worksheet.Column(startColD).Width = 25;
            worksheet.Column(startColD + 1).Width = 35;
            worksheet.Column(startColD + 2).Width = 10;
            worksheet.Column(startColD + 3).Width = 15;
            worksheet.Column(startColD + 4).Width = 15;
            worksheet.Column(startColD + 5).Width = 18;
            worksheet.Column(startColD + 6).Width = 18;
            worksheet.Column(startColD + 7).Width = 22;
            worksheet.Column(startColD + 8).Width = 15;
        }

        private void AddMPContentStructure(IXLWorksheet worksheet)
        {
            int lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 5;
            worksheet.Range(5, 4, Math.Max(5, lastRowUsed), 15).Clear(XLClearOptions.All);

            int currentRow = 5;
            int startColD = 4;

            var title1 = worksheet.Cell(currentRow, startColD);
            title1.Value = "AS/NZS 3000: Table C1";
            title1.Style.Font.Bold = true;
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 8).Merge(); // REMOVED MERGE
            currentRow++;

            var title2 = worksheet.Cell(currentRow, startColD);
            title2.Value = "Maximum Demand - Single and Multiple Domestic Electrical Installations";
            title2.Style.Font.Bold = true;
            // worksheet.Range(currentRow, startColD, currentRow, startColD + 8).Merge(); // REMOVED MERGE
            currentRow++;
            currentRow++;

            worksheet.Cell(currentRow, startColD + 2).Value = "Total Blocks of Living Units:";
            worksheet.Cell(currentRow, startColD + 3).Value = 29;
            var totalsHeaderCell = worksheet.Cell(currentRow, startColD + 6);
            totalsHeaderCell.Value = "TOTALS";
            totalsHeaderCell.Style.Font.Bold = true;
            // worksheet.Range(currentRow, startColD + 6, currentRow, startColD + 8).Merge().Style.Font.Bold = true; // REMOVED MERGE
            currentRow++;
            worksheet.Cell(currentRow, startColD + 2).Value = "Living Units Per Phase:";
            worksheet.Cell(currentRow, startColD + 3).Value = 10;
            currentRow++;
            var loadAssociatedCell = worksheet.Cell(currentRow, startColD + 2);
            loadAssociatedCell.Value = "Load associated with individual units";
            loadAssociatedCell.Style.Font.Bold = true;
            // worksheet.Range(currentRow, startColD + 2, currentRow, startColD + 5).Merge().Style.Font.Bold = true; // REMOVED MERGE
            currentRow++;

            string[] mpHeaders = {
                "Load Group",
                "",
                "Single Domestic Electrial Installation or individual living unit per phase",
                "2 to 5 living units per phase",
                "6 to 20 living units per phase",
                "21 or more living units per phase",
                "TOTALS"
            };
            for (int i = 0; i < mpHeaders.Length; i++)
            {
                var headerCell = worksheet.Cell(currentRow, startColD + i);
                headerCell.Value = mpHeaders[i];
                ApplyHeaderStyle_ClosedXML(headerCell);
                headerCell.Style.Alignment.WrapText = true;
            }
            worksheet.Row(currentRow).Height = 45;
            currentRow++;

            Action<string, string, string, string, string, string, string, int> addMpLoadGroupRow =
                (group, desc, valF, valG, valH, valI, valJ, blankRowsAfter) =>
                {
                    var groupCell = worksheet.Cell(currentRow, startColD);
                    groupCell.Value = group;
                    groupCell.Style.Font.Bold = true;
                    groupCell.Style.Alignment.WrapText = true;

                    var descCell = worksheet.Cell(currentRow, startColD + 1);
                    descCell.Value = desc;
                    descCell.Style.Alignment.WrapText = true;

                    worksheet.Cell(currentRow, startColD + 2).Value = valF;
                    worksheet.Cell(currentRow, startColD + 3).Value = valG;
                    worksheet.Cell(currentRow, startColD + 4).Value = valH;
                    worksheet.Cell(currentRow, startColD + 5).Value = valI;
                    worksheet.Cell(currentRow, startColD + 6).Value = valJ;

                    worksheet.Row(currentRow).AdjustToContents();
                    currentRow += 1 + blankRowsAfter;
                };

            addMpLoadGroupRow("A. Lighting", "(i) Expect as in (ii) and Load Group H below", "3 A for 1 to 20 points + 2 A for each additional 20 points or part thereof", "6A", "5 + 0.25 A per living unit", "0.5 A per living unit", "8A", 1);
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Number of Points:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 10;
            addMpLoadGroupRow("", "(ii) Outdoor lighting exceeding a total of 1000W", "75% Connected Load", "No assessement for the purpose of maximum demand", "", "", "", 0);

            addMpLoadGroupRow("B. General Power", "(i) Socket-outlets not exceeding 10A. Permanently connected electrical equipment not exceeding 10A and not included in other load group", "10A for 1 to 20 Points + 5A for each additional 20 points or part there of", "10A + 5A per living unit", "15A + 3.75A per living unit", "50A + 1.9A per living unit", "53A", 1);
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Number of Points:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 5;
            addMpLoadGroupRow("", "(ii) Where the electrical installation includes one or more 15A socket-outlets other than socket-outlets provided to supply electrical equipement set out in group C, D, E, F G and L", "10A", "10A", "10A", "10A", "10A", 0);
            addMpLoadGroupRow("", "(iii) Where the electrical installation includes one or more 20A socket-outlets other than socket-outlets provided to supply electrical equipment set out in gorups C, D, E, F, G and L", "15A", "15A", "15A", "15A", "15A", 0);

            addMpLoadGroupRow("C. Appliances Power", "Ranges, cooking appliances, laundry equipment or socket-outlets rated at more than 10A for the connection thereof", "50% Connected Load", "15A", "2.8A per living unit", "2.8A per living unit", "28A", 1);
            worksheet.Cell(currentRow - 2, startColD + 2).Value = "0.5";
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Connected Load:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 20;

            addMpLoadGroupRow("D. Heating and Air-Conditioning", "Fixed space heating or air-conditioning equipement, saunas or socket-outlets rated at more than 10A for the connection thereof", "75% Connected Load", "75% Connected Load", "75% Connected Load", "75% Connected Load", "A", 1);
            worksheet.Cell(currentRow - 2, startColD + 2).Value = "0.75";
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Connected Load:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 0;

            addMpLoadGroupRow("E.Instantaneous water heaters", "Instantaneous Water Heaters", "33.3% Connected Load", "6A per living unit", "6A per living unit", "100A + 0.8A per living unit", "", 1);
            worksheet.Cell(currentRow - 2, startColD + 2).Value = "0.333";
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Connected Load:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 0;

            addMpLoadGroupRow("F. Storage Water Heaters", "Storage Water Heaters", "Full-Load Current", "6A per living unit", "6A per living unit", "100A + 0.8A per living unit", "", 1);
            worksheet.Cell(currentRow - 2, startColD + 3).Value = "Connected Load:";
            worksheet.Cell(currentRow - 1, startColD + 3).Value = 0;

            addMpLoadGroupRow("G. Spa and Swmming Pool Heater", "Spa and Swimming pool heater", "75% of the largest spa, plus 75% of the largest swimming pool, plus 25% of the remainder", "", "", "", "", 0);
            currentRow++;

            worksheet.Cell(currentRow, startColD + 5).Value = "Total Amps per Phase";
            worksheet.Cell(currentRow, startColD + 6).Value = "88 A";
            worksheet.Cell(currentRow, startColD + 5).Style.Font.Bold = true; // Apply bold to the label cell
            currentRow++;
            worksheet.Cell(currentRow, startColD + 5).Value = "Total in kVA";
            worksheet.Cell(currentRow, startColD + 6).Value = "60.90kVA";
            worksheet.Cell(currentRow, startColD + 5).Style.Font.Bold = true; // Apply bold to the label cell
            currentRow++;
            worksheet.Cell(currentRow, startColD + 5).Value = "kVA Per Apt";
            worksheet.Cell(currentRow, startColD + 6).Value = "2.10kVA";
            worksheet.Cell(currentRow, startColD + 5).Style.Font.Bold = true; // Apply bold to the label cell

            worksheet.Column(startColD).Width = 20;
            worksheet.Column(startColD + 1).Width = 40;
            worksheet.Column(startColD + 2).Width = 30;
            worksheet.Column(startColD + 3).Width = 25;
            worksheet.Column(startColD + 4).Width = 25;
            worksheet.Column(startColD + 5).Width = 25;
            worksheet.Column(startColD + 6).Width = 15;
        }

        private void CreateSheetLayout_ClosedXML(IXLWorksheet worksheet, string mainPageTitle, XLWorkbook workbook,
                                                string revitProjectName, string revitProjectNumber, string revitProjectAddress, bool populateTOCItems)
        {
            int labelColumn = 2;
            int valueColumn = 3;
            int mainContentColumn = 3;

            if (worksheet.Name == "Cover Page" || worksheet.Name == "TempTemplateSheetForCopying")
            {
                if (worksheet.Column(1).Width != 1.4) worksheet.Column(1).InsertColumnsBefore(1);
                if (worksheet.Column(2).Width != 1.4) worksheet.Column(1).InsertColumnsAfter(1);
                if (worksheet.Column(3).Width != 1.4) worksheet.Column(2).InsertColumnsAfter(1);

                worksheet.Column(1).Width = 1.4;
                worksheet.Column(2).Width = 1.4;
                worksheet.Column(3).Width = 1.4;

                labelColumn = 4;
                valueColumn = 5;
                mainContentColumn = 4;

                worksheet.Column(labelColumn).Width = 25;
                worksheet.Column(valueColumn).Width = 40;
                worksheet.Column(valueColumn + 1).Width = 20;
                worksheet.Column(valueColumn + 2).Width = 5;
            }
            worksheet.ShowGridLines = false;

            #region Sheet Content
            int currentRow = 2;

            var titleCell = worksheet.Cell(currentRow, mainContentColumn);
            titleCell.Value = mainPageTitle;
            ApplyTitleStyle_ClosedXML(titleCell); // Centering is part of this style
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(16, 1.5, 15);
            currentRow++;

            var sheetNameCell = worksheet.Cell(currentRow, mainContentColumn);
            sheetNameCell.FormulaA1 = @"MID(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))+1,255)";
            ApplyHeading1Style_ClosedXML(sheetNameCell); // Centering is part of this style
            worksheet.Row(currentRow).Height = GetRowHeightForStyle(14, 1.5, 15);
            currentRow += 3;

            if (worksheet.Name == "Cover Page" || worksheet.Name == "TempTemplateSheetForCopying")
            {
                var projectDetailsHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
                projectDetailsHeaderCell.Value = "Project Details";
                ApplyHeading2Style_ClosedXML(projectDetailsHeaderCell); // Centering is part of this style
                worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
                currentRow++;

                Action<string, string> addProjectDetailRow = (label, value) =>
                {
                    var labelCell = worksheet.Cell(currentRow, labelColumn);
                    labelCell.Value = label;
                    ApplyInfoTextStyle_ClosedXML(labelCell); // Left aligned

                    var valueCell = worksheet.Cell(currentRow, valueColumn);
                    valueCell.Value = value;
                    ApplyInfoTextStyle_ClosedXML(valueCell); // Left aligned
                    currentRow++;
                };

                addProjectDetailRow("Document Number:", "E-MD-001");
                addProjectDetailRow("Revision:", "1");
                addProjectDetailRow("Project Name:", revitProjectName);
                addProjectDetailRow("Project Number:", revitProjectNumber);
                addProjectDetailRow("Project Address:", revitProjectAddress);
                currentRow += 2;

                addProjectDetailRow("Engineer:", "Kyle Vorster");
                addProjectDetailRow("Date:", DateTime.Now.ToString("dd.MM.yyyy"));
                addProjectDetailRow("Issue Purpose:", "For Review");
                addProjectDetailRow("Reference Schematic:", "");
                currentRow += 3;

                var tocHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
                tocHeaderCell.Value = "Table of Contents";
                ApplyHeading2Style_ClosedXML(tocHeaderCell); // Centering is part of this style
                worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
                currentRow++;

                if (!populateTOCItems && worksheet.Name == "TempTemplateSheetForCopying")
                {
                    int estimatedTocRows = 10;
                    for (int i = 0; i < estimatedTocRows; i++)
                    {
                        ApplyGenericContentStyle_ClosedXML(worksheet.Cell(currentRow + i, labelColumn));
                        ApplyGenericContentStyle_ClosedXML(worksheet.Cell(currentRow + i, valueColumn));
                    }
                    currentRow += estimatedTocRows;
                }

                var assumptionsHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
                assumptionsHeaderCell.Value = "Assumptions";
                ApplyHeading2Style_ClosedXML(assumptionsHeaderCell); // Centering is part of this style
                worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
                currentRow++;
                var assumptionsTextCell = worksheet.Cell(currentRow, labelColumn);
                assumptionsTextCell.Value = "[Specify assumptions here...]";
                ApplyGenericContentStyle_ClosedXML(assumptionsTextCell); // Left aligned, wrap text
                // if (worksheet.Name == "Cover Page" || worksheet.Name == "TempTemplateSheetForCopying") { worksheet.Range(currentRow, labelColumn, currentRow, valueColumn).Merge(); } // REMOVED MERGE
                currentRow += 3;

                var notesHeaderCell = worksheet.Cell(currentRow, mainContentColumn);
                notesHeaderCell.Value = "General Notes";
                ApplyHeading2Style_ClosedXML(notesHeaderCell); // Centering is part of this style
                worksheet.Row(currentRow).Height = GetRowHeightForStyle(12, 1.5, 15);
                currentRow++;

                var note1Cell = worksheet.Cell(currentRow, labelColumn);
                note1Cell.Value = "1. Where possible, AS3000 Max Demand methods have been used.";
                ApplyGenericContentStyle_ClosedXML(note1Cell); // Left aligned, wrap text
                // if (worksheet.Name == "Cover Page" || worksheet.Name == "TempTemplateSheetForCopying") { worksheet.Range(currentRow, labelColumn, currentRow, valueColumn).Merge(); } // REMOVED MERGE
                currentRow++;

                var note2Cell = worksheet.Cell(currentRow, labelColumn);
                note2Cell.Value = "2. Where deviations or engineering assessment have been applied, it is noted within the page and justified.";
                ApplyGenericContentStyle_ClosedXML(note2Cell); // Left aligned, wrap text
                // if (worksheet.Name == "Cover Page" || worksheet.Name == "TempTemplateSheetForCopying") { worksheet.Range(currentRow, labelColumn, currentRow, valueColumn).Merge(); } // REMOVED MERGE
                currentRow++;
            }
            #endregion

            #region Page Setup
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.Margins.Top = 0.75;
            worksheet.PageSetup.Margins.Bottom = 0.75;
            worksheet.PageSetup.Margins.Left = 0.7;
            worksheet.PageSetup.Margins.Right = 0.7;
            worksheet.PageSetup.Margins.Header = 0.3;
            worksheet.PageSetup.Margins.Footer = 0.3;
            worksheet.PageSetup.CenterHorizontally = true;
            #endregion

            if (worksheet.Name == "Cover Page")
            {
                var rangeD2D3_Titles = worksheet.Range(2, labelColumn, 3, labelColumn);
                rangeD2D3_Titles.Style.NumberFormat.Format = "@";
                rangeD2D3_Titles.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                rangeD2D3_Titles.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rangeD2D3_Titles.Style.Alignment.WrapText = false;
                rangeD2D3_Titles.Style.Alignment.TextRotation = 0;
                rangeD2D3_Titles.Style.Alignment.Indent = 0;

                var colorRedC00000 = XLColor.FromArgb(0xC0, 0x00, 0x00);
                var colorBlack000000 = XLColor.FromArgb(0x00, 0x00, 0x00);
                var colorTeal009A93 = XLColor.FromArgb(0x00, 0x9A, 0x93);

                int[] rowsToColorRedAndSetHeight = { 1, 4 };
                foreach (int rowNum in rowsToColorRedAndSetHeight)
                {
                    worksheet.Row(rowNum).Style.Fill.BackgroundColor = colorRedC00000;
                    worksheet.Row(rowNum).Height = 7.5;
                }

                worksheet.Row(2).Style.Fill.BackgroundColor = colorTeal009A93;
                worksheet.Row(3).Style.Fill.BackgroundColor = colorTeal009A93;

                worksheet.Column(1).Style.Fill.BackgroundColor = colorBlack000000;
                worksheet.Column(2).Style.Fill.BackgroundColor = colorRedC00000;
                worksheet.Column(3).Style.Fill.SetBackgroundColor(XLColor.NoColor);

                var colD_Content_Style = worksheet.Column(labelColumn).Style;
                colD_Content_Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                colD_Content_Style.Alignment.TextRotation = 0;
                colD_Content_Style.Alignment.Indent = 0;
                colD_Content_Style.NumberFormat.Format = "@";
                colD_Content_Style.Alignment.WrapText = false;

                var colE_Content_Style = worksheet.Column(valueColumn).Style;
                colE_Content_Style.NumberFormat.Format = "@";
                colE_Content_Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                colE_Content_Style.Alignment.TextRotation = 0;

                worksheet.PageSetup.PrintAreas.Clear();
                int lastPrintRow = worksheet.LastRowUsed()?.RowNumber() ?? 69;
                worksheet.PageSetup.PrintAreas.Add($"A1:G{lastPrintRow}");
                worksheet.PageSetup.FitToPages(1, 1);
                worksheet.PageSetup.CenterVertically = true;
            }

            #region Footer Setup
            var hfRightLayout = worksheet.PageSetup.Footer.Right;
            hfRightLayout.Clear();
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
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
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
            cell.Style.Font.FontColor = XLColor.FromArgb(0, 154, 147);
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void ApplyHeading3Style_ClosedXML(IXLCell cell)
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
            cell.Style.Font.FontSize = 10;
            cell.Style.Font.FontColor = XLColor.Black;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            cell.Style.Alignment.WrapText = true;
        }
        #endregion
    }
}
