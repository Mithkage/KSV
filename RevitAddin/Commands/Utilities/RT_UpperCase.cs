using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions; // Ensure this namespace is included
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RTS.Commands.Support;

namespace RTS.Commands.Utilities
{
    /// <summary>
    /// Revit External Command to uppercase specified text parameters on Views and Sheets.
    /// This script iterates through View Names, Sheet Names, Sheet Numbers, and custom
    /// parameters like 'Title on sheet', 'Approved By', 'Designed By', and 'Checked By'.
    /// It includes an exception list for words like 'kg', 'kW/h', 'mm', and 'kW' which
    /// will retain their original casing even if the rest of the string is uppercased.
    /// </summary>
    /// <remarks>
    /// To use this script:
    /// 1. Create a new C# Class Library project in Visual Studio.
    /// 2. Add references to RevitAPI.dll and RevitAPIUI.dll (typically found in
    ///    C:\Program Files\Autodesk\Revit 2022\).
    /// 3. Copy this code into the RT_UpperCaseClass.cs file.
    /// 4. Build the project.
    /// 5. Create a .addin manifest file pointing to the compiled DLL and place it
    ///    in one of Revit's AddIns folders (e.g., C:\ProgramData\Autodesk\Revit\Addins\2022\).
    /// 6. The custom parameters ('Title on sheet', 'Approved By', 'Designed By', 'Checked By')
    ///    must exist as Project Parameters or Shared Parameters applied to View Sheets
    ///    in your Revit project for the script to find and modify them.
    /// </remarks>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RT_UpperCaseClass : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the application and document
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // List to store details of updated elements for user feedback
            List<string> updatedElements = new List<string>();

            // Define the exception list for words that should not be uppercased
            // These words will be re-inserted in their original casing after the
            // rest of the string has been converted to uppercase.
            // These exceptions will only apply to standalone whole words.
            List<string> exceptionWords = new List<string> { "kg", "kW/h", "mm", "kW" };

            // Start a new transaction for modifying the model
            using (Transaction trans = new Transaction(doc, "Uppercase View and Sheet Data"))
            {
                try
                {
                    trans.Start();

                    // --- Process Views ---
                    // Filter for all View elements in the document
                    FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
                    ICollection<Element> views = viewCollector.OfClass(typeof(View)).ToElements();

                    foreach (View view in views)
                    {
                        // Exclude title sheets, schedules, and un-named views which might cause issues
                        if (view.IsTemplate || view.ViewType == ViewType.Schedule || string.IsNullOrWhiteSpace(view.Name))
                        {
                            continue; // Skip templates, schedules, and views without names
                        }

                        // Get the current view name
                        string originalViewName = view.Name;
                        string processedViewName = ApplyUppercaseWithExceptions(originalViewName, exceptionWords);

                        // If the name has changed, update it
                        if (originalViewName != processedViewName)
                        {
                            try
                            {
                                view.Name = processedViewName;
                                updatedElements.Add($"View Name: '{originalViewName}' -> '{processedViewName}'");
                            }
                            catch (Exception ex)
                            {
                                // Handle cases where view name cannot be changed (e.g., duplicated names)
                                updatedElements.Add($"Error updating View Name '{originalViewName}': {ex.Message}");
                            }
                        }
                    }

                    // --- Process Sheets (ViewSheet elements) ---
                    // Filter for all ViewSheet elements in the document
                    FilteredElementCollector sheetCollector = new FilteredElementCollector(doc);
                    ICollection<Element> sheets = sheetCollector.OfClass(typeof(ViewSheet)).ToElements();

                    foreach (ViewSheet sheet in sheets)
                    {
                        // Get current values for comparison
                        string originalSheetName = sheet.Name;
                        string originalSheetNumber = sheet.SheetNumber;

                        // Uppercase Sheet Name
                        string processedSheetName = ApplyUppercaseWithExceptions(originalSheetName, exceptionWords);
                        if (originalSheetName != processedSheetName)
                        {
                            try
                            {
                                sheet.Name = processedSheetName;
                                updatedElements.Add($"Sheet Name: '{originalSheetName}' -> '{processedSheetName}' (Sheet Number: {originalSheetNumber})");
                            }
                            catch (Exception ex)
                            {
                                updatedElements.Add($"Error updating Sheet Name '{originalSheetName}' (Number: {originalSheetNumber}): {ex.Message}");
                            }
                        }

                        // Uppercase Sheet Number
                        string processedSheetNumber = ApplyUppercaseWithExceptions(originalSheetNumber, exceptionWords);
                        if (originalSheetNumber != processedSheetNumber)
                        {
                            try
                            {
                                sheet.SheetNumber = processedSheetNumber;
                                updatedElements.Add($"Sheet Number: '{originalSheetNumber}' -> '{processedSheetNumber}' (Sheet Name: {originalSheetName})");
                            }
                            catch (Exception ex)
                            {
                                updatedElements.Add($"Error updating Sheet Number '{originalSheetNumber}' (Name: {originalSheetName}): {ex.Message}");
                            }
                        }

                        // Process custom parameters on the sheet
                        ProcessSheetParameter(sheet, "Title on sheet", exceptionWords, updatedElements);
                        ProcessSheetParameter(sheet, "Approved By", exceptionWords, updatedElements);
                        ProcessSheetParameter(sheet, "Designed By", exceptionWords, updatedElements);
                        ProcessSheetParameter(sheet, "Checked By", exceptionWords, updatedElements);
                    }

                    // Commit the transaction if successful
                    trans.Commit();

                    // Show success message to the user
                    if (updatedElements.Any())
                    {
                        string successMessage = "Successfully uppercased elements:\n\n" + string.Join("\n", updatedElements);
                        CustomTaskDialog.Show("RT_UpperCase - Success", successMessage);
                    }
                    else
                    {
                        CustomTaskDialog.Show("RT_UpperCase - No Changes", "No eligible view or sheet parameters were found that required uppercasing based on the defined criteria and exception list, or no changes were necessary.");
                    }

                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    // If an error occurs, rollback the transaction
                    if (trans.GetStatus() == TransactionStatus.Started)
                    {
                        trans.RollBack();
                    }
                    message = "An error occurred during the uppercasing process: " + ex.Message;
                    CustomTaskDialog.Show("RT_UpperCase - Error", message);
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Applies uppercasing to a string,
        /// while preserving the casing of specified exception words, but only when they are standalone.
        /// </summary>
        /// <param name="originalString">The string to process.</param>
        /// <param name="exceptions">A list of words whose casing should be preserved (standalone).</param>
        /// <returns>The processed string.</returns>
        private string ApplyUppercaseWithExceptions(string originalString, List<string> exceptions)
        {
            if (string.IsNullOrWhiteSpace(originalString))
            {
                return originalString;
            }

            // Convert the entire string to uppercase first
            string uppercasedString = originalString.ToUpper();

            // Iterate through each exception word and replace its uppercased version
            // in the processed string with its original casing, only if it's a whole word.
            foreach (string exception in exceptions)
            {
                // Construct the regex pattern with word boundaries (\b).
                // Regex.Escape is used to handle special regex characters within the exception word (e.g., / in kW/h).
                // The pattern looks for the uppercased version of the exception word (e.g., "KG", "MM")
                // as a whole word (standalone), and replaces it with the original casing ("kg", "mm").
                // RegexOptions.IgnoreCase ensures it matches "KG", "Kg", "kG" etc.
                string pattern = @"\b" + Regex.Escape(exception.ToUpper()) + @"\b";

                uppercasedString = Regex.Replace(uppercasedString, pattern, exception, RegexOptions.IgnoreCase);
            }

            return uppercasedString;
        }

        /// <summary>
        /// Helper method to process a specific parameter on a ViewSheet.
        /// </summary>
        /// <param name="sheet">The ViewSheet element.</param>
        /// <param name="paramName">The name of the parameter to process.</param>
        /// <param name="exceptionWords">List of words to exclude from uppercasing (standalone).</param>
        /// <param name="updatedElements">List to log updates.</param>
        private void ProcessSheetParameter(
            ViewSheet sheet,
            string paramName,
            List<string> exceptionWords,
            List<string> updatedElements)
        {
            // Get the parameter by its name
            Parameter param = sheet.LookupParameter(paramName);

            // Check if the parameter exists and is a string type
            if (param != null && param.StorageType == StorageType.String)
            {
                string originalValue = param.AsString();

                // Only process if the parameter has a value
                if (!string.IsNullOrWhiteSpace(originalValue))
                {
                    string processedValue = ApplyUppercaseWithExceptions(originalValue, exceptionWords);

                    // If the value has changed, update it
                    if (originalValue != processedValue)
                    {
                        try
                        {
                            param.Set(processedValue);
                            updatedElements.Add($"Sheet '{sheet.SheetNumber} - {sheet.Name}' -> Parameter '{paramName}': '{originalValue}' -> '{processedValue}'");
                        }
                        catch (Exception ex)
                        {
                            updatedElements.Add($"Error updating Sheet '{sheet.SheetNumber} - {sheet.Name}' Parameter '{paramName}': {ex.Message}");
                        }
                    }
                }
            }
            // No else block needed here; if parameter not found or not a string, we simply don't process it.
            // No warning is logged to the user for missing parameters to keep the success message clean.
        }
    }
}
