//
// File:          RTS.cs
//
// Namespace:     RTS.App
//
// Class:         App
//
// Function:      This is the main application class for the Revit add-in. It handles the
//                OnStartup and OnShutdown events and is responsible for creating the
//                custom ribbon tab ("RTS") and all associated panels and buttons.
//
// Author:        Kyle Vorster (Modified by AI)
//
// Log:
// 2025-07-09:    Added System.Diagnostics namespace to resolve 'Debug' class errors.
// 2025-07-09:    Refactored ribbon panel creation to match the updated folder structure.
//                - Renamed panels to "PowerCAD", "Power Reticulation", "PDF Tools", and "Utilities".
//                - Updated all button command paths to reflect their new namespaces.
// 2025-07-09:    Corrected TaskDialog call to use the self-contained CustomTaskDialog.
// 2025-07-09:    Updated GetEmbeddedPng to correctly reference icons within the 'Icons' subfolder
//                of the 'Resources' directory, based on the Solution Explorer screenshot.
// 2025-07-09:    Temporarily commented out the CreateRibbonPanel call in OnStartup for debugging
//                TypeLoadException to isolate the issue.
//
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics; // FIX: Added for the Debug class
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using RTS.Commands.Support; // This using statement is important
#endregion

// The main application class for the add-in.
namespace RTS.App
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        // Path to this assembly.
        static string ExecutingAssemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Temporarily commented out for debugging TypeLoadException.
                // If the add-in loads successfully after this, the issue is within CreateRibbonPanel or its dependencies.
                // Create the ribbon panel during startup.
                // CreateRibbonPanel(application); 
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If anything goes wrong, show the detailed error and fail the startup.
                CustomTaskDialog.Show("RTS Add-in Load Failed", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Creates the ribbon tab and panels for the add-in.
        /// </summary>
        /// <param name="application">The UIControlledApplication instance.</param>
        private void CreateRibbonPanel(UIControlledApplication application)
        {
            string tabName = "RTS";

            // Create the ribbon tab. Handle the case where it might already exist.
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Tab already exists, this is fine.
            }

            // --- Create Ribbon Panels based on new folder structure ---
            RibbonPanel powerCadPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "PowerCAD") ?? application.CreateRibbonPanel(tabName, "PowerCAD");
            RibbonPanel powerReticulationPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Power Reticulation") ?? application.CreateRibbonPanel(tabName, "Power Reticulation");
            RibbonPanel pdfToolsPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "PDF Tools") ?? application.CreateRibbonPanel(tabName, "PDF Tools");
            RibbonPanel utilitiesPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Utilities") ?? application.CreateRibbonPanel(tabName, "Utilities");


            // --- Load default and specific icons ---
            BitmapImage defaultIcon = GetEmbeddedPng("Icon.png");

            // --- PowerCAD Panel Buttons ---
            AddButton(powerCadPanel, "CmdPcSwbExporter", "Export\nSWB Data", "RTS.Commands.PowerCAD.PC_SWB_ExporterClass", "Exports data for PowerCAD SWB import.", defaultIcon);
            AddButton(powerCadPanel, "CmdPcSwbImporter", "Import\nSWB Data", "RTS.Commands.PowerCAD.PC_SWB_ImporterClass", "Imports cable data from a PowerCAD SWB CSV export file.", defaultIcon);
            AddButton(powerCadPanel, "CmdPcCableImporter", "Import Cable\nSummary", "RTS.Commands.PowerCAD.PC_Cable_ImporterClass", "Imports Cable Summary data from PowerCAD into SLD.", defaultIcon);
            AddButton(powerCadPanel, "CmdPcClearData", "Clear PCAD\nData", "RTS.Commands.PowerCAD.PC_Clear_DataClass", "Clears specific PowerCAD-related parameters.", defaultIcon);
            AddButton(powerCadPanel, "CmdPcGenerateMd", "Generate\nMD Report", "RTS.Commands.PowerCAD.PC_Generate_MDClass", "Generates an Excel report with Cover Page, Authority, and Submains data.", GetEmbeddedPng("PC_Generate_MD.png") ?? defaultIcon);
            AddButton(powerCadPanel, "CmdPcExtensible", "Process & Save\nCable Data", "RTS.Commands.PowerCAD.PC_ExtensibleClass", "Processes the Cleaned Cable Schedule and saves its data to project extensible storage.", GetEmbeddedPng("PC_Extensible.png") ?? defaultIcon);
            AddButton(powerCadPanel, "CmdPcWireData", "Update Electrical\nWires", "RTS.Commands.PowerCAD.PC_WireDataClass", "Reads cleaned cable data from extensible storage and updates electrical wires.", GetEmbeddedPng("PC_WireData.png") ?? defaultIcon);
            AddButton(powerCadPanel, "CmdPcUpdater", "Update PCAD\nCSV", "RTS.Commands.PowerCAD.PC_UpdaterClass", "Updates 'Cable Length' in a PowerCAD CSV export.", GetEmbeddedPng("PC_Updater.png") ?? defaultIcon);
            AddButton(powerCadPanel, "CmdMdExcelUpdater", "Update SWB\nLoads (Excel)", "RTS.Commands.PowerCAD.MD_ImporterClass", "Updates 'PC_SWB Load' parameter for Detail Items from an Excel file.", GetEmbeddedPng("MD_Importer.png") ?? defaultIcon);

            // --- Power Reticulation Panel Buttons ---
            AddButton(powerReticulationPanel, "CmdRTCableLengths", "Update Cable\nLengths", "RTS.Commands.PowerReticulation.RTCableLengthsCommand", "Calculates and updates 'PC_Cable Length' on Detail Items.", GetEmbeddedPng("RT_CableLengths.png") ?? defaultIcon);
            AddButton(powerReticulationPanel, "CmdRtPanelConnect", "Connect\nPanels", "RTS.Commands.PowerReticulation.RT_PanelConnectClass", "Powers electrical panels by connecting them to their source panel based on a CSV file.", defaultIcon);
            AddButton(powerReticulationPanel, "CmdRtTrayOccupancy", "Process\nTray Data", "RTS.Commands.PowerReticulation.RT_TrayOccupancyClass", "Parses a PowerCAD cable schedule to extract and clean data for export.", GetEmbeddedPng("RT_TrayOccupancy.png") ?? defaultIcon);
            AddButton(powerReticulationPanel, "CmdRtTrayId", "Update Tray\nIDs", "RTS.Commands.PowerReticulation.RT_TrayIDClass", "Generates and updates unique IDs for cable tray elements.", defaultIcon);
            AddButton(powerReticulationPanel, "CmdRtTrayConduits", "Create Tray\nConduits", "RTS.Commands.PowerReticulation.RT_TrayConduitsClass", "Models conduits along cable trays based on cable data.", GetEmbeddedPng("RT_TrayConduits.png") ?? defaultIcon);
            AddButton(powerReticulationPanel, "CmdRtWireRoute", "Wire\nRoute", "RTS.Commands.PowerReticulation.RT_WireRouteClass", "Routes electrical wires through conduits based on matching RTS_ID.", GetEmbeddedPng("RT_WireRoute.png") ?? defaultIcon);

            // --- PDF Tools Panel Buttons ---
            AddButton(pdfToolsPanel, "CmdBbCableLengthImport", "Import\nBB Lengths", "RTS.Commands.PDFTools.BB_CableLengthImport", "Imports Bluebeam cable measurements into SLD components.", GetEmbeddedPng("BB_Import.png") ?? defaultIcon);

            // --- Utilities Panel Buttons ---
            AddButton(utilitiesPanel, "CmdRtsInitiate", "Initiate\nRTS Params", "RTS.Commands.Utilities.RTS_InitiateClass", "Ensures required RTS Shared Parameters are added to the project.", GetEmbeddedPng("RTS_Initiate.png") ?? defaultIcon);
            AddButton(utilitiesPanel, "CmdRtsMapCables", "Map\nCables", "RTS.Commands.Utilities.RTS_MapCablesClass", "Maps RTS parameters with client parameters using a CSV mapping file.", GetEmbeddedPng("RTS_MapCables.png") ?? defaultIcon);
            AddButton(utilitiesPanel, "CmdRtsGenerateSchedules", "Generate\nSchedules", "RTS.Commands.Utilities.RTS_SchedulesClass", "Deletes and recreates a standard set of project schedules.", GetEmbeddedPng("RTS_Schedules.png") ?? defaultIcon);
            AddButton(utilitiesPanel, "CmdRtsReports", "RTS\nReports", "RTS.Commands.Utilities.ShowReportsWindowCommand", "Generates various reports from extensible storage data.", GetEmbeddedPng("RTS_Reports.png") ?? defaultIcon);
            AddButton(utilitiesPanel, "CmdRtUppercase", "Uppercase\nText", "RTS.Commands.Utilities.RT_UpperCaseClass", "Converts view names, sheet names, and sheet numbers to uppercase.", GetEmbeddedPng("RT_UpperCase.png") ?? defaultIcon);
            AddButton(utilitiesPanel, "CmdRtIsolate", "Isolate by\nID/Cable", "RTS.Commands.Utilities.RT_IsolateClass", "Isolates elements in the active view based on 'RTS_ID' or 'RTS_Cable_XX' parameter values.", GetEmbeddedPng("RT_Isolate.png") ?? defaultIcon);
        }

        /// <summary>
        /// Helper method to create and add a push button to a ribbon panel.
        /// </summary>
        private void AddButton(RibbonPanel panel, string name, string text, string className, string toolTip, BitmapImage image)
        {
            PushButtonData buttonData = new PushButtonData(name, text, ExecutingAssemblyPath, className)
            {
                ToolTip = toolTip,
                LargeImage = image
            };
            panel.AddItem(buttonData);
        }

        /// <summary>
        /// Helper method to get an embedded PNG image from the assembly resources.
        /// </summary>
        private BitmapImage GetEmbeddedPng(string imageName)
        {
            try
            {
                // Note: The resource path might need adjustment if your resources are in a subfolder within the project.
                // Assuming resources are in a "Resources" folder and marked as "Embedded Resource".
                // Corrected resource path to include the 'Icons' subfolder.
                string resourceName = "RTS.Resources.Icons." + imageName;
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // Log or debug that the resource was not found.
                        Debug.WriteLine($"Embedded resource not found: {resourceName}");
                        return null;
                    }
                    BitmapImage bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = stream;
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading embedded image '{imageName}': {ex.Message}");
                return null;
            }
        }
    }
}

// NOTE: The `RTS_Reports` namespace and `ShowReportsWindowCommand` class have been moved
// into the `RTS.Commands.Utilities` namespace in the corresponding file (`RTS_Reports.cs`).
// This App.cs file should only contain the IExternalApplication implementation.
