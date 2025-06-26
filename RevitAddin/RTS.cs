//
// File: RTS.cs
//
// Description: Main application class for creating a custom Revit ribbon tab and buttons.
// Includes buttons for various PowerCAD related tools and RTS Setup.
//
// Author: ReTick Solutions
//
// Date: June 20, 2024
//
#region Namespaces
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
#endregion

// Define a namespace for your application
namespace RTS
{
    public class App : IExternalApplication
    {
        // Path to this assembly.
        static string ExecutingAssemblyPath = Assembly.GetExecutingAssembly().Location;
        // Helper to get the directory of the assembly, useful for finding resources.
        static string ExecutingAssemblyDirectory = Path.GetDirectoryName(ExecutingAssemblyPath);

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create a custom ribbon tab
                string tabName = "RTS";
                application.CreateRibbonTab(tabName);

                // --- Create Ribbon Panels (Order of creation determines position) ---
                RibbonPanel pcadPanel = application.CreateRibbonPanel(tabName, "PCAD Tools");
                RibbonPanel revitToolsPanel = application.CreateRibbonPanel(tabName, "Revit Tools");
                RibbonPanel miscPanel = application.CreateRibbonPanel(tabName, "Misc Tools");
                RibbonPanel rtsSetupPanel = application.CreateRibbonPanel(tabName, "RTS"); // New panel, placed on the far right


                // --- Define PushButtonData for each command ---
                string genericIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_PC_Launcher.png");

                // --- PCAD Tools Buttons ---

                // 1. PC SWB Exporter
                PushButtonData pbdPcSwbExporter = new PushButtonData(
                    "CmdPcSwbExporter", "Export\nSWB Data", ExecutingAssemblyPath, "PC_SWB_Exporter.PC_SWB_ExporterClass")
                {
                    ToolTip = "Exports data for PowerCAD SWB import."
                };
                if (File.Exists(genericIconPath)) pbdPcSwbExporter.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 2. PC SWB Importer
                PushButtonData pbdPcSwbImporter = new PushButtonData(
                    "CmdPcSwbImporter", "Import\nSWB Data", ExecutingAssemblyPath, "PC_SWB_Importer.PC_SWB_ImporterClass")
                {
                    ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters."
                };
                if (File.Exists(genericIconPath)) pbdPcSwbImporter.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 3. Import Cable Summary Data
                PushButtonData pbdPcCableImporter = new PushButtonData(
                    "CmdPcCableImporter", "Import Cable\nSummary", ExecutingAssemblyPath, "PC_Cable_Importer.PC_Cable_ImporterClass")
                {
                    ToolTip = "Imports Cable Summary data from PowerCAD into SLD."
                };
                if (File.Exists(genericIconPath)) pbdPcCableImporter.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 4. PC Clear Data
                PushButtonData pbdPcClearData = new PushButtonData(
                    "CmdPcClearData", "Clear PCAD\nData", ExecutingAssemblyPath, "PC_Clear_Data.PC_Clear_DataClass")
                {
                    ToolTip = "Clears specific PowerCAD-related parameters from Detail Items where PC_PowerCAD is 'Yes'."
                };
                if (File.Exists(genericIconPath)) pbdPcClearData.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 5. PC Generate MD Report
                PushButtonData pbdPcGenerateMd = new PushButtonData(
                    "CmdPcGenerateMd", "Generate\nMD Report", ExecutingAssemblyPath, "PC_Generate_MD.PC_Generate_MDClass")
                {
                    ToolTip = "Generates an Excel report with Cover Page, Authority, and Submains data."
                };
                string mdReportIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_MD_Report.png");
                pbdPcGenerateMd.LargeImage = new BitmapImage(new Uri(File.Exists(mdReportIconPath) ? mdReportIconPath : genericIconPath));


                // --- Revit Tools Buttons ---

                // 1. RT Cable Lengths
                PushButtonData pbdRtCableLengths = new PushButtonData(
                    "CmdRTCableLengths", "Update Cable\nLengths", ExecutingAssemblyPath, "RTCableLengths.RTCableLengthsCommand")
                {
                    ToolTip = "Calculates and updates 'PC_Cable Length' on Detail Items based on 'PC_SWB To' and summed lengths of Conduits/Cable Trays."
                };
                if (File.Exists(genericIconPath)) pbdRtCableLengths.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 2. RT Panel Connect
                PushButtonData pbdRtPanelConnect = new PushButtonData(
                    "CmdRtPanelConnect", "Connect\nPanels", ExecutingAssemblyPath, "RT_PanelConnect.RT_PanelConnectClass")
                {
                    ToolTip = "Powers electrical panels by connecting them to their source panel based on a CSV file."
                };
                string rtPanelConnectIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Panel_Connect.png");
                pbdRtPanelConnect.LargeImage = new BitmapImage(new Uri(File.Exists(rtPanelConnectIconPath) ? rtPanelConnectIconPath : genericIconPath));

                // 3. RT Tray Occupancy
                PushButtonData pbdRtTrayOccupancy = new PushButtonData(
                    "CmdRtTrayOccupancy", "Process\nTray Data", ExecutingAssemblyPath, "RT_TrayOccupancy.RT_TrayOccupancyClass")
                {
                    ToolTip = "Parses a PowerCAD cable schedule to extract and clean data for export."
                };
                if (File.Exists(genericIconPath)) pbdRtTrayOccupancy.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 4. RT Tray ID - NEW BUTTON ADDED HERE
                PushButtonData pbdRtTrayId = new PushButtonData(
                    "CmdRtTrayId", "Update Tray\nIDs", ExecutingAssemblyPath, "RT_TrayID.RT_TrayIDClass") // Ensure this matches your RT_TrayID class namespace and name
                {
                    ToolTip = "Generates and updates unique IDs for cable tray elements."
                };
                // You can specify a custom icon for RT_TrayID or reuse genericIconPath
                string rtTrayIdIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_TrayID.png"); // Example custom icon path
                pbdRtTrayId.LargeImage = new BitmapImage(new Uri(File.Exists(rtTrayIdIconPath) ? rtTrayIdIconPath : genericIconPath));

                // 5. RT Tray Conduits - NEW BUTTON ADDED HERE
                PushButtonData pbdRtTrayConduits = new PushButtonData(
                    "CmdRtTrayConduits", "Create Tray\nConduits", ExecutingAssemblyPath, "RT_TrayConduits.RT_TrayConduitsClass") // Match namespace and class name
                {
                    ToolTip = "Models conduits along cable trays based on cable data."
                };
                // Example custom icon path, or reuse genericIconPath
                string rtTrayConduitsIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Conduit.png"); // Assuming you might have a conduit-related icon
                pbdRtTrayConduits.LargeImage = new BitmapImage(new Uri(File.Exists(rtTrayConduitsIconPath) ? rtTrayConduitsIconPath : genericIconPath));


                // --- Misc Tools Buttons ---

                // 1. Import BB Cable Lengths
                PushButtonData pbdBbImport = new PushButtonData(
                    "CmdBbCableLengthImport", "Import\n BB Lengths", ExecutingAssemblyPath, "BB_Import.BB_CableLengthImport")
                {
                    ToolTip = "Imports Bluebeam cable measurements into SLD components."
                };
                if (File.Exists(genericIconPath)) pbdBbImport.LargeImage = new BitmapImage(new Uri(genericIconPath));

                // 2. MD Importer - Update Detail Item Loads from Excel
                PushButtonData pbdMdImporter = new PushButtonData(
                    "CmdMdExcelUpdater", "Update SWB\nLoads (Excel)", ExecutingAssemblyPath, "MD_Importer.MD_ImporterClass")
                {
                    ToolTip = "Updates 'PC_SWB Load' parameter for Detail Items from an Excel file ('TB_Submains' table) based on the 'PC_SWB To' parameter."
                };
                string mdImporterIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Excel_Import.png");
                pbdMdImporter.LargeImage = new BitmapImage(new Uri(File.Exists(mdImporterIconPath) ? mdImporterIconPath : genericIconPath));


                // --- RTS Setup Panel Buttons ---

                // 1. RTS Initiate Parameters
                PushButtonData pbdRtsInitiate = new PushButtonData(
                    "CmdRtsInitiate", "Initiate\nRTS Params", ExecutingAssemblyPath, "RTS_Initiate.RTS_InitiateClass")
                {
                    ToolTip = "Ensures required RTS Shared Parameters (like PC_PowerCAD) are added to the project."
                };
                string rtsInitiateIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Setup.png");
                pbdRtsInitiate.LargeImage = new BitmapImage(new Uri(File.Exists(rtsInitiateIconPath) ? rtsInitiateIconPath : genericIconPath));

                // 2. RTS Map Cables
                PushButtonData pbdRtsMapCables = new PushButtonData(
                    "CmdRtsMapCables", "Map\nCables", ExecutingAssemblyPath, "RTS_MapCables.RTS_MapCablesClass")
                {
                    ToolTip = "Maps RTS parameters with client parameters using a CSV mapping file."
                };
                string rtsMapCablesIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Map_Cables.png");
                pbdRtsMapCables.LargeImage = new BitmapImage(new Uri(File.Exists(rtsMapCablesIconPath) ? rtsMapCablesIconPath : genericIconPath));

                // 3. RTS Generate Schedules
                PushButtonData pbdRtsGenerateSchedules = new PushButtonData(
                    "CmdRtsGenerateSchedules", "Generate\nSchedules", ExecutingAssemblyPath, "RTS_Schedules.RTS_SchedulesClass")
                {
                    ToolTip = "Deletes and recreates a standard set of project schedules."
                };
                // Reusing the setup icon for consistency in the panel.
                string rtsSchedulesIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_Setup.png");
                pbdRtsGenerateSchedules.LargeImage = new BitmapImage(new Uri(File.Exists(rtsSchedulesIconPath) ? rtsSchedulesIconPath : genericIconPath));


                // --- Add buttons to their respective panels ---
                pcadPanel.AddItem(pbdPcSwbExporter);
                pcadPanel.AddItem(pbdPcSwbImporter);
                pcadPanel.AddItem(pbdPcCableImporter);
                pcadPanel.AddItem(pbdPcClearData);
                pcadPanel.AddItem(pbdPcGenerateMd);

                revitToolsPanel.AddItem(pbdRtCableLengths);
                revitToolsPanel.AddItem(pbdRtPanelConnect);
                revitToolsPanel.AddItem(pbdRtTrayOccupancy);
                revitToolsPanel.AddItem(pbdRtTrayId); // ADDED THE NEW BUTTON TO THE REVIT TOOLS PANEL
                revitToolsPanel.AddItem(pbdRtTrayConduits); // ADDED THE NEW RT_TRAYCONDUITS BUTTON HERE

                miscPanel.AddItem(pbdBbImport);
                miscPanel.AddItem(pbdMdImporter);

                rtsSetupPanel.AddItem(pbdRtsInitiate);
                rtsSetupPanel.AddItem(pbdRtsMapCables);
                rtsSetupPanel.AddItem(pbdRtsGenerateSchedules); // Added the new button to the RTS panel.

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RTS Add-In Error",
                    "Failed to initialize ReTick Solutions ribbon items.\n" +
                    "Error: " + ex.Message +
                    (ex.InnerException != null ? "\nInner Exception: " + ex.InnerException.Message : "") +
                    "\nStackTrace: " + ex.StackTrace);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
