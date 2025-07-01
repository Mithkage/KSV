//
// File: RTS.cs
//
// Description: Main application class for creating a custom Revit ribbon tab and buttons.
// Includes buttons for various PowerCAD related tools and RTS Setup.
//
// Author: ReTick Solutions
//
// Date: July 1, 2025 (Updated)
//
#region Namespaces
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging; // Required for BitmapImage
using Autodesk.Revit.DB; // Required for TaskDialog
#endregion

// Define a namespace for your application
namespace RTS
{
    public class App : IExternalApplication
    {
        // Path to this assembly.
        static string ExecutingAssemblyPath = Assembly.GetExecutingAssembly().Location;

        /// <summary>
        /// Helper method to load images from embedded resources.
        /// Ensure your .png files are in a folder named 'Resources' in your project
        /// and their 'Build Action' is set to 'Embedded Resource'.
        /// The resource name format is typically: DefaultNamespace.FolderName.FileName.Extension
        /// E.g., "RTS.Resources.Icon_PC_Launcher.png"
        /// </summary>
        /// <param name="imageName">The name of the image file (e.g., "Icon_PC_Launcher.png").</param>
        /// <returns>A BitmapImage for the icon, or null if not found or an error occurs.</returns>
        private BitmapImage GetEmbeddedPng(string imageName)
        {
            try
            {
                string resourceName = GetType().Namespace + ".Resources." + imageName;

                Assembly assembly = Assembly.GetExecutingAssembly();

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // Optionally log or show a TaskDialog during development to indicate missing resources
                        // TaskDialog.Show("Image Load Warning", $"Could not find embedded resource: {resourceName}");
                        return null;
                    }

                    BitmapImage bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = stream;
                    bmp.CacheOption = BitmapCacheOption.OnLoad; // Cache the image data
                    bmp.EndInit();
                    bmp.Freeze(); // Freeze the image to make it usable on any thread
                    return bmp;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Image Load Error", $"Failed to load image '{imageName}'. Error: {ex.Message}");
                return null;
            }
        }

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
                RibbonPanel rtsSetupPanel = application.CreateRibbonPanel(tabName, "RTS");

                // --- Load default and specific icons ---
                BitmapImage defaultIcon = GetEmbeddedPng("Icon.png"); // This is your new default

                // --- PCAD Tools Buttons ---

                // 1. PC SWB Exporter
                PushButtonData pbdPcSwbExporter = new PushButtonData(
                    "CmdPcSwbExporter", "Export\nSWB Data", ExecutingAssemblyPath, "PC_SWB_Exporter.PC_SWB_ExporterClass")
                {
                    ToolTip = "Exports data for PowerCAD SWB import.",
                    LargeImage = defaultIcon // No specific icon provided, using default
                };

                // 2. PC SWB Importer
                PushButtonData pbdPcSwbImporter = new PushButtonData(
                    "CmdPcSwbImporter", "Import\nSWB Data", ExecutingAssemblyPath, "PC_SWB_Importer.PC_SWB_ImporterClass")
                {
                    ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters.",
                    LargeImage = defaultIcon // No specific icon provided, using default
                };

                // 3. Import Cable Summary Data
                PushButtonData pbdPcCableImporter = new PushButtonData(
                    "CmdPcCableImporter", "Import Cable\nSummary", ExecutingAssemblyPath, "PC_Cable_Importer.PC_Cable_ImporterClass")
                {
                    ToolTip = "Imports Cable Summary data from PowerCAD into SLD.",
                    LargeImage = defaultIcon // No specific icon provided, using default
                };

                // 4. PC Clear Data
                PushButtonData pbdPcClearData = new PushButtonData(
                    "CmdPcClearData", "Clear PCAD\nData", ExecutingAssemblyPath, "PC_Clear_Data.PC_Clear_DataClass")
                {
                    ToolTip = "Clears specific PowerCAD-related parameters from Detail Items where PC_PowerCAD is 'Yes'.",
                    LargeImage = defaultIcon // No specific icon provided, using default
                };

                // 5. PC Generate MD Report
                PushButtonData pbdPcGenerateMd = new PushButtonData(
                    "CmdPcGenerateMd", "Generate\nMD Report", ExecutingAssemblyPath, "PC_Generate_MD.PC_Generate_MDClass")
                {
                    ToolTip = "Generates an Excel report with Cover Page, Authority, and Submains data.",
                    LargeImage = GetEmbeddedPng("PC_Generate_MD.png") ?? defaultIcon // Updated to PC_Generate_MD.png
                };

                // 6. PC_Extensible: Process & Save Cable Data
                PushButtonData pbdPcExtensible = new PushButtonData(
                    "CmdPcExtensible", "Process & Save\nCable Data", ExecutingAssemblyPath, "PC_Extensible.PC_ExtensibleClass")
                {
                    ToolTip = "Processes the Cleaned Cable Schedule and saves its data to project extensible storage.",
                    LargeImage = GetEmbeddedPng("PC_Extensible.png") ?? defaultIcon
                };

                // 7. PC_WireData: Update Electrical Wires (Updated Name)
                PushButtonData pbdPcWireData = new PushButtonData(
                    "CmdPcWireData", "Update Electrical\nWires", ExecutingAssemblyPath, "PC_WireData.PC_WireDataClass")
                {
                    ToolTip = "Reads cleaned cable data from extensible storage and updates electrical wires in the model.",
                    LargeImage = GetEmbeddedPng("PC_WireData.png") ?? defaultIcon
                };

                // --- Revit Tools Buttons ---

                // 1. RT Cable Lengths
                PushButtonData pbdRtCableLengths = new PushButtonData(
                    "CmdRTCableLengths", "Update Cable\nLengths", ExecutingAssemblyPath, "RTCableLengths.RTCableLengthsCommand")
                {
                    ToolTip = "Calculates and updates 'PC_Cable Length' on Detail Items based on 'PC_SWB To' and summed lengths of Conduits/Cable Trays.",
                    LargeImage = GetEmbeddedPng("RT_CableLengths.png") ?? defaultIcon
                };

                // 2. RT Panel Connect
                PushButtonData pbdRtPanelConnect = new PushButtonData(
                    "CmdRtPanelConnect", "Connect\nPanels", ExecutingAssemblyPath, "RT_PanelConnect.RT_PanelConnectClass")
                {
                    ToolTip = "Powers electrical panels by connecting them to their source panel based on a CSV file.",
                    LargeImage = defaultIcon // Updated to defaultIcon as no specific new icon was provided
                };

                // 3. RT Tray Occupancy
                PushButtonData pbdRtTrayOccupancy = new PushButtonData(
                    "CmdRtTrayOccupancy", "Process\nTray Data", ExecutingAssemblyPath, "RT_TrayOccupancy.RT_TrayOccupancyClass")
                {
                    ToolTip = "Parses a PowerCAD cable schedule to extract and clean data for export.",
                    LargeImage = GetEmbeddedPng("RT_TrayOccupancy.png") ?? defaultIcon
                };

                // 4. RT Tray ID
                PushButtonData pbdRtTrayId = new PushButtonData(
                    "CmdRtTrayId", "Update Tray\nIDs", ExecutingAssemblyPath, "RT_TrayID.RT_TrayIDClass")
                {
                    ToolTip = "Generates and updates unique IDs for cable tray elements.",
                    LargeImage = defaultIcon // Updated to defaultIcon as no specific new icon was provided
                };

                // 5. RT Tray Conduits
                PushButtonData pbdRtTrayConduits = new PushButtonData(
                    "CmdRtTrayConduits", "Create Tray\nConduits", ExecutingAssemblyPath, "RT_TrayConduits.RT_TrayConduitsClass")
                {
                    ToolTip = "Models conduits along cable trays based on cable data.",
                    LargeImage = GetEmbeddedPng("RT_TrayConduits.png") ?? defaultIcon
                };

                // 6. RT Uppercase
                PushButtonData pbdRtUppercase = new PushButtonData(
                    "CmdRtUppercase", "Uppercase\nText", ExecutingAssemblyPath, "RT_UpperCase.RT_UpperCaseClass")
                {
                    ToolTip = "Converts view names, sheet names, sheet numbers, and specific sheet parameters to uppercase, with exceptions.",
                    LargeImage = GetEmbeddedPng("RT_UpperCase.png") ?? defaultIcon
                };

                // 7. RT Wire Route (NEW BUTTON)
                PushButtonData pbdRtWireRoute = new PushButtonData(
                    "CmdRtWireRoute", "Wire\nRoute", ExecutingAssemblyPath, "RT_WireRoute.RT_WireRouteClass")
                {
                    ToolTip = "Routes electrical wires through conduits based on matching RTS_ID.",
                    LargeImage = GetEmbeddedPng("RT_WireRoute.png") ?? defaultIcon
                };


                // --- Misc Tools Buttons ---

                // 1. Import BB Cable Lengths
                PushButtonData pbdBbImport = new PushButtonData(
                    "CmdBbCableLengthImport", "Import\n BB Lengths", ExecutingAssemblyPath, "BB_Import.BB_CableLengthImport")
                {
                    ToolTip = "Imports Bluebeam cable measurements into SLD components.",
                    LargeImage = GetEmbeddedPng("BB_Import.png") ?? defaultIcon
                };

                // 2. MD Importer - Update Detail Item Loads from Excel
                PushButtonData pbdMdImporter = new PushButtonData(
                    "CmdMdExcelUpdater", "Update SWB\nLoads (Excel)", ExecutingAssemblyPath, "MD_Importer.MD_ImporterClass")
                {
                    ToolTip = "Updates 'PC_SWB Load' parameter for Detail Items from an Excel file ('TB_Submains' table) based on the 'PC_SWB To' parameter.",
                    LargeImage = GetEmbeddedPng("MD_Importer.png") ?? defaultIcon // Updated to MD_Importer.png
                };

                // --- RTS Setup Panel Buttons ---

                // 1. RTS Initiate Parameters
                PushButtonData pbdRtsInitiate = new PushButtonData(
                    "CmdRtsInitiate", "Initiate\nRTS Params", ExecutingAssemblyPath, "RTS_Initiate.RTS_InitiateClass")
                {
                    ToolTip = "Ensures required RTS Shared Parameters (like PC_PowerCAD) are added to the project.",
                    LargeImage = GetEmbeddedPng("RTS_Initiate.png") ?? defaultIcon
                };

                // 2. RTS Map Cables
                PushButtonData pbdRtsMapCables = new PushButtonData(
                    "CmdRtsMapCables", "Map\nCables", ExecutingAssemblyPath, "RTS_MapCables.RTS_MapCablesClass")
                {
                    ToolTip = "Maps RTS parameters with client parameters using a CSV mapping file.",
                    LargeImage = GetEmbeddedPng("RTS_MapCables.png") ?? defaultIcon
                };

                // 3. RTS Generate Schedules
                PushButtonData pbdRtsGenerateSchedules = new PushButtonData(
                    "CmdRtsGenerateSchedules", "Generate\nSchedules", ExecutingAssemblyPath, "RTS_Schedules.RTS_SchedulesClass")
                {
                    ToolTip = "Deletes and recreates a standard set of project schedules.",
                    LargeImage = GetEmbeddedPng("RTS_Schedules.png") ?? defaultIcon
                };

                // --- Add buttons to their respective panels ---
                pcadPanel.AddItem(pbdPcSwbExporter);
                pcadPanel.AddItem(pbdPcSwbImporter);
                pcadPanel.AddItem(pbdPcCableImporter);
                pcadPanel.AddItem(pbdPcClearData);
                pcadPanel.AddItem(pbdPcGenerateMd);
                pcadPanel.AddItem(pbdPcExtensible);
                pcadPanel.AddItem(pbdPcWireData);

                revitToolsPanel.AddItem(pbdRtCableLengths);
                revitToolsPanel.AddItem(pbdRtPanelConnect);
                revitToolsPanel.AddItem(pbdRtTrayOccupancy);
                revitToolsPanel.AddItem(pbdRtTrayId);
                revitToolsPanel.AddItem(pbdRtTrayConduits);
                revitToolsPanel.AddItem(pbdRtUppercase);
                revitToolsPanel.AddItem(pbdRtWireRoute); // Add the new Wire Route button

                miscPanel.AddItem(pbdBbImport);
                miscPanel.AddItem(pbdMdImporter);

                rtsSetupPanel.AddItem(pbdRtsInitiate);
                rtsSetupPanel.AddItem(pbdRtsMapCables);
                rtsSetupPanel.AddItem(pbdRtsGenerateSchedules);

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
