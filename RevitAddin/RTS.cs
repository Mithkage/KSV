//
// File: RTS.cs
//
// Namespace: RTS
//
// Class: App
//
// Function: This is the main application class for the RTS Revit add-in. It handles
//           the OnStartup and OnShutdown events, and is responsible for creating the
//           custom ribbon tab ("RTS") and all associated panels and buttons that
//           launch the various tools.
//
// Author: Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - July 30, 2025: Added buttons for ScheduleFloor and CastCeiling commands.
// - July 16, 2025: Added standard file header comment.
// - July 16, 2025: Added the 'Link Manager' button to the 'Revit Tools' panel.
//

#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
#endregion

// The main application class for the add-in.
namespace RTS
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
                // Create the ribbon panel during startup.
                CreateRibbonPanel(application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If anything goes wrong, show the detailed error and fail the startup.
                Autodesk.Revit.UI.TaskDialog.Show("RTS Add-in Load Failed", ex.ToString());
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

            // --- Create Ribbon Panels (Order of creation determines position) ---
            RibbonPanel pcadPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "PCAD Tools") ?? application.CreateRibbonPanel(tabName, "PCAD Tools");
            RibbonPanel revitToolsPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Revit Tools") ?? application.CreateRibbonPanel(tabName, "Revit Tools");
            RibbonPanel miscPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Misc Tools") ?? application.CreateRibbonPanel(tabName, "Misc Tools");
            RibbonPanel rtsSetupPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "RTS") ?? application.CreateRibbonPanel(tabName, "RTS");

            // --- Load default and specific icons ---
            BitmapImage defaultIcon = GetEmbeddedPng("Icon.png");

            // --- PCAD Tools Buttons ---

            // 1. PC SWB Exporter
            PushButtonData pbdPcSwbExporter = new PushButtonData("CmdPcSwbExporter", "Export\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.PC_SWB_ExporterClass")
            {
                ToolTip = "Exports data for PowerCAD SWB import.",
                LargeImage = defaultIcon
            };

            // 2. PC SWB Importer
            PushButtonData pbdPcSwbImporter = new PushButtonData("CmdPcSwbImporter", "Import\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.PC_SWB_ImporterClass")
            {
                ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters.",
                LargeImage = defaultIcon
            };

            // 3. Import Cable Summary Data
            PushButtonData pbdPcCableImporter = new PushButtonData("CmdPcCableImporter", "Import Cable\nSummary", ExecutingAssemblyPath, "RTS.Commands.PC_Cable_ImporterClass")
            {
                ToolTip = "Imports Cable Summary data from PowerCAD into SLD.",
                LargeImage = defaultIcon
            };

            // 4. PC Clear Data
            PushButtonData pbdPcClearData = new PushButtonData("CmdPcClearData", "Clear PCAD\nData", ExecutingAssemblyPath, "RTS.Commands.PC_Clear_DataClass")
            {
                ToolTip = "Clears specific PowerCAD-related parameters from Detail Items where PC_PowerCAD is 'Yes'.",
                LargeImage = defaultIcon
            };

            // 5. PC Generate MD Report
            PushButtonData pbdPcGenerateMd = new PushButtonData("CmdPcGenerateMd", "Generate\nMD Report", ExecutingAssemblyPath, "RTS.Commands.PC_Generate_MDClass")
            {
                ToolTip = "Generates an Excel report with Cover Page, Authority, and Submains data.",
                LargeImage = GetEmbeddedPng("PC_Generate_MD.png") ?? defaultIcon
            };

            // 6. PC_Extensible: Process & Save Cable Data
            PushButtonData pbdPcExtensible = new PushButtonData("CmdPcExtensible", "Process & Save\nCable Data", ExecutingAssemblyPath, "PC_Extensible.PC_ExtensibleClass")
            {
                ToolTip = "Processes the Cleaned Cable Schedule and saves its data to project extensible storage.",
                LargeImage = GetEmbeddedPng("PC_Extensible.png") ?? defaultIcon
            };

            // 7. PC_WireData: Update Electrical Wires
            PushButtonData pbdPcWireData = new PushButtonData("CmdPcWireData", "Update Electrical\nWires", ExecutingAssemblyPath, "RTS.Commands.PC_WireDataClass")
            {
                ToolTip = "Reads cleaned cable data from extensible storage and updates electrical wires in the model.",
                LargeImage = GetEmbeddedPng("PC_WireData.png") ?? defaultIcon
            };

            // 8. PC_Updater: Update CSV file
            PushButtonData pbdPcUpdater = new PushButtonData("CmdPcUpdater", "Update PCAD\nCSV", ExecutingAssemblyPath, "RTS.Commands.PC_UpdaterClass")
            {
                ToolTip = "Updates 'Cable Length' in a PowerCAD CSV export using lengths from Model Generated Data.",
                LargeImage = GetEmbeddedPng("PC_Updater.png") ?? defaultIcon
            };

            // --- Revit Tools Buttons ---

            // 1. RT Cable Lengths
            PushButtonData pbdRtCableLengths = new PushButtonData("CmdRTCableLengths", "Update Cable\nLengths", ExecutingAssemblyPath, "RTS.Commands.RTCableLengthsCommand")
            {
                ToolTip = "Calculates and updates 'PC_Cable Length' on Detail Items based on 'PC_SWB To' and summed lengths of Conduits/Cable Trays.",
                LargeImage = GetEmbeddedPng("RT_CableLengths.png") ?? defaultIcon
            };

            // 2. RT Panel Connect
            PushButtonData pbdRtPanelConnect = new PushButtonData("CmdRtPanelConnect", "Connect\nPanels", ExecutingAssemblyPath, "RTS.Commands.RT_PanelConnectClass")
            {
                ToolTip = "Powers electrical panels by connecting them to their source panel based on a CSV file.",
                LargeImage = defaultIcon
            };

            // 3. RT Tray Occupancy
            PushButtonData pbdRtTrayOccupancy = new PushButtonData("CmdRtTrayOccupancy", "Process\nTray Data", ExecutingAssemblyPath, "RTS.Commands.RT_TrayOccupancyClass")
            {
                ToolTip = "Parses a PowerCAD cable schedule to extract and clean data for export.",
                LargeImage = GetEmbeddedPng("RT_TrayOccupancy.png") ?? defaultIcon
            };

            // 4. RT Tray ID
            PushButtonData pbdRtTrayId = new PushButtonData("CmdRtTrayId", "Update Tray\nIDs", ExecutingAssemblyPath, "RTS.Commands.RT_TrayIDClass")
            {
                ToolTip = "Generates and updates unique IDs for cable tray elements.",
                LargeImage = defaultIcon
            };

            // 5. RT Tray Conduits
            PushButtonData pbdRtTrayConduits = new PushButtonData("CmdRtTrayConduits", "Create Tray\nConduits", ExecutingAssemblyPath, "RTS.Commands.RT_TrayConduitsClass")
            {
                ToolTip = "Models conduits along cable trays based on cable data.",
                LargeImage = GetEmbeddedPng("RT_TrayConduits.png") ?? defaultIcon
            };

            // 6. RT Uppercase
            PushButtonData pbdRtUppercase = new PushButtonData("CmdRtUppercase", "Uppercase\nText", ExecutingAssemblyPath, "RTS.Commands.RT_UpperCaseClass")
            {
                ToolTip = "Converts view names, sheet names, sheet numbers, and specific sheet parameters to uppercase, with exceptions.",
                LargeImage = GetEmbeddedPng("RT_UpperCase.png") ?? defaultIcon
            };

            // 7. RT Wire Route
            PushButtonData pbdRtWireRoute = new PushButtonData("CmdRtWireRoute", "Wire\nRoute", ExecutingAssemblyPath, "RTS.Commands.RT_WireRouteClass")
            {
                ToolTip = "Routes electrical wires through conduits based on matching RTS_ID.",
                LargeImage = GetEmbeddedPng("RT_WireRoute.png") ?? defaultIcon
            };

            // 8. RT_Isolate
            PushButtonData pbdRtIsolate = new PushButtonData("CmdRtIsolate", "Isolate by\nID/Cable", ExecutingAssemblyPath, "RTS.Commands.RT_IsolateClass")
            {
                ToolTip = "Isolates elements in the active view based on 'RTS_ID' or 'RTS_Cable_XX' parameter values.",
                LargeImage = GetEmbeddedPng("RT_Isolate.png") ?? defaultIcon
            };

            // 9. Link Manager
            PushButtonData pbdLinkManager = new PushButtonData("CmdLinkManager", "Link\nManager", ExecutingAssemblyPath, "RTS.Commands.LinkManagerCommand")
            {
                ToolTip = "Opens a window to manage Revit links and their associated metadata.",
                LargeImage = GetEmbeddedPng("LinkManager.png") ?? defaultIcon
            };

            // 10. Schedule Level
            PushButtonData pbdScheduleLevel = new PushButtonData("CmdScheduleLevel", "Schedule\nLevel", ExecutingAssemblyPath, "RTS.Commands.ScheduleFloorClass")
            {
                ToolTip = "Updates the schedule level of lighting fixtures based on the closest floor in a linked model.",
                LargeImage = defaultIcon // TODO: Create a specific icon
            };

            // 11. Cast to Ceiling/Slab
            PushButtonData pbdCastCeiling = new PushButtonData("CmdCastCeiling", "Cast to\nCeiling/Slab", ExecutingAssemblyPath, "RTS.Commands.CastCeilingClass")
            {
                ToolTip = "Moves MEP elements vertically to the nearest ceiling or slab from linked models.",
                LargeImage = defaultIcon // TODO: Create a specific icon
            };

            // 12. Viewport Alignment Tool
            PushButtonData pbdViewportAlignment = new PushButtonData(
                "CmdViewportAlignment",
                "Align\nViewports",
                ExecutingAssemblyPath,
                "RTS.Commands.ViewportAlignmentTool")
            {
                ToolTip = "Aligns the center of selected viewports to a reference viewport on the active sheet.",
                LargeImage = GetEmbeddedPng("ViewportAlignment.png") ?? defaultIcon // Use a specific icon if available
            };

            // --- Misc Tools Buttons ---

            // 1. Import BB Cable Lengths
            PushButtonData pbdBbImport = new PushButtonData("CmdBbCableLengthImport", "Import\n BB Lengths", ExecutingAssemblyPath, "RTS.Commands.BB_CableLengthImport")
            {
                ToolTip = "Imports Bluebeam cable measurements into SLD components.",
                LargeImage = GetEmbeddedPng("BB_Import.png") ?? defaultIcon
            };

            // 2. MD Importer - Update Detail Item Loads from Excel
            PushButtonData pbdMdImporter = new PushButtonData("CmdMdExcelUpdater", "Update SWB\nLoads (Excel)", ExecutingAssemblyPath, "RTS.Commands.MD_ImporterClass")
            {
                ToolTip = "Updates 'PC_SWB Load' parameter for Detail Items from an Excel file ('TB_Submains' table) based on the 'PC_SWB To' parameter.",
                LargeImage = GetEmbeddedPng("MD_Importer.png") ?? defaultIcon
            };

            // 3. Copy Relative (Door/Light)
            PushButtonData pbdCopyRelative = new PushButtonData(
                "CmdCopyRelative",
                "Copy\nRelative",
                ExecutingAssemblyPath,
                "RTS.Commands.CopyRelativeClass")
            {
                ToolTip = "Places light fixtures in the host model at locations relative to doors in a linked model, based on a template selection.",
                LargeImage = GetEmbeddedPng("CopyRelative.png") ?? defaultIcon // Use a specific icon if available
            };

            // --- RTS Setup Panel Buttons ---

            // 1. RTS Initiate Parameters
            PushButtonData pbdRtsInitiate = new PushButtonData("CmdRtsInitiate", "Initiate\nRTS Params", ExecutingAssemblyPath, "RTS.Commands.RTS_InitiateClass")
            {
                ToolTip = "Ensures required RTS Shared Parameters (like PC_PowerCAD) are added to the project.",
                LargeImage = GetEmbeddedPng("RTS_Initiate.png") ?? defaultIcon
            };

            // 2. RTS Map Cables
            PushButtonData pbdRtsMapCables = new PushButtonData("CmdRtsMapCables", "Map\nCables", ExecutingAssemblyPath, "RTS.Commands.RTS_MapCablesClass")
            {
                ToolTip = "Maps RTS parameters with client parameters using a CSV mapping file.",
                LargeImage = GetEmbeddedPng("RTS_MapCables.png") ?? defaultIcon
            };

            // 3. RTS Generate Schedules
            PushButtonData pbdRtsGenerateSchedules = new PushButtonData("CmdRtsGenerateSchedules", "Generate\nSchedules", ExecutingAssemblyPath, "RTS.Addin.Commands.RTS_Schedules.RTS_SchedulesClass")
            {
                ToolTip = "Deletes and recreates a standard set of project schedules.",
                LargeImage = GetEmbeddedPng("RTS_Schedules.png") ?? defaultIcon
            };

            // 4. RTS Reports
            PushButtonData pbdRtsReports = new PushButtonData("CmdRtsReports", "RTS\nReports", ExecutingAssemblyPath, "RTS.Addin.Commands.RTS_Reports.ShowReportsWindowCommand")
            {
                ToolTip = "Generates various reports from extensible storage data.",
                LargeImage = GetEmbeddedPng("RTS_Reports.png") ?? defaultIcon
            };

            // --- Add buttons to their respective panels ---
            pcadPanel.AddItem(pbdPcSwbExporter);
            pcadPanel.AddItem(pbdPcSwbImporter);
            pcadPanel.AddItem(pbdPcCableImporter);
            pcadPanel.AddItem(pbdPcClearData);
            pcadPanel.AddItem(pbdPcGenerateMd);
            pcadPanel.AddItem(pbdPcExtensible);
            pcadPanel.AddItem(pbdPcWireData);
            pcadPanel.AddItem(pbdPcUpdater);

            revitToolsPanel.AddItem(pbdRtCableLengths);
            revitToolsPanel.AddItem(pbdRtPanelConnect);
            revitToolsPanel.AddItem(pbdRtTrayOccupancy);
            revitToolsPanel.AddItem(pbdRtTrayId);
            revitToolsPanel.AddItem(pbdRtTrayConduits);
            revitToolsPanel.AddItem(pbdRtUppercase);
            revitToolsPanel.AddItem(pbdRtWireRoute);
            revitToolsPanel.AddItem(pbdRtIsolate);
            revitToolsPanel.AddItem(pbdLinkManager);
            revitToolsPanel.AddItem(pbdScheduleLevel);
            revitToolsPanel.AddItem(pbdCastCeiling);
            revitToolsPanel.AddItem(pbdViewportAlignment);

            miscPanel.AddItem(pbdBbImport);
            miscPanel.AddItem(pbdMdImporter);
            miscPanel.AddItem(pbdCopyRelative);

            rtsSetupPanel.AddItem(pbdRtsInitiate);
            rtsSetupPanel.AddItem(pbdRtsMapCables);
            rtsSetupPanel.AddItem(pbdRtsGenerateSchedules);
            rtsSetupPanel.AddItem(pbdRtsReports);
        }

        private BitmapImage GetEmbeddedPng(string imageName)
        {
            try
            {
                // Assuming "Resources" is still a direct subfolder of the assembly's root namespace (RTS).
                string resourceName = GetType().Namespace + ".Resources." + imageName;
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return null;
                    BitmapImage bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = stream;
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}

// This namespace now contains the command to launch your window.
// Ensure all other command classes (e.g., PC_SWB_ExporterClass) are defined in their respective namespaces.
// This specific namespace for RTS_Reports is assumed to be at the root level,
// or its path would also need to be updated if it moved under 'Commands'.
namespace RTS.Addin.Commands.RTS_Reports // Updated namespace
{
    [Transaction(TransactionMode.Manual)]
    public class ShowReportsWindowCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // This command now launches your existing ReportSelectionWindow
                // Ensure ReportSelectionWindow is in RTS.UI namespace
                var window = new RTS.UI.ReportSelectionWindow(commandData); // Updated namespace
                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
