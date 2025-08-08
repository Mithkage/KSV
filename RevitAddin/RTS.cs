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
// - August 15, 2025: Updated split button names and icons to use generic category representations.
// - August 14, 2025: Updated all button tooltips to better describe functionality and relationship to workflow.
// - August 12, 2025: Major reorganization of ribbon panels into functional groups: "Project Setup", "Data Exchange", "MEP Tools", and "Document Tools".
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
            // New functional panels based on workflows
            RibbonPanel projectSetupPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Project Setup") 
                ?? application.CreateRibbonPanel(tabName, "Project Setup");
                
            RibbonPanel dataExchangePanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Data Exchange") 
                ?? application.CreateRibbonPanel(tabName, "Data Exchange");
                
            RibbonPanel mepToolsPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "MEP Tools") 
                ?? application.CreateRibbonPanel(tabName, "MEP Tools");
                
            RibbonPanel documentToolsPanel = application.GetRibbonPanels(tabName).FirstOrDefault(p => p.Name == "Document Tools") 
                ?? application.CreateRibbonPanel(tabName, "Document Tools");

            // --- Load default and specific icons ---
            BitmapImage defaultIcon = GetEmbeddedPng("Icon.png");

            // --- Load category icons for split buttons ---
            BitmapImage parameterSetupIcon = GetEmbeddedPng("ParameterSetup_Icon.png") ?? defaultIcon;
            BitmapImage importDataIcon = GetEmbeddedPng("ImportData_Icon.png") ?? defaultIcon;
            BitmapImage exportDataIcon = GetEmbeddedPng("ExportData_Icon.png") ?? defaultIcon;
            BitmapImage electricalToolsIcon = GetEmbeddedPng("ElectricalTools_Icon.png") ?? defaultIcon;
            BitmapImage cableTrayToolsIcon = GetEmbeddedPng("CableTrayTools_Icon.png") ?? defaultIcon;
            BitmapImage positioningToolsIcon = GetEmbeddedPng("PositioningTools_Icon.png") ?? defaultIcon;
            BitmapImage viewToolsIcon = GetEmbeddedPng("ViewTools_Icon.png") ?? defaultIcon;

            // --- Create All Button Definitions ---

            // Project Setup - Parameter Setup
            PushButtonData pbdRtsInitiate = new PushButtonData("CmdRtsInitiate", "Initialize\nParameters", ExecutingAssemblyPath, "RTS.Commands.RTS_InitiateClass")
            {
                ToolTip = "Sets up required shared parameters needed for RTS tools. Must be run before using other tools.",
                LargeImage = GetEmbeddedPng("RTS_Initiate.png") ?? defaultIcon
            };

            PushButtonData pbdRtsMapCables = new PushButtonData("CmdRtsMapCables", "Map\nParameters", ExecutingAssemblyPath, "RTS.Commands.RTS_MapCablesClass")
            {
                ToolTip = "Maps standard RTS parameters to custom project parameters using a CSV file.",
                LargeImage = GetEmbeddedPng("RTS_MapCables.png") ?? defaultIcon
            };

            PushButtonData pbdPcClearData = new PushButtonData("CmdPcClearData", "Clear\nData", ExecutingAssemblyPath, "RTS.Commands.PC_Clear_DataClass")
            {
                ToolTip = "Clears PowerCAD-related parameters from elements while keeping core identifiers intact.",
                LargeImage = defaultIcon
            };

            // Project Setup - Documentation
            PushButtonData pbdRtsGenerateSchedules = new PushButtonData("CmdRtsGenerateSchedules", "Generate\nSchedules", ExecutingAssemblyPath, "RTS.Addin.Commands.RTS_Schedules.RTS_SchedulesClass")
            {
                ToolTip = "Creates or refreshes a standard set of electrical schedules based on RTS parameters.",
                LargeImage = GetEmbeddedPng("RTS_Schedules.png") ?? defaultIcon
            };

            PushButtonData pbdRtsReports = new PushButtonData("CmdRtsReports", "Generate\nReports", ExecutingAssemblyPath, "RTS.Addin.Commands.RTS_Reports.ShowReportsWindowCommand")
            {
                ToolTip = "Creates detailed reports from project data for documentation and coordination.",
                LargeImage = GetEmbeddedPng("RTS_Reports.png") ?? defaultIcon
            };

            // Data Exchange - Import Tools
            PushButtonData pbdPcSwbImporter = new PushButtonData("CmdPcSwbImporter", "Import\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.PC_SWB_ImporterClass")
            {
                ToolTip = "Imports switchboard and cable data from PowerCAD CSV exports into Revit.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdPcCableImporter = new PushButtonData("CmdPcCableImporter", "Import Cable\nSummary", ExecutingAssemblyPath, "RTS.Commands.PC_Cable_ImporterClass")
            {
                ToolTip = "Imports cable information from PowerCAD Cable Summary documents into Single Line Diagrams.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdBbImport = new PushButtonData("CmdBbCableLengthImport", "Import\nBluebeam", ExecutingAssemblyPath, "RTS.Commands.BB_CableLengthImport")
            {
                ToolTip = "Imports cable length measurements from Bluebeam markups into Revit elements.",
                LargeImage = GetEmbeddedPng("BB_Import.png") ?? defaultIcon
            };

            PushButtonData pbdMdImporter = new PushButtonData("CmdMdExcelUpdater", "Update SWB\nLoads", ExecutingAssemblyPath, "RTS.Commands.MD_ImporterClass")
            {
                ToolTip = "Updates switchboard load parameters from Excel spreadsheets using TB_Submains table.",
                LargeImage = GetEmbeddedPng("MD_Importer.png") ?? defaultIcon
            };

            // Data Exchange - Export Tools
            PushButtonData pbdPcSwbExporter = new PushButtonData("CmdPcSwbExporter", "Export\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.PC_SWB_ExporterClass")
            {
                ToolTip = "Exports switchboard and cable data for use in PowerCAD or other tools.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdPcGenerateMd = new PushButtonData("CmdPcGenerateMd", "Generate MD\nReport", ExecutingAssemblyPath, "RTS.Commands.PC_Generate_MDClass")
            {
                ToolTip = "Creates a Maximum Demand Excel report with Cover Page, Authority, and Submains details.",
                LargeImage = GetEmbeddedPng("PC_Generate_MD.png") ?? defaultIcon
            };

            // Data Exchange - Data Management
            PushButtonData pbdPcExtensible = new PushButtonData("CmdPcExtensible", "Process & Store\nCable Data", ExecutingAssemblyPath, "PC_Extensible.PC_ExtensibleClass")
            {
                ToolTip = "Processes cable schedules and stores data in project extensible storage for use by other tools.",
                LargeImage = GetEmbeddedPng("PC_Extensible.png") ?? defaultIcon
            };

            PushButtonData pbdPcUpdater = new PushButtonData("CmdPcUpdater", "Update PCAD\nCSV", ExecutingAssemblyPath, "RTS.Commands.PC_UpdaterClass")
            {
                ToolTip = "Updates cable lengths in PowerCAD CSV exports using data from the Revit model.",
                LargeImage = GetEmbeddedPng("PC_Updater.png") ?? defaultIcon
            };

            // MEP Tools - Electrical
            PushButtonData pbdPcWireData = new PushButtonData("CmdPcWireData", "Update\nWires", ExecutingAssemblyPath, "RTS.Commands.PC_WireDataClass")
            {
                ToolTip = "Creates or updates electrical wires in the Revit model based on cable data.",
                LargeImage = GetEmbeddedPng("PC_WireData.png") ?? defaultIcon
            };

            PushButtonData pbdRtWireRoute = new PushButtonData("CmdRtWireRoute", "Route\nWires", ExecutingAssemblyPath, "RTS.Commands.RT_WireRouteClass")
            {
                ToolTip = "Automatically routes electrical wires through conduits based on matching RTS_ID parameters.",
                LargeImage = GetEmbeddedPng("RT_WireRoute.png") ?? defaultIcon
            };

            PushButtonData pbdRtCableLengths = new PushButtonData("CmdRTCableLengths", "Update Cable\nLengths", ExecutingAssemblyPath, "RTS.Commands.RTCableLengthsCommand")
            {
                ToolTip = "Calculates and updates cable lengths based on conduit/tray paths and connections.",
                LargeImage = GetEmbeddedPng("RT_CableLengths.png") ?? defaultIcon
            };

            PushButtonData pbdRtPanelConnect = new PushButtonData("CmdRtPanelConnect", "Connect\nPanels", ExecutingAssemblyPath, "RTS.Commands.RT_PanelConnectClass")
            {
                ToolTip = "Creates power connections between electrical panels based on source/destination data.",
                LargeImage = defaultIcon
            };

            // MEP Tools - Cable Tray
            PushButtonData pbdRtTrayOccupancy = new PushButtonData("CmdRtTrayOccupancy", "Process\nTray Data", ExecutingAssemblyPath, "RTS.Commands.RT_TrayOccupancyClass")
            {
                ToolTip = "Analyzes cable tray requirements based on cable schedule data and routing.",
                LargeImage = GetEmbeddedPng("RT_TrayOccupancy.png") ?? defaultIcon
            };

            PushButtonData pbdRtTrayId = new PushButtonData("CmdRtTrayId", "Update Tray\nIDs", ExecutingAssemblyPath, "RTS.Commands.RT_TrayIDClass")
            {
                ToolTip = "Assigns unique identifiers to cable tray elements for coordination and scheduling.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdRtTrayConduits = new PushButtonData("CmdRtTrayConduits", "Create Tray\nConduits", ExecutingAssemblyPath, "RTS.Commands.RT_TrayConduitsClass")
            {
                ToolTip = "Creates conduit runs along cable trays based on cable requirements and routes.",
                LargeImage = GetEmbeddedPng("RT_TrayConduits.png") ?? defaultIcon
            };

            // MEP Tools - Element Positioning
            PushButtonData pbdCopyRelative = new PushButtonData("CmdCopyRelative", "Copy\nRelative", ExecutingAssemblyPath, "RTS.Commands.CopyRelativeClass")
            {
                ToolTip = "Places elements (like fixtures) in consistent positions relative to reference elements (like doors).",
                LargeImage = GetEmbeddedPng("CopyRelative.png") ?? defaultIcon
            };

            PushButtonData pbdScheduleLevel = new PushButtonData("CmdScheduleLevel", "Schedule\nLevel", ExecutingAssemblyPath, "RTS.Commands.ScheduleFloorClass")
            {
                ToolTip = "Automatically updates element schedule level parameters based on nearby floor elements.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdCastCeiling = new PushButtonData("CmdCastCeiling", "Cast to\nCeiling", ExecutingAssemblyPath, "RTS.Commands.CastCeilingClass")
            {
                ToolTip = "Moves MEP elements vertically to align with the nearest ceiling or slab surface.",
                LargeImage = defaultIcon
            };

            // Document Tools - View Tools
            PushButtonData pbdRtIsolate = new PushButtonData("CmdRtIsolate", "Isolate by\nID/Cable", ExecutingAssemblyPath, "RTS.Commands.RT_IsolateClass")
            {
                ToolTip = "Temporarily isolates elements in the current view based on RTS_ID or cable parameters.",
                LargeImage = GetEmbeddedPng("RT_Isolate.png") ?? defaultIcon
            };

            PushButtonData pbdViewportAlignment = new PushButtonData("CmdViewportAlignment", "Align\nViewports", ExecutingAssemblyPath, "RTS.Commands.ViewportAlignmentTool")
            {
                ToolTip = "Aligns selected viewports on sheets by centering them with a reference viewport.",
                LargeImage = GetEmbeddedPng("ViewportAlignment.png") ?? defaultIcon
            };

            PushButtonData pbdRtUppercase = new PushButtonData("CmdRtUppercase", "Uppercase\nText", ExecutingAssemblyPath, "RTS.Commands.RT_UpperCaseClass")
            {
                ToolTip = "Converts view names, sheet titles, numbers, and other text elements to uppercase.",
                LargeImage = GetEmbeddedPng("RT_UpperCase.png") ?? defaultIcon
            };

            // Document Tools - Project Management
            PushButtonData pbdLinkManager = new PushButtonData("CmdLinkManager", "Link\nManager", ExecutingAssemblyPath, "RTS.Commands.LinkManagerCommand")
            {
                ToolTip = "Centralized interface for managing Revit links and tracking associated metadata.",
                LargeImage = GetEmbeddedPng("LinkManager.png") ?? defaultIcon
            };

            // --- Add buttons to their respective panels based on the new organization ---

            // Project Setup Panel
            SplitButtonData splitSetupData = new SplitButtonData("ProjectSetupSplit", "Parameter\nTools")
            {
                LargeImage = parameterSetupIcon
            };
            SplitButton splitSetupButton = projectSetupPanel.AddItem(splitSetupData) as SplitButton;
            splitSetupButton.AddPushButton(pbdRtsInitiate);
            splitSetupButton.AddPushButton(pbdRtsMapCables);
            splitSetupButton.AddPushButton(pbdPcClearData);
            projectSetupPanel.AddItem(pbdRtsGenerateSchedules);
            projectSetupPanel.AddItem(pbdRtsReports);

            // Data Exchange Panel
            SplitButtonData splitImportData = new SplitButtonData("ImportDataSplit", "Import\nTools")
            {
                LargeImage = importDataIcon
            };
            SplitButton splitImportButton = dataExchangePanel.AddItem(splitImportData) as SplitButton;
            splitImportButton.AddPushButton(pbdPcSwbImporter);
            splitImportButton.AddPushButton(pbdPcCableImporter);
            splitImportButton.AddPushButton(pbdBbImport);
            splitImportButton.AddPushButton(pbdMdImporter);

            SplitButtonData splitExportData = new SplitButtonData("ExportDataSplit", "Export\nTools")
            {
                LargeImage = exportDataIcon
            };
            SplitButton splitExportButton = dataExchangePanel.AddItem(splitExportData) as SplitButton;
            splitExportButton.AddPushButton(pbdPcSwbExporter);
            splitExportButton.AddPushButton(pbdPcGenerateMd);
            
            dataExchangePanel.AddItem(pbdPcExtensible);
            dataExchangePanel.AddItem(pbdPcUpdater);

            // MEP Tools Panel
            SplitButtonData splitElectricalData = new SplitButtonData("ElectricalSplit", "Electrical\nTools")
            {
                LargeImage = electricalToolsIcon
            };
            SplitButton splitElectricalButton = mepToolsPanel.AddItem(splitElectricalData) as SplitButton;
            splitElectricalButton.AddPushButton(pbdPcWireData);
            splitElectricalButton.AddPushButton(pbdRtWireRoute);
            splitElectricalButton.AddPushButton(pbdRtCableLengths);
            splitElectricalButton.AddPushButton(pbdRtPanelConnect);

            SplitButtonData splitTrayData = new SplitButtonData("CableTraySplit", "Cable Tray\nTools")
            {
                LargeImage = cableTrayToolsIcon
            };
            SplitButton splitTrayButton = mepToolsPanel.AddItem(splitTrayData) as SplitButton;
            splitTrayButton.AddPushButton(pbdRtTrayOccupancy);
            splitTrayButton.AddPushButton(pbdRtTrayId);
            splitTrayButton.AddPushButton(pbdRtTrayConduits);

            SplitButtonData splitPositioningData = new SplitButtonData("PositioningSplit", "Positioning\nTools")
            {
                LargeImage = positioningToolsIcon
            };
            SplitButton splitPositioningButton = mepToolsPanel.AddItem(splitPositioningData) as SplitButton;
            splitPositioningButton.AddPushButton(pbdCopyRelative);
            splitPositioningButton.AddPushButton(pbdScheduleLevel);
            splitPositioningButton.AddPushButton(pbdCastCeiling);

            // Document Tools Panel
            SplitButtonData splitViewData = new SplitButtonData("ViewToolsSplit", "View\nTools")
            {
                LargeImage = viewToolsIcon
            };
            SplitButton splitViewButton = documentToolsPanel.AddItem(splitViewData) as SplitButton;
            splitViewButton.AddPushButton(pbdRtIsolate);
            splitViewButton.AddPushButton(pbdViewportAlignment);
            splitViewButton.AddPushButton(pbdRtUppercase);
            
            documentToolsPanel.AddItem(pbdLinkManager);
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

// This namespace contains the command to launch your window.
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
