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
// - September 17, 2025: Corrected namespace to match the .addin manifest file.
// - September 16, 2025: Added backing commands for Diagnostics panel.
// - September 16, 2025: Added Diagnostics panel and proactive resource verification at startup.
// - September 15, 2025: Corrected DiagnosticsManager initialization call to handle OnStartup context.
// - September 15, 2025: Integrated DiagnosticsManager for centralized logging and crash reporting.
// - August 16, 2025: Reorganized codebase to use new folder structure with centralized utilities.
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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using RTS.Utilities; // Centralized utilities including DiagnosticsManager
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
                // Initialize the DiagnosticsManager using the UIControlledApplication.
                // This is the new, preferred method that correctly captures the Revit version.
                DiagnosticsManager.Initialize(application);

                // Create the ribbon panel during startup.
                CreateRibbonPanel(application);

                // Log successful startup.
                DiagnosticsManager.LogMessage(DiagnosticsManager.LogLevel.Info, "RTS Add-in started successfully.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If anything goes wrong, use the DiagnosticsManager to show a detailed
                // crash report dialog and log the exception.
                DiagnosticsManager.ShowCrashDialog(ex, "RTS Add-in failed to load.");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Log the shutdown event for diagnostic purposes.
            DiagnosticsManager.LogMessage(DiagnosticsManager.LogLevel.Info, "RTS Add-in is shutting down.");
            return Result.Succeeded;
        }

        /// <summary>
        /// Creates the ribbon tab and panels for the add-in.
        /// </summary>
        /// <param name="application">The UIControlledApplication instance.</param>
        private void CreateRibbonPanel(UIControlledApplication application)
        {
            string tabName = "RTS";
            // List to hold all expected icon names for verification
            var expectedIcons = new List<string>();

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
            RibbonPanel projectSetupPanel = application.CreateRibbonPanel(tabName, "Project Setup");
            RibbonPanel dataExchangePanel = application.CreateRibbonPanel(tabName, "Data Exchange");
            RibbonPanel mepToolsPanel = application.CreateRibbonPanel(tabName, "MEP Tools");
            RibbonPanel documentToolsPanel = application.CreateRibbonPanel(tabName, "Document Tools");
            RibbonPanel diagnosticsPanel = application.CreateRibbonPanel(tabName, "Diagnostics");

            // --- Load default and specific icons ---
            expectedIcons.Add("Icon.png");
            BitmapImage defaultIcon = GetEmbeddedPng("Icon.png");

            // --- Load category icons for split buttons ---
            expectedIcons.AddRange(new[] {
                "ParameterSetup_Icon.png", "ImportData_Icon.png", "ExportData_Icon.png",
                "ElectricalTools_Icon.png", "CableTrayTools_Icon.png", "PositioningTools_Icon.png",
                "ViewTools_Icon.png", "Diagnostics_Icon.png"
            });
            BitmapImage parameterSetupIcon = GetEmbeddedPng("ParameterSetup_Icon.png") ?? defaultIcon;
            BitmapImage importDataIcon = GetEmbeddedPng("ImportData_Icon.png") ?? defaultIcon;
            BitmapImage exportDataIcon = GetEmbeddedPng("ExportData_Icon.png") ?? defaultIcon;
            BitmapImage electricalToolsIcon = GetEmbeddedPng("ElectricalTools_Icon.png") ?? defaultIcon;
            BitmapImage cableTrayToolsIcon = GetEmbeddedPng("CableTrayTools_Icon.png") ?? defaultIcon;
            BitmapImage positioningToolsIcon = GetEmbeddedPng("PositioningTools_Icon.png") ?? defaultIcon;
            BitmapImage viewToolsIcon = GetEmbeddedPng("ViewTools_Icon.png") ?? defaultIcon;
            BitmapImage diagnosticsIcon = GetEmbeddedPng("Diagnostics_Icon.png") ?? defaultIcon;

            // --- Create All Button Definitions ---

            // Helper function to add icon names to the verification list
            void AddIcon(string name) { if (!string.IsNullOrEmpty(name)) expectedIcons.Add(name); }

            // Project Setup - Parameter Setup
            AddIcon("RTS_Initiate.png");
            PushButtonData pbdRtsInitiate = new PushButtonData("CmdRtsInitiate", "Initialize\nParameters", ExecutingAssemblyPath, "RTS.Commands.ProjectSetup.RTS_InitiateClass")
            {
                ToolTip = "Sets up required shared parameters needed for RTS tools. Must be run before using other tools.",
                LargeImage = GetEmbeddedPng("RTS_Initiate.png") ?? defaultIcon
            };

            AddIcon("RTS_MapCables.png");
            PushButtonData pbdRtsMapCables = new PushButtonData("CmdRtsMapCables", "Map\nParameters", ExecutingAssemblyPath, "RTS.Commands.ProjectSetup.RTS_MapCablesClass")
            {
                ToolTip = "Maps standard RTS parameters to custom project parameters using a CSV file.",
                LargeImage = GetEmbeddedPng("RTS_MapCables.png") ?? defaultIcon
            };

            PushButtonData pbdPcClearData = new PushButtonData("CmdPcClearData", "Clear\nData", ExecutingAssemblyPath, "RTS.Commands.ProjectSetup.PC_Clear_DataClass")
            {
                ToolTip = "Clears PowerCAD-related parameters from elements while keeping core identifiers intact.",
                LargeImage = defaultIcon
            };

            // Project Setup - Documentation
            AddIcon("RTS_Schedules.png");
            PushButtonData pbdRtsGenerateSchedules = new PushButtonData("CmdRtsGenerateSchedules", "Generate\nSchedules", ExecutingAssemblyPath, "RTS.Commands.ProjectSetup.RTS_SchedulesClass")
            {
                ToolTip = "Creates or refreshes a standard set of electrical schedules based on RTS parameters.",
                LargeImage = GetEmbeddedPng("RTS_Schedules.png") ?? defaultIcon
            };

            AddIcon("RTS_Reports.png");
            PushButtonData pbdRtsReports = new PushButtonData("CmdRtsReports", "Generate\nReports", ExecutingAssemblyPath, "RTS.Commands.ProjectSetup.RTS_ReportsClass")
            {
                ToolTip = "Creates detailed reports from project data for documentation and coordination.",
                LargeImage = GetEmbeddedPng("RTS_Reports.png") ?? defaultIcon
            };

            // Data Exchange - Import Tools
            PushButtonData pbdPcSwbImporter = new PushButtonData("CmdPcSwbImporter", "Import\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Import.PC_SWB_ImporterClass")
            {
                ToolTip = "Imports switchboard and cable data from PowerCAD CSV exports into Revit.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdPcCableImporter = new PushButtonData("CmdPcCableImporter", "Import Cable\nSummary", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Import.PC_Cable_ImporterClass")
            {
                ToolTip = "Imports cable information from PowerCAD Cable Summary documents into Single Line Diagrams.",
                LargeImage = defaultIcon
            };

            AddIcon("BB_Import.png");
            PushButtonData pbdBbImport = new PushButtonData("CmdBbCableLengthImport", "Import\nBluebeam", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Import.BB_CableLengthImport")
            {
                ToolTip = "Imports cable length measurements from Bluebeam markups into Revit elements.",
                LargeImage = GetEmbeddedPng("BB_Import.png") ?? defaultIcon
            };

            AddIcon("MD_Importer.png");
            PushButtonData pbdMdImporter = new PushButtonData("CmdMdExcelUpdater", "Update SWB\nLoads", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Import.MD_ImporterClass")
            {
                ToolTip = "Updates switchboard load parameters from Excel spreadsheets using TB_Submains table.",
                LargeImage = GetEmbeddedPng("MD_Importer.png") ?? defaultIcon
            };

            // Data Exchange - Export Tools
            PushButtonData pbdPcSwbExporter = new PushButtonData("CmdPcSwbExporter", "Export\nSWB Data", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Export.PC_SWB_ExporterClass")
            {
                ToolTip = "Exports switchboard and cable data for use in PowerCAD or other tools.",
                LargeImage = defaultIcon
            };

            AddIcon("PC_Generate_MD.png");
            PushButtonData pbdPcGenerateMd = new PushButtonData("CmdPcGenerateMd", "Generate MD\nReport", ExecutingAssemblyPath, "RTS.Commands.DataExchange.Export.PC_Generate_MDClass")
            {
                ToolTip = "Creates a Maximum Demand Excel report with Cover Page, Authority, and Submains details.",
                LargeImage = GetEmbeddedPng("PC_Generate_MD.png") ?? defaultIcon
            };

            // Data Exchange - Data Management
            AddIcon("PC_Extensible.png");
            PushButtonData pbdPcExtensible = new PushButtonData("CmdPcExtensible", "Process & Store\nCable Data", ExecutingAssemblyPath, "RTS.Commands.DataExchange.DataManagement.PC_ExtensibleClass")
            {
                ToolTip = "Processes cable schedules and stores data in project extensible storage for use by other tools.",
                LargeImage = GetEmbeddedPng("PC_Extensible.png") ?? defaultIcon
            };

            AddIcon("PC_Updater.png");
            PushButtonData pbdPcUpdater = new PushButtonData("CmdPcUpdater", "Update PCAD\nCSV", ExecutingAssemblyPath, "RTS.Commands.DataExchange.DataManagement.PC_UpdaterClass")
            {
                ToolTip = "Updates cable lengths in PowerCAD CSV exports using data from the Revit model.",
                LargeImage = GetEmbeddedPng("PC_Updater.png") ?? defaultIcon
            };

            // MEP Tools - Electrical
            AddIcon("PC_WireData.png");
            PushButtonData pbdPcWireData = new PushButtonData("CmdPcWireData", "Update\nWires", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Electrical.PC_WireDataClass")
            {
                ToolTip = "Creates or updates electrical wires in the Revit model based on cable data.",
                LargeImage = GetEmbeddedPng("PC_WireData.png") ?? defaultIcon
            };

            AddIcon("RT_WireRoute.png");
            PushButtonData pbdRtWireRoute = new PushButtonData("CmdRtWireRoute", "Route\nWires", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Electrical.RT_WireRouteClass")
            {
                ToolTip = "Automatically routes electrical wires through conduits based on matching RTS_ID parameters.",
                LargeImage = GetEmbeddedPng("RT_WireRoute.png") ?? defaultIcon
            };

            AddIcon("RT_CableLengths.png");
            PushButtonData pbdRtCableLengths = new PushButtonData("CmdRTCableLengths", "Update Cable\nLengths", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Electrical.RTCableLengthsCommand")
            {
                ToolTip = "Calculates and updates cable lengths based on conduit/tray paths and connections.",
                LargeImage = GetEmbeddedPng("RT_CableLengths.png") ?? defaultIcon
            };

            PushButtonData pbdRtPanelConnect = new PushButtonData("CmdRtPanelConnect", "Connect\nPanels", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Electrical.RT_PanelConnectClass")
            {
                ToolTip = "Creates power connections between electrical panels based on source/destination data.",
                LargeImage = defaultIcon
            };

            // MEP Tools - Cable Tray
            AddIcon("RT_TrayOccupancy.png");
            PushButtonData pbdRtTrayOccupancy = new PushButtonData("CmdRtTrayOccupancy", "Process\nTray Data", ExecutingAssemblyPath, "RTS.Commands.MEPTools.CableTray.RT_TrayOccupancyClass")
            {
                ToolTip = "Analyzes cable tray requirements based on cable schedule data and routing.",
                LargeImage = GetEmbeddedPng("RT_TrayOccupancy.png") ?? defaultIcon
            };

            PushButtonData pbdRtTrayId = new PushButtonData("CmdRtTrayId", "Update Tray\nIDs", ExecutingAssemblyPath, "RTS.Commands.MEPTools.CableTray.RT_TrayIDClass")
            {
                ToolTip = "Assigns unique identifiers to cable tray elements for coordination and scheduling.",
                LargeImage = defaultIcon
            };

            AddIcon("RT_TrayConduits.png");
            PushButtonData pbdRtTrayConduits = new PushButtonData("CmdRtTrayConduits", "Create Tray\nConduits", ExecutingAssemblyPath, "RTS.Commands.MEPTools.CableTray.RT_TrayConduitsClass")
            {
                ToolTip = "Creates conduit runs along cable trays based on cable requirements and routes.",
                LargeImage = GetEmbeddedPng("RT_TrayConduits.png") ?? defaultIcon
            };

            // MEP Tools - Element Positioning
            AddIcon("CopyRelative.png");
            PushButtonData pbdCopyRelative = new PushButtonData("CmdCopyRelative", "Copy\nRelative", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Positioning.CopyRelativeClass")
            {
                ToolTip = "Places elements (like fixtures) in consistent positions relative to reference elements (like doors).",
                LargeImage = GetEmbeddedPng("CopyRelative.png") ?? defaultIcon
            };

            PushButtonData pbdScheduleLevel = new PushButtonData("CmdScheduleLevel", "Schedule\nLevel", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Positioning.ScheduleFloorClass")
            {
                ToolTip = "Automatically updates element schedule level parameters based on nearby floor elements.",
                LargeImage = defaultIcon
            };

            PushButtonData pbdCastCeiling = new PushButtonData("CmdCastCeiling", "Cast to\nCeiling", ExecutingAssemblyPath, "RTS.Commands.MEPTools.Positioning.CastCeilingClass")
            {
                ToolTip = "Moves MEP elements vertically to align with the nearest ceiling or slab surface.",
                LargeImage = defaultIcon
            };

            // Document Tools - View Tools
            AddIcon("RT_Isolate.png");
            PushButtonData pbdRtIsolate = new PushButtonData("CmdRtIsolate", "Isolate by\nID/Cable", ExecutingAssemblyPath, "RTS.Commands.DocumentTools.ViewTools.RT_IsolateClass")
            {
                ToolTip = "Temporarily isolates elements in the current view based on RTS_ID or cable parameters.",
                LargeImage = GetEmbeddedPng("RT_Isolate.png") ?? defaultIcon
            };

            AddIcon("ViewportAlignment.png");
            PushButtonData pbdViewportAlignment = new PushButtonData("CmdViewportAlignment", "Align\nViewports", ExecutingAssemblyPath, "RTS.Commands.DocumentTools.ViewTools.ViewportAlignmentTool")
            {
                ToolTip = "Aligns selected viewports on sheets by centering them with a reference viewport.",
                LargeImage = GetEmbeddedPng("ViewportAlignment.png") ?? defaultIcon
            };

            AddIcon("RT_UpperCase.png");
            PushButtonData pbdRtUppercase = new PushButtonData("CmdRtUppercase", "Uppercase\nText", ExecutingAssemblyPath, "RTS.Commands.DocumentTools.ViewTools.RT_UpperCaseClass")
            {
                ToolTip = "Converts view names, sheet titles, numbers, and other text elements to uppercase.",
                LargeImage = GetEmbeddedPng("RT_UpperCase.png") ?? defaultIcon
            };

            // Document Tools - Project Management
            AddIcon("LinkManager.png");
            PushButtonData pbdLinkManager = new PushButtonData("CmdLinkManager", "Link\nManager", ExecutingAssemblyPath, "RTS.Commands.DocumentTools.ProjectManagement.LinkManagerCommand")
            {
                ToolTip = "Centralized interface for managing Revit links and tracking associated metadata.",
                LargeImage = GetEmbeddedPng("LinkManager.png") ?? defaultIcon
            };

            // Diagnostics Tools
            AddIcon("LogFolder_Icon.png");
            PushButtonData pbdOpenLogFolder = new PushButtonData("CmdOpenLogFolder", "Open Log\nFolder", ExecutingAssemblyPath, "RTS.Commands.Diagnostics.OpenLogFolderCommand")
            {
                ToolTip = "Opens the folder containing all RTS log and crash report files.",
                LargeImage = GetEmbeddedPng("LogFolder_Icon.png") ?? diagnosticsIcon
            };

            AddIcon("LogLevel_Icon.png");
            PushButtonData pbdSetLogLevel = new PushButtonData("CmdSetLogLevel", "Set Log\nLevel", ExecutingAssemblyPath, "RTS.Commands.Diagnostics.SetLogLevelCommand")
            {
                ToolTip = "Sets the detail level for the RTS log files. Higher detail can help with troubleshooting.",
                LargeImage = GetEmbeddedPng("LogLevel_Icon.png") ?? diagnosticsIcon
            };

            AddIcon("CheckResources_Icon.png");
            PushButtonData pbdCheckResources = new PushButtonData("CmdCheckResources", "Check\nResources", ExecutingAssemblyPath, "RTS.Commands.Diagnostics.CheckResourcesCommand")
            {
                ToolTip = "Confirms that all required embedded resources like icons are available.",
                LargeImage = GetEmbeddedPng("CheckResources_Icon.png") ?? diagnosticsIcon
            };


            // --- Add buttons to their respective panels ---

            // Project Setup Panel
            var splitSetupButton = projectSetupPanel.AddItem(new SplitButtonData("ProjectSetupSplit", "Parameter\nTools") { LargeImage = parameterSetupIcon }) as SplitButton;
            splitSetupButton.AddPushButton(pbdRtsInitiate);
            splitSetupButton.AddPushButton(pbdRtsMapCables);
            splitSetupButton.AddPushButton(pbdPcClearData);
            projectSetupPanel.AddItem(pbdRtsGenerateSchedules);
            projectSetupPanel.AddItem(pbdRtsReports);

            // Data Exchange Panel
            var splitImportButton = dataExchangePanel.AddItem(new SplitButtonData("ImportDataSplit", "Import\nTools") { LargeImage = importDataIcon }) as SplitButton;
            splitImportButton.AddPushButton(pbdPcSwbImporter);
            splitImportButton.AddPushButton(pbdPcCableImporter);
            splitImportButton.AddPushButton(pbdBbImport);
            splitImportButton.AddPushButton(pbdMdImporter);

            var splitExportButton = dataExchangePanel.AddItem(new SplitButtonData("ExportDataSplit", "Export\nTools") { LargeImage = exportDataIcon }) as SplitButton;
            splitExportButton.AddPushButton(pbdPcSwbExporter);
            splitExportButton.AddPushButton(pbdPcGenerateMd);

            dataExchangePanel.AddStackedItems(pbdPcExtensible, pbdPcUpdater);

            // MEP Tools Panel
            var splitElectricalButton = mepToolsPanel.AddItem(new SplitButtonData("ElectricalSplit", "Electrical\nTools") { LargeImage = electricalToolsIcon }) as SplitButton;
            splitElectricalButton.AddPushButton(pbdPcWireData);
            splitElectricalButton.AddPushButton(pbdRtWireRoute);
            splitElectricalButton.AddPushButton(pbdRtCableLengths);
            splitElectricalButton.AddPushButton(pbdRtPanelConnect);

            var splitTrayButton = mepToolsPanel.AddItem(new SplitButtonData("CableTraySplit", "Cable Tray\nTools") { LargeImage = cableTrayToolsIcon }) as SplitButton;
            splitTrayButton.AddPushButton(pbdRtTrayOccupancy);
            splitTrayButton.AddPushButton(pbdRtTrayId);
            splitTrayButton.AddPushButton(pbdRtTrayConduits);

            var splitPositioningButton = mepToolsPanel.AddItem(new SplitButtonData("PositioningSplit", "Positioning\nTools") { LargeImage = positioningToolsIcon }) as SplitButton;
            splitPositioningButton.AddPushButton(pbdCopyRelative);
            splitPositioningButton.AddPushButton(pbdScheduleLevel);
            splitPositioningButton.AddPushButton(pbdCastCeiling);

            // Document Tools Panel
            var splitViewButton = documentToolsPanel.AddItem(new SplitButtonData("ViewToolsSplit", "View\nTools") { LargeImage = viewToolsIcon }) as SplitButton;
            splitViewButton.AddPushButton(pbdRtIsolate);
            splitViewButton.AddPushButton(pbdViewportAlignment);
            splitViewButton.AddPushButton(pbdRtUppercase);

            documentToolsPanel.AddItem(pbdLinkManager);

            // Diagnostics Panel
            diagnosticsPanel.AddItem(pbdOpenLogFolder);
            diagnosticsPanel.AddItem(pbdSetLogLevel);
            diagnosticsPanel.AddItem(pbdCheckResources);

            // --- Perform Resource Verification ---
            // Now that all buttons are defined, check if their icons exist.
            DiagnosticsManager.VerifyEmbeddedResources(Assembly.GetExecutingAssembly(), "RTS.Resources.", expectedIcons);
        }

        private BitmapImage GetEmbeddedPng(string imageName)
        {
            try
            {
                // Assuming "Resources" is a direct subfolder of the assembly's root namespace (RTS).
                string resourceName = "RTS.Resources." + imageName;
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // The verification method will log a consolidated warning.
                        // We no longer need to log here.
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
                // Log the exception for debugging purposes.
                DiagnosticsManager.LogException(ex, $"Failed to load embedded PNG: {imageName}");
                return null;
            }
        }
    }
}
