// Add these using statements at the top of your .cs file
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes; // Required for TransactionMode, etc. if your commands use it
using System;
using System.Reflection; // To get assembly path
using System.IO; // For Path operations
using System.Windows.Media.Imaging; // For BitmapImage

// Define a namespace for your application
namespace RTS
{
    public class App : IExternalApplication
    {
        // Path to this assembly
        static string ExecutingAssemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create a custom ribbon tab
                string tabName = "ReTick Solutions";
                application.CreateRibbonTab(tabName);

                // Create a ribbon panel
                RibbonPanel rtsPanel = application.CreateRibbonPanel(tabName, "Data Exchange");
                // Or, to add to the existing "Add-Ins" tab:
                // RibbonPanel rtsPanel = application.CreateRibbonPanel("Data Exchange");


                // --- Define PushButtonData for each command ---

                // 1. PC SWB Exporter
                PushButtonData pbdPcSwbExporter = new PushButtonData(
                    "CmdPcSwbExporter",                           // Internal name, must be unique
                    "Export\nSWB Data",                           // Text displayed on the button
                    ExecutingAssemblyPath,                        // Assembly path
                    "PC_SWB_Exporter.PC_SWB_ExporterClass"        // Full class name of the command
                );
                pbdPcSwbExporter.ToolTip = "Exports data for PowerCAD SWB import.";
                // Optional: Add an icon (32x32 PNG recommended)
                // pbdPcSwbExporter.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(ExecutingAssemblyPath), "Resources", "ExportIcon32.png")));
                // pbdPcSwbExporter.ToolTipImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(ExecutingAssemblyPath), "Resources", "ExportTooltip.png")));


                // 2. PC SWB Importer
                PushButtonData pbdPcSwbImporter = new PushButtonData(
                    "CmdPcSwbImporter",
                    "Import\nSWB Data",
                    ExecutingAssemblyPath,
                    "PC_SWB_Importer.PC_SWB_ImporterClass"
                );
                pbdPcSwbImporter.ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters.";
                // pbdPcSwbImporter.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(ExecutingAssemblyPath), "Resources", "ImportIcon32.png")));


                // 3. Import BB Cable Lengths
                PushButtonData pbdBbImport = new PushButtonData(
                    "CmdBbCableLengthImport",
                    "Import BB\nLengths",
                    ExecutingAssemblyPath,
                    "BB_Import.BB_CableLengthImport"
                );
                pbdBbImport.ToolTip = "Imports Bluebeam cable measurements into SLD components.";
                // pbdBbImport.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(ExecutingAssemblyPath), "Resources", "BbImportIcon32.png")));


                // 4. Import Cable Summary Data
                PushButtonData pbdPcCableImporter = new PushButtonData(
                    "CmdPcCableImporter",
                    "Import Cable\nSummary",
                    ExecutingAssemblyPath,
                    "PC_Cable_Importer.PC_Cable_ImporterClass"
                );
                pbdPcCableImporter.ToolTip = "Imports Cable Summary data from PowerCAD into SLD.";
                // pbdPcCableImporter.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(ExecutingAssemblyPath), "Resources", "CableSummaryIcon32.png")));


                // --- Add buttons to the panel ---
                // You can add them individually or as stacked items
                
                // Example: Add them as individual large buttons
                rtsPanel.AddItem(pbdPcSwbExporter);
                rtsPanel.AddItem(pbdPcSwbImporter);
                rtsPanel.AddItem(pbdBbImport);
                rtsPanel.AddItem(pbdPcCableImporter);

                // Example: If you wanted to stack some of them
                // PulldownButtonData pdGroup = new PulldownButtonData("DataGroup", "Data Tools");
                // PulldownButton pulldown = rtsPanel.AddItem(pdGroup) as PulldownButton;
                // pulldown.AddPushButton(pbdPcSwbExporter);
                // pulldown.AddPushButton(pbdPcSwbImporter);
                // rtsPanel.AddSeparator();
                // rtsPanel.AddItem(pbdBbImport);
                // rtsPanel.AddItem(pbdPcCableImporter);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RTS Add-In Error", "Failed to initialize ReTick Solutions ribbon items:\n" + ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Nothing to do in this example
            return Result.Succeeded;
        }
    }
}