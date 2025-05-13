// RTS.cs file
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;        // Often needed, good to have
using Autodesk.Revit.Attributes; // Required for TransactionMode, etc. if your commands use it
using System;
using System.Reflection;      // To get assembly path
using System.IO;              // For Path operations
using System.Windows.Media.Imaging; // For BitmapImage

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

                // Create a ribbon panel on your custom tab
                RibbonPanel rtsPanel = application.CreateRibbonPanel(tabName, "PCAD Tools");
                // Or, to add to the existing "Add-Ins" tab (uncomment one line below and comment the one above):
                // RibbonPanel rtsPanel = application.CreateRibbonPanel("Data Exchange");


                // --- Define PushButtonData for each command ---
                // For icons to work with the paths below, you need to:
                // 1. Create a folder in your Visual Studio project (e.g., named "Resources").
                // 2. Add your .png icon files (e.g., "ExportIcon32.png") to this "Resources" folder.
                // 3. In Solution Explorer (Visual Studio), select each icon file.
                // 4. In the Properties window for each icon file:
                //    - Set "Build Action" to "Content".
                //    - Set "Copy to Output Directory" to "Copy if newer" (or "Copy always").
                // This ensures icons are copied to a "Resources" subfolder alongside your RTS.dll.

                // 1. PC SWB Exporter
                // IMPORTANT: Verify the full class name "Namespace.ClassName" for your command.
                PushButtonData pbdPcSwbExporter = new PushButtonData(
                    "CmdPcSwbExporter",                          // Internal name, must be unique
                    "Export\nSWB Data",                          // Text displayed on the button
                    ExecutingAssemblyPath,                       // Assembly path
                    "PC_SWB_Exporter.PC_SWB_ExporterClass"       // Full class name (e.g., Namespace.ClassName)
                );
                pbdPcSwbExporter.ToolTip = "Exports data for PowerCAD SWB import.";
                // Uncomment and adjust if you have an icon (e.g., ExportIcon32.png):
                string exportIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_PC_Launcher.png");
                if (File.Exists(exportIconPath))
                {
                    pbdPcSwbExporter.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }
                // Optional: Tooltip image
                // string exportTooltipImagePath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "ExportTooltip.png");
                // if (File.Exists(exportTooltipImagePath))
                // {
                //     pbdPcSwbExporter.ToolTipImage = new BitmapImage(new Uri(exportTooltipImagePath));
                // }


                // 2. PC SWB Importer
                // IMPORTANT: Verify the full class name "Namespace.ClassName" for your command.
                // Based on your provided PC_SWB_Importer.cs, this class is in the "PC_SWB_Importer" namespace.
                PushButtonData pbdPcSwbImporter = new PushButtonData(
                    "CmdPcSwbImporter",
                    "Import\nSWB Data",
                    ExecutingAssemblyPath,
                    "PC_SWB_Importer.PC_SWB_ImporterClass"       // Full class name
                );
                pbdPcSwbImporter.ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters.";
                // Uncomment and adjust if you have an icon (e.g., ImportIcon32.png):
                // string importIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "ImportIcon32.png");
                // if (File.Exists(importIconPath))
                // {
                //     pbdPcSwbImporter.LargeImage = new BitmapImage(new Uri(importIconPath));
                // }

                // 3. Import BB Cable Lengths
                // IMPORTANT: Verify the full class name "Namespace.ClassName" for your command.
                PushButtonData pbdBbImport = new PushButtonData(
                    "CmdBbCableLengthImport",
                    "Import\n BB Lengths",
                    ExecutingAssemblyPath,
                    "BB_Import.BB_CableLengthImport"             // Full class name (e.g., Namespace.ClassName)
                );
                pbdBbImport.ToolTip = "Imports Bluebeam cable measurements into SLD components.";
                // Uncomment and adjust if you have an icon (e.g., BbImportIcon32.png):
                // string bbImportIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "BbImportIcon32.png");
                // if (File.Exists(bbImportIconPath))
                // {
                //     pbdBbImport.LargeImage = new BitmapImage(new Uri(bbImportIconPath));
                // }

                // 4. Import Cable Summary Data
                // IMPORTANT: Verify the full class name "Namespace.ClassName" for your command.
                PushButtonData pbdPcCableImporter = new PushButtonData(
                    "CmdPcCableImporter",
                    "Import Cable\nSummary",
                    ExecutingAssemblyPath,
                    "PC_Cable_Importer.PC_Cable_ImporterClass"   // Full class name (e.g., Namespace.ClassName)
                );
                pbdPcCableImporter.ToolTip = "Imports Cable Summary data from PowerCAD into SLD.";
                // Uncomment and adjust if you have an icon (e.g., CableSummaryIcon32.png):
                // string cableSummaryIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "CableSummaryIcon32.png");
                // if (File.Exists(cableSummaryIconPath))
                // {
                //     pbdPcCableImporter.LargeImage = new BitmapImage(new Uri(cableSummaryIconPath));
                // }

                // 5. PC Clear Data
                // IMPORTANT: Verify the full class name "Namespace.ClassName" for your command.
                // This class is in the "PC_Clear_Data" namespace and is named "PC_Clear_DataClass".
                PushButtonData pbdPcClearData = new PushButtonData(
                    "CmdPcClearData",                            // Internal name, must be unique
                    "Clear PCAD\nData",                          // Text displayed on the button
                    ExecutingAssemblyPath,                       // Assembly path
                    "PC_Clear_Data.PC_Clear_DataClass"           // Full class name
                );
                pbdPcClearData.ToolTip = "Clears specific PowerCAD-related parameters from Detail Items where PC_PowerCAD is 'Yes'.";
                // Uncomment and adjust if you have an icon (e.g., ClearDataIcon32.png):
                // string clearDataIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "ClearDataIcon32.png"); // Example icon name
                // if (File.Exists(clearDataIconPath))
                // {
                //     pbdPcClearData.LargeImage = new BitmapImage(new Uri(clearDataIconPath));
                // }


                // --- Add buttons to the panel ---
                // Example: Add them as individual large buttons
                rtsPanel.AddItem(pbdPcSwbExporter);
                rtsPanel.AddItem(pbdPcSwbImporter);
                rtsPanel.AddItem(pbdBbImport);
                rtsPanel.AddItem(pbdPcCableImporter);
                rtsPanel.AddItem(pbdPcClearData); // Added the new button here

                // Example: If you wanted to stack some of them using a PulldownButton or SplitButton
                // Create a PulldownButtonData
                // PulldownButtonData pdGroup = new PulldownButtonData("DataToolsGroup", "Data Tools");
                // pdGroup.ToolTip = "Access various data import/export tools.";
                // // Optional icon for the pulldown button itself
                // // string groupIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "DataGroupIcon32.png");
                // // if(File.Exists(groupIconPath)) { pdGroup.LargeImage = new BitmapImage(new Uri(groupIconPath)); }

                // Add the PulldownButton to the panel
                // PulldownButton pulldownButton = rtsPanel.AddItem(pdGroup) as PulldownButton;

                // Add previously defined PushButtonData objects to the PulldownButton
                // if (pulldownButton != null)
                // {
                //     pulldownButton.AddPushButton(pbdPcSwbExporter);
                //     pulldownButton.AddPushButton(pbdPcSwbImporter);
                // }
                // rtsPanel.AddSeparator(); // Optional separator
                // // Add other buttons individually
                // rtsPanel.AddItem(pbdBbImport);
                // rtsPanel.AddItem(pbdPcCableImporter);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Provide more details in the error message if possible
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
            // Clean up resources if needed, though often not necessary for simple ribbon items.
            return Result.Succeeded;
        }
    }
}