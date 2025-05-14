// RTS.cs file
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;         // Often needed, good to have
using Autodesk.Revit.Attributes; // Required for TransactionMode, etc. if your commands use it
using System;
using System.Reflection;         // To get assembly path
using System.IO;                 // For Path operations
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
                PushButtonData pbdPcSwbExporter = new PushButtonData(
                    "CmdPcSwbExporter",                      // Internal name, must be unique
                    "Export\nSWB Data",                      // Text displayed on the button
                    ExecutingAssemblyPath,                   // Assembly path
                    "PC_SWB_Exporter.PC_SWB_ExporterClass"   // Full class name (e.g., Namespace.ClassName)
                );
                pbdPcSwbExporter.ToolTip = "Exports data for PowerCAD SWB import.";
                string exportIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_PC_Launcher.png");
                if (File.Exists(exportIconPath))
                {
                    pbdPcSwbExporter.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }

                // 2. PC SWB Importer
                PushButtonData pbdPcSwbImporter = new PushButtonData(
                    "CmdPcSwbImporter",
                    "Import\nSWB Data",
                    ExecutingAssemblyPath,
                    "PC_SWB_Importer.PC_SWB_ImporterClass"   // Full class name
                );
                pbdPcSwbImporter.ToolTip = "Imports cable data from a PowerCAD SWB CSV export file into filtered Detail Item parameters.";
                // Icon for Importer (assuming same as exporter for now, or create a new one)
                if (File.Exists(exportIconPath)) // Re-using exportIconPath for example
                {
                    pbdPcSwbImporter.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }


                // 3. Import BB Cable Lengths
                PushButtonData pbdBbImport = new PushButtonData(
                    "CmdBbCableLengthImport",
                    "Import\n BB Lengths",
                    ExecutingAssemblyPath,
                    "BB_Import.BB_CableLengthImport"         // Full class name (e.g., Namespace.ClassName)
                );
                pbdBbImport.ToolTip = "Imports Bluebeam cable measurements into SLD components.";
                // Icon for BB Import (assuming same as exporter for now, or create a new one)
                if (File.Exists(exportIconPath)) // Re-using exportIconPath for example
                {
                    pbdBbImport.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }

                // 4. Import Cable Summary Data
                PushButtonData pbdPcCableImporter = new PushButtonData(
                    "CmdPcCableImporter",
                    "Import Cable\nSummary",
                    ExecutingAssemblyPath,
                    "PC_Cable_Importer.PC_Cable_ImporterClass"   // Full class name (e.g., Namespace.ClassName)
                );
                pbdPcCableImporter.ToolTip = "Imports Cable Summary data from PowerCAD into SLD.";
                // Icon for Cable Importer (assuming same as exporter for now, or create a new one)
                if (File.Exists(exportIconPath)) // Re-using exportIconPath for example
                {
                    pbdPcCableImporter.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }


                // 5. PC Clear Data
                PushButtonData pbdPcClearData = new PushButtonData(
                    "CmdPcClearData",                        // Internal name, must be unique
                    "Clear PCAD\nData",                      // Text displayed on the button
                    ExecutingAssemblyPath,                   // Assembly path
                    "PC_Clear_Data.PC_Clear_DataClass"       // Full class name
                );
                pbdPcClearData.ToolTip = "Clears specific PowerCAD-related parameters from Detail Items where PC_PowerCAD is 'Yes'.";
                // Icon for Clear Data (assuming same as exporter for now, or create a new one)
                if (File.Exists(exportIconPath)) // Re-using exportIconPath for example
                {
                    pbdPcClearData.LargeImage = new BitmapImage(new Uri(exportIconPath));
                }

                // 6. PC Generate MD Report (NEW BUTTON)
                PushButtonData pbdPcGenerateMd = new PushButtonData(
                    "CmdPcGenerateMd",                       // Internal name, must be unique
                    "Generate\nMD Report",                   // Text displayed on the button
                    ExecutingAssemblyPath,                   // Assembly path
                    "PC_Generate_MD.PC_Generate_MDClass"     // Full class name from the immersive artifact
                );
                pbdPcGenerateMd.ToolTip = "Generates an Excel report with Cover Page, Authority, and Submains data.";
                // Optional: Set an icon for this new button
                string mdReportIconPath = Path.Combine(ExecutingAssemblyDirectory, "Resources", "Icon_MD_Report.png"); // Example icon name
                if (File.Exists(mdReportIconPath))
                {
                    pbdPcGenerateMd.LargeImage = new BitmapImage(new Uri(mdReportIconPath));
                }
                // Optional: If you have a command availability class for this command
                // pbdPcGenerateMd.AvailabilityClassName = "PC_Generate_MD.CommandAvailability"; // Example


                // --- Add buttons to the panel ---
                // Example: Add them as individual large buttons
                rtsPanel.AddItem(pbdPcSwbExporter);
                rtsPanel.AddItem(pbdPcSwbImporter);
                rtsPanel.AddItem(pbdBbImport);
                rtsPanel.AddItem(pbdPcCableImporter);
                rtsPanel.AddItem(pbdPcClearData);
                rtsPanel.AddItem(pbdPcGenerateMd); // Added the new MD Report button here

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
