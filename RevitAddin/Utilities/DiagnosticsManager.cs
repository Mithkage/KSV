//
// File: DiagnosticsManager.cs
//
// Namespace: RTS.Utilities
//
// Class: DiagnosticsManager
//
// Function: Provides centralized crash reporting, logging, and diagnostic utilities
//           for better troubleshooting of Revit add-in issues. Supports both
//           version-specific and common error handling patterns with configurable
//           logging levels and format. Includes performance timing and resource verification.
//
// Author: GitHub Copilot / Kyle Vorster
// Company: ReTick Solutions (RTS)
//
// Log:
// - September 17, 2025: Added static constructor and failsafe log for early startup diagnostics.
// - September 16, 2025: Made LogDirectory property public for external access.
// - September 16, 2025: Implemented proposed improvements: performance timing, resource verification, and improved initialization.
// - September 15, 2025: Initial creation of the diagnostic utility class.
//

#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
#endregion

namespace RTS.Utilities
{
    /// <summary>
    /// Central diagnostics manager for RTS add-in error handling, logging, and crash reporting
    /// </summary>
    public static class DiagnosticsManager
    {
        #region Properties and Fields

        // Constants for logging configuration
        private const string LOG_DIRECTORY = "RTS_Logs";
        private const string CRASH_FILE_NAME = "RTS_Crash_Report_{0}.txt";
        private const string LOG_FILE_NAME = "RTS_Log_{0}.txt";
        private const int MAX_LOG_FILES = 10;
        private const int MAX_STACK_FRAMES = 20;

        // Log levels
        public enum LogLevel { Debug, Info, Warning, Error, Fatal }

        // Current log level threshold - messages below this level won't be logged
        public static LogLevel CurrentLogLevel { get; set; } = LogLevel.Info;

        // Path to log directory, now public for access by diagnostic commands.
        public static string LogDirectory => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            LOG_DIRECTORY);

        // Current crash report file path (set when a crash report is generated)
        private static string _currentCrashFilePath;

        // Track currently active operations for better context in crash reports
        private static readonly Dictionary<int, string> _activeOperations = new Dictionary<int, string>();

        // Cache Revit version for reporting
        private static string _revitVersion;

        #endregion

        #region Constructor

        /// <summary>
        /// Static constructor to run a failsafe log as soon as the class is loaded by the .NET runtime.
        /// This happens before any methods are called, providing a way to diagnose pre-startup failures.
        /// </summary>
        static DiagnosticsManager()
        {
            FailsafeLog("DiagnosticsManager static constructor called. .NET runtime has loaded the assembly.");
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Initialize the diagnostics manager at application startup from a UIControlledApplication.
        /// </summary>
        /// <param name="app">The UIControlledApplication from OnStartup.</param>
        public static void Initialize(UIControlledApplication app)
        {
            FailsafeLog("Initialize method called.");
            InitializeInternal(app?.ControlledApplication?.VersionName);
        }

        /// <summary>
        /// Initialize the diagnostics manager from a UIApplication context (e.g., inside a command).
        /// </summary>
        /// <param name="app">Revit application object.</param>
        public static void Initialize(UIApplication app)
        {
            InitializeInternal(app?.Application?.VersionName);
        }

        /// <summary>
        /// Log a message with specified level
        /// </summary>
        public static void LogMessage(LogLevel level, string message)
        {
            if (level < CurrentLogLevel)
                return;

            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
                WriteToLogFile(logMessage);

                // Also write to Debug output for development
                Debug.WriteLine(logMessage);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error writing to log: " + ex.Message);
            }
        }

        /// <summary>
        /// Log exception details with contextual information
        /// </summary>
        public static void LogException(Exception ex, string context = null, Document doc = null)
        {
            try
            {
                string currentOperation = GetCurrentOperation();
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"EXCEPTION at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Context: {context ?? "Not specified"}");
                sb.AppendLine($"Current Operation: {currentOperation}");
                sb.AppendLine($"Revit Version: {_revitVersion}");
                if (doc != null)
                {
                    sb.AppendLine($"Document: {doc.Title} (v{doc.WorksharingCentralGUID})");
                }

                AppendExceptionDetails(sb, ex);

                string logMessage = sb.ToString();
                WriteToLogFile(logMessage);
                Debug.WriteLine(logMessage);
            }
            catch (Exception logEx)
            {
                Debug.WriteLine("Error logging exception: " + logEx.Message);
            }
        }

        /// <summary>
        /// Verifies that a list of expected embedded resources exist in the assembly.
        /// </summary>
        /// <param name="assembly">The assembly to check.</param>
        /// <param name="baseResourcePath">The base path for resources, e.g., "RTS.Resources."</param>
        /// <param name="expectedResourceNames">A list of simple file names (e.g., "Icon.png").</param>
        public static void VerifyEmbeddedResources(Assembly assembly, string baseResourcePath, IEnumerable<string> expectedResourceNames)
        {
            LogMessage(LogLevel.Debug, "Starting embedded resource verification...");
            var manifestResourceNames = new HashSet<string>(assembly.GetManifestResourceNames());
            var missingResources = new List<string>();

            foreach (var resourceName in expectedResourceNames)
            {
                string fullResourcePath = baseResourcePath + resourceName;
                if (!manifestResourceNames.Contains(fullResourcePath))
                {
                    missingResources.Add(resourceName);
                }
            }

            if (missingResources.Any())
            {
                LogMessage(LogLevel.Warning, $"Missing {missingResources.Count} embedded resources: {string.Join(", ", missingResources)}");
            }
            else
            {
                LogMessage(LogLevel.Info, "All expected embedded resources found.");
            }
        }


        /// <summary>
        /// Generates a detailed crash report for the exception
        /// </summary>
        public static string GenerateCrashReport(Exception ex, string context = null, Document doc = null)
        {
            try
            {
                string crashFilePath = Path.Combine(LogDirectory,
                    string.Format(CRASH_FILE_NAME, DateTime.Now.ToString("yyyyMMdd_HHmmss")));

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("==================================================");
                sb.AppendLine("RTS ADD-IN CRASH REPORT");
                sb.AppendLine("==================================================");
                sb.AppendLine($"Date/Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"RTS Add-in Version: {Assembly.GetExecutingAssembly().GetName().Version}");
                sb.AppendLine($"Revit Version: {_revitVersion}");
                sb.AppendLine();

                sb.AppendLine("CONTEXT INFORMATION");
                sb.AppendLine("--------------------------------------------------");
                sb.AppendLine($"Context: {context ?? "Not specified"}");
                sb.AppendLine($"Current Operation: {GetCurrentOperation()}");
                if (doc != null)
                {
                    sb.AppendLine($"Document: {doc.Title}");
                    sb.AppendLine($"Document GUID: {doc.WorksharingCentralGUID}");
                    sb.AppendLine($"Document Path: {doc.PathName}");
                }
                sb.AppendLine();

                sb.AppendLine("EXCEPTION DETAILS");
                sb.AppendLine("--------------------------------------------------");
                AppendExceptionDetails(sb, ex);
                sb.AppendLine();

                sb.AppendLine("SYSTEM INFORMATION");
                sb.AppendLine("--------------------------------------------------");
                sb.AppendLine($"OS: {Environment.OSVersion}");
                sb.AppendLine($"64-bit OS: {Environment.Is64BitOperatingSystem}");
                sb.AppendLine($"Machine Name: {Environment.MachineName}");
                sb.AppendLine($"Processor Count: {Environment.ProcessorCount}");
                sb.AppendLine($"CLR Version: {Environment.Version}");
                sb.AppendLine($"Working Set: {Environment.WorkingSet / (1024 * 1024)} MB");

                // Save the crash report to file
                File.WriteAllText(crashFilePath, sb.ToString());
                _currentCrashFilePath = crashFilePath;

                // Also log the exception
                LogException(ex, context, doc);

                return crashFilePath;
            }
            catch (Exception reportEx)
            {
                Debug.WriteLine("Error generating crash report: " + reportEx.Message);
                return null;
            }
        }

        /// <summary>
        /// Shows a dialog with crash information and options to view the report or send it
        /// </summary>
        public static void ShowCrashDialog(Exception ex, string context = null, Document doc = null)
        {
            try
            {
                string crashFilePath = GenerateCrashReport(ex, context, doc);

                TaskDialog dialog = new TaskDialog("RTS Add-in Error")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                    MainInstruction = "The RTS Add-in encountered an error",
                    MainContent = $"Error: {ex.Message}\n\n" +
                                  $"A crash report has been saved to:\n{crashFilePath}\n\n" +
                                  $"Would you like to view the crash report or send it to support?",
                    CommonButtons = TaskDialogCommonButtons.Close,
                    AllowCancellation = true
                };

                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "View Crash Report");
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Open Log Folder");
#if !REVIT2022
                // Email support option only available in Revit 2024+
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Email Support");
#endif

                TaskDialogResult result = dialog.Show();

                if (result == TaskDialogResult.CommandLink1 && !string.IsNullOrEmpty(crashFilePath))
                {
                    Process.Start(crashFilePath);
                }
                else if (result == TaskDialogResult.CommandLink2)
                {
                    Process.Start(LogDirectory);
                }
#if !REVIT2022
                else if (result == TaskDialogResult.CommandLink3)
                {
                    // Compose email to support with crash report attached
                    string subject = $"RTS Add-in Crash Report - {DateTime.Now:yyyy-MM-dd}";
                    string body = $"RTS Add-in Version: {Assembly.GetExecutingAssembly().GetName().Version}\n" +
                                  $"Revit Version: {_revitVersion}\n" +
                                  $"Error: {ex.Message}\n\n" +
                                  "Please describe what you were doing when the error occurred:";

                    string mailtoUrl = $"mailto:support@retick.com.au?subject={Uri.EscapeDataString(subject)}" +
                                       $"&body={Uri.EscapeDataString(body)}";

                    Process.Start(mailtoUrl);

                    // Show a follow-up dialog to remind about attaching the file
                    TaskDialog attachReminder = new TaskDialog("Attach Crash Report")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                        MainInstruction = "Please attach the crash report to your email",
                        MainContent = $"The crash report is located at:\n{crashFilePath}\n\n" +
                                      "Would you like to open the folder to find the file?",
                        CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                    };

                    if (attachReminder.Show() == TaskDialogResult.Yes)
                    {
                        Process.Start(LogDirectory);
                    }
                }
#endif
            }
            catch (Exception dialogEx)
            {
                // Last resort error handling
                Debug.WriteLine("Error showing crash dialog: " + dialogEx.Message);
                TaskDialog.Show("Error", "An error occurred while displaying the crash dialog.\n\n" +
                                         $"Original error: {ex.Message}\n" +
                                         $"Dialog error: {dialogEx.Message}");
            }
        }

        /// <summary>
        /// Track the start of an operation for better context in error reports
        /// </summary>
        public static void StartOperation(string operationName)
        {
            int threadId = Thread.CurrentThread.ManagedThreadId;
            lock (_activeOperations)
            {
                _activeOperations[threadId] = operationName;
            }
            LogMessage(LogLevel.Debug, $"Started operation: {operationName}");
        }

        /// <summary>
        /// End tracking of the current operation
        /// </summary>
        public static void EndOperation()
        {
            int threadId = Thread.CurrentThread.ManagedThreadId;
            string operationName;
            lock (_activeOperations)
            {
                if (_activeOperations.TryGetValue(threadId, out operationName))
                {
                    _activeOperations.Remove(threadId);
                    // The completion message is now handled in ExecuteWithDiagnostics
                }
            }
        }

        /// <summary>
        /// Execute an action with diagnostic tracking, including performance timing.
        /// </summary>
        public static void ExecuteWithDiagnostics(string operationName, Action action)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                StartOperation(operationName);
                action();
            }
            catch (Exception ex)
            {
                LogException(ex, operationName);
                throw; // Re-throw to allow caller to handle
            }
            finally
            {
                stopwatch.Stop();
                LogMessage(LogLevel.Debug, $"Operation '{operationName}' completed in {stopwatch.ElapsedMilliseconds}ms.");
                EndOperation();
            }
        }

        /// <summary>
        /// Execute a function with diagnostic tracking and return its result, including performance timing.
        /// </summary>
        public static T ExecuteWithDiagnostics<T>(string operationName, Func<T> function)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                StartOperation(operationName);
                return function();
            }
            catch (Exception ex)
            {
                LogException(ex, operationName);
                throw; // Re-throw to allow caller to handle
            }
            finally
            {
                stopwatch.Stop();
                LogMessage(LogLevel.Debug, $"Operation '{operationName}' completed in {stopwatch.ElapsedMilliseconds}ms.");
                EndOperation();
            }
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// A lightweight, failsafe logging method that has no dependencies.
        /// It's used to log the earliest stages of add-in loading before the main logger is initialized.
        /// </summary>
        /// <param name="message">The message to log.</param>
        private static void FailsafeLog(string message)
        {
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), LOG_DIRECTORY);
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                string logFilePath = Path.Combine(logDir, string.Format(LOG_FILE_NAME, DateTime.Now.ToString("yyyyMMdd")));
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [FAILSAFE] {message}{Environment.NewLine}";
                File.AppendAllText(logFilePath, logMessage);
            }
            catch
            {
                // Suppress all errors; this is a last resort logger.
            }
        }

        /// <summary>
        /// Internal core initialization logic.
        /// </summary>
        private static void InitializeInternal(string revitVersion)
        {
            try
            {
                // Create log directory if it doesn't exist
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }

                // Set up unhandled exception handler for the AppDomain
                AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
                    LogUnhandledException(args.ExceptionObject as Exception, "AppDomain.UnhandledException");

                // Cache Revit version
                _revitVersion = revitVersion ?? "Unknown";

                // Clean up old log files
                CleanupOldLogFiles();

                // Log successful initialization
                LogMessage(LogLevel.Info, "DiagnosticsManager initialized successfully. Revit version: " + _revitVersion);
            }
            catch (Exception ex)
            {
                // Last resort error handling if we can't set up the diagnostics
                Debug.WriteLine("Failed to initialize DiagnosticsManager: " + ex.Message);
                TaskDialog.Show("Diagnostics Initialization Failed",
                    "Failed to initialize diagnostic services: " + ex.Message);
            }
        }


        /// <summary>
        /// Write a message to the log file
        /// </summary>
        private static void WriteToLogFile(string message)
        {
            string logFilePath = Path.Combine(LogDirectory,
                string.Format(LOG_FILE_NAME, DateTime.Now.ToString("yyyyMMdd")));

            // Use a lock to prevent multiple threads writing at the same time
            lock (typeof(DiagnosticsManager))
            {
                File.AppendAllText(logFilePath, message + Environment.NewLine);
            }
        }

        /// <summary>
        /// Clean up old log files to prevent disk space issues
        /// </summary>
        private static void CleanupOldLogFiles()
        {
            try
            {
                // Get all log files
                var logFiles = Directory.GetFiles(LogDirectory, "RTS_Log_*.txt")
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .ToList();

                // Keep only the most recent MAX_LOG_FILES files
                if (logFiles.Count > MAX_LOG_FILES)
                {
                    foreach (var file in logFiles.Skip(MAX_LOG_FILES))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // Ignore errors when cleaning up
                        }
                    }
                }

                // Also clean up crash reports older than 30 days
                var crashFiles = Directory.GetFiles(LogDirectory, "RTS_Crash_Report_*.txt");
                foreach (var file in crashFiles)
                {
                    if (File.GetLastWriteTime(file) < DateTime.Now.AddDays(-30))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // Ignore errors when cleaning up
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error cleaning up old log files: " + ex.Message);
            }
        }

        /// <summary>
        /// Get the current operation for the calling thread
        /// </summary>
        private static string GetCurrentOperation()
        {
            int threadId = Thread.CurrentThread.ManagedThreadId;
            lock (_activeOperations)
            {
                string operation;
                return _activeOperations.TryGetValue(threadId, out operation)
                    ? operation
                    : "Unknown";
            }
        }

        /// <summary>
        /// Append detailed exception information to a StringBuilder
        /// </summary>
        private static void AppendExceptionDetails(StringBuilder sb, Exception ex, int level = 0)
        {
            if (ex == null) return;

            // Indent based on level for inner exceptions
            string indent = new string(' ', level * 2);

            sb.AppendLine($"{indent}Exception Type: {ex.GetType().FullName}");
            sb.AppendLine($"{indent}Message: {ex.Message}");
            sb.AppendLine($"{indent}Source: {ex.Source}");

            // Add Revit-specific exception details if available
            if (ex is Autodesk.Revit.Exceptions.ApplicationException revitEx)
            {
                sb.AppendLine($"{indent}Revit API Error: {revitEx.Message}");
#if !REVIT2022
                // Revit 2024+ has additional exception data
                sb.AppendLine($"{indent}Severity: {revitEx.GetType().Name}");
#endif
            }

            // Format the stack trace with line numbers if available
            if (!string.IsNullOrEmpty(ex.StackTrace))
            {
                sb.AppendLine($"{indent}Stack Trace:");
                string[] frames = ex.StackTrace.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                // Limit the number of stack frames to keep the log manageable
                int frameCount = Math.Min(frames.Length, MAX_STACK_FRAMES);
                for (int i = 0; i < frameCount; i++)
                {
                    sb.AppendLine($"{indent}  {frames[i].Trim()}");
                }

                if (frames.Length > MAX_STACK_FRAMES)
                {
                    sb.AppendLine($"{indent}  ... {frames.Length - MAX_STACK_FRAMES} more frames ...");
                }
            }

            // Include additional exception data
            if (ex.Data.Count > 0)
            {
                sb.AppendLine($"{indent}Additional Data:");
                foreach (System.Collections.DictionaryEntry entry in ex.Data)
                {
                    sb.AppendLine($"{indent}  {entry.Key}: {entry.Value}");
                }
            }

            // Recursively process inner exceptions
            if (ex.InnerException != null)
            {
                sb.AppendLine($"{indent}Inner Exception:");
                AppendExceptionDetails(sb, ex.InnerException, level + 1);
            }

            // Handle aggregate exceptions (multiple inner exceptions)
            if (ex is AggregateException aggEx)
            {
                sb.AppendLine($"{indent}Aggregate Exceptions ({aggEx.InnerExceptions.Count}):");
                int innerIndex = 0;
                foreach (var innerEx in aggEx.InnerExceptions)
                {
                    sb.AppendLine($"{indent}Inner Exception #{++innerIndex}:");
                    AppendExceptionDetails(sb, innerEx, level + 1);
                }
            }
        }

        /// <summary>
        /// Handler for unhandled exceptions at the AppDomain level
        /// </summary>
        private static void LogUnhandledException(Exception ex, string source)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"UNHANDLED EXCEPTION from {source} at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Current Operation: {GetCurrentOperation()}");
                sb.AppendLine($"Revit Version: {_revitVersion}");

                AppendExceptionDetails(sb, ex);

                // Write to crash report file
                string crashFilePath = Path.Combine(LogDirectory,
                    string.Format(CRASH_FILE_NAME, DateTime.Now.ToString("yyyyMMdd_HHmmss")));
                File.WriteAllText(crashFilePath, sb.ToString());

                // Also log to regular log file
                WriteToLogFile(sb.ToString());

                Debug.WriteLine(sb.ToString());
            }
            catch (Exception logEx)
            {
                // Last resort if we can't even log the exception
                Debug.WriteLine($"FATAL: Failed to log unhandled exception: {logEx.Message}");
                Debug.WriteLine($"Original exception: {ex?.Message}");
            }
        }

        #endregion
    }
}
