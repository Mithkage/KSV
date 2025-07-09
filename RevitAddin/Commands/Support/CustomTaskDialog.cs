//
// Copyright (c) 2024 [Your Name/Company Here]
//
//
// File:         CustomTaskDialog.cs
//
// Description:  A feature-rich, self-contained implementation of a Windows Task Dialog.
//               This class uses P/Invoke to call the native Windows API (comctl32.dll),
//               which avoids the need for external libraries like the WindowsAPICodePack.
//               This prevents assembly binding and versioning conflicts that can occur
//               within the Revit environment. It supports command links, common buttons,
//               and detailed content.
//
// Author:       Kyle Vorster (Modified by AI)
//
// Change Log:
// 2025-07-09:   Major rewrite. Converted from a simple static class to a full-featured
//               instance-based class to support command links, common buttons, and
//               detailed configurations, removing the dependency on Autodesk.Revit.UI.TaskDialog.
// 2025-07-09:   Initial creation. Renamed from TaskDialog.cs to CustomTaskDialog.cs
//               to prevent ambiguous reference conflicts.
//
//

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace RTS.Commands.Support
{
    #region Public Enums (for API consistency)

    /// <summary>
    /// Defines standard results returned from a Task Dialog.
    /// </summary>
    public enum CustomTaskDialogResult
    {
        None,
        Ok,
        Cancel,
        Yes,
        No,
        Retry,
        Close,
        CommandLink1,
        CommandLink2,
        CommandLink3,
        CommandLink4
    }

    /// <summary>
    /// Defines the IDs for command links.
    /// </summary>
    public enum CustomTaskDialogCommandLinkId
    {
        CommandLink1 = 1001,
        CommandLink2 = 1002,
        CommandLink3 = 1003,
        CommandLink4 = 1004,
    }

    /// <summary>
    /// Defines standard buttons that can be added to a Task Dialog.
    /// </summary>
    [Flags]
    public enum CustomTaskDialogCommonButtons
    {
        None = 0,
        Ok = 0x0001,
        Yes = 0x0002,
        No = 0x0004,
        Cancel = 0x0008,
        Retry = 0x0010,
        Close = 0x0020
    }

    #endregion

    /// <summary>
    /// A custom implementation of a Task Dialog using Windows API P/Invoke.
    /// </summary>
    public class CustomTaskDialog
    {
        #region Public Properties

        public string Title { get; set; }
        public string MainInstruction { get; set; }
        public string MainContent { get; set; }
        public CustomTaskDialogCommonButtons CommonButtons { get; set; }
        public CustomTaskDialogResult DefaultButton { get; set; }

        private readonly List<TASKDIALOG_BUTTON> _commandLinks = new List<TASKDIALOG_BUTTON>();

        #endregion

        #region Constructor

        public CustomTaskDialog(string title)
        {
            Title = title;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Adds a command link to the dialog.
        /// </summary>
        public void AddCommandLink(CustomTaskDialogCommandLinkId id, string mainText, string supportingText = "")
        {
            _commandLinks.Add(new TASKDIALOG_BUTTON { nButtonID = (int)id, pszButtonText = $"{mainText}\n{supportingText}" });
        }

        /// <summary>
        /// Displays the configured task dialog.
        /// </summary>
        public CustomTaskDialogResult Show()
        {
            var config = new TASKDIALOGCONFIG();
            config.cbSize = (uint)Marshal.SizeOf(typeof(TASKDIALOGCONFIG));
            config.hwndParent = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            config.dwFlags = TASKDIALOG_FLAGS.TDF_ALLOW_DIALOG_CANCELLATION | TASKDIALOG_FLAGS.TDF_POSITION_RELATIVE_TO_WINDOW;
            if (_commandLinks.Count > 0)
            {
                config.dwFlags |= TASKDIALOG_FLAGS.TDF_USE_COMMAND_LINKS;
            }

            config.dwCommonButtons = (TASKDIALOG_COMMON_BUTTON_FLAGS)CommonButtons;
            config.pszWindowTitle = Title;
            config.pszMainInstruction = MainInstruction;
            config.pszContent = MainContent;

            // Handle command links
            var commandButtons = _commandLinks.ToArray();
            var pButtons = IntPtr.Zero;
            GCHandle handle = new GCHandle();

            if (commandButtons.Length > 0)
            {
                handle = GCHandle.Alloc(commandButtons, GCHandleType.Pinned);
                pButtons = handle.AddrOfPinnedObject();
                config.cButtons = (uint)commandButtons.Length;
                config.pButtons = pButtons;
            }

            // Set default button
            config.nDefaultButton = MapResultToButtonId(DefaultButton);

            try
            {
                int buttonResult;
                TaskDialogIndirect(ref config, out buttonResult, out _, out _);
                return MapButtonIdToResult(buttonResult);
            }
            finally
            {
                if (handle.IsAllocated)
                {
                    handle.Free();
                }
            }
        }

        /// <summary>
        /// Shows a simple task dialog with a title and main content.
        /// </summary>
        public static void Show(string title, string mainContent)
        {
            var dialog = new CustomTaskDialog(title)
            {
                MainContent = mainContent,
                CommonButtons = CustomTaskDialogCommonButtons.Ok
            };
            dialog.Show();
        }

        #endregion

        #region Private Helper Methods

        private int MapResultToButtonId(CustomTaskDialogResult result)
        {
            switch (result)
            {
                case CustomTaskDialogResult.Ok: return 1;
                case CustomTaskDialogResult.Cancel: return 2;
                case CustomTaskDialogResult.Yes: return 6;
                case CustomTaskDialogResult.No: return 7;
                case CustomTaskDialogResult.Retry: return 4;
                case CustomTaskDialogResult.Close: return 8;
                case CustomTaskDialogResult.CommandLink1: return (int)CustomTaskDialogCommandLinkId.CommandLink1;
                case CustomTaskDialogResult.CommandLink2: return (int)CustomTaskDialogCommandLinkId.CommandLink2;
                case CustomTaskDialogResult.CommandLink3: return (int)CustomTaskDialogCommandLinkId.CommandLink3;
                case CustomTaskDialogResult.CommandLink4: return (int)CustomTaskDialogCommandLinkId.CommandLink4;
                default: return 0;
            }
        }

        private CustomTaskDialogResult MapButtonIdToResult(int buttonId)
        {
            switch (buttonId)
            {
                case 1: return CustomTaskDialogResult.Ok;
                case 2: return CustomTaskDialogResult.Cancel;
                case 6: return CustomTaskDialogResult.Yes;
                case 7: return CustomTaskDialogResult.No;
                case 4: return CustomTaskDialogResult.Retry;
                case 8: return CustomTaskDialogResult.Close;
                case (int)CustomTaskDialogCommandLinkId.CommandLink1: return CustomTaskDialogResult.CommandLink1;
                case (int)CustomTaskDialogCommandLinkId.CommandLink2: return CustomTaskDialogResult.CommandLink2;
                case (int)CustomTaskDialogCommandLinkId.CommandLink3: return CustomTaskDialogResult.CommandLink3;
                case (int)CustomTaskDialogCommandLinkId.CommandLink4: return CustomTaskDialogResult.CommandLink4;
                default: return CustomTaskDialogResult.None;
            }
        }

        #endregion

        #region P/Invoke Definitions

        [DllImport("comctl32.dll", CharSet = CharSet.Unicode, EntryPoint = "TaskDialogIndirect")]
        private static extern int TaskDialogIndirect(
            [In] ref TASKDIALOGCONFIG pTaskConfig,
            [Out] out int pnButton,
            [Out] out int pnRadioButton,
            [Out] out bool pfVerificationFlagChecked);

        [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Unicode)]
        private struct TASKDIALOGCONFIG
        {
            public uint cbSize;
            public IntPtr hwndParent;
            public IntPtr hInstance;
            public TASKDIALOG_FLAGS dwFlags;
            public TASKDIALOG_COMMON_BUTTON_FLAGS dwCommonButtons;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszWindowTitle;
            public IntPtr hMainIcon;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszMainInstruction;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszContent;
            public uint cButtons;
            public IntPtr pButtons;
            public int nDefaultButton;
            public uint cRadioButtons;
            public IntPtr pRadioButtons;
            public int nDefaultRadioButton;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszVerificationText;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszExpandedInformation;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszExpandedControlText;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszCollapsedControlText;
            public IntPtr hFooterIcon;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszFooter;
            public IntPtr pfCallback;
            public IntPtr lpCallbackData;
            public uint cxWidth;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Unicode)]
        private struct TASKDIALOG_BUTTON
        {
            public int nButtonID;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszButtonText;
        }

        [Flags]
        private enum TASKDIALOG_FLAGS : uint
        {
            TDF_ALLOW_DIALOG_CANCELLATION = 0x0008,
            TDF_USE_COMMAND_LINKS = 0x0010,
            TDF_POSITION_RELATIVE_TO_WINDOW = 0x1000,
        }

        [Flags]
        private enum TASKDIALOG_COMMON_BUTTON_FLAGS : uint
        {
            TDCBF_OK_BUTTON = 0x0001,
            TDCBF_YES_BUTTON = 0x0002,
            TDCBF_NO_BUTTON = 0x0004,
            TDCBF_CANCEL_BUTTON = 0x0008,
            TDCBF_RETRY_BUTTON = 0x0010,
            TDCBF_CLOSE_BUTTON = 0x0020
        }

        #endregion
    }
}
