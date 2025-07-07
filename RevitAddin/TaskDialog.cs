using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace RTS
{
    /// <summary>
    /// TaskDialog.cs file
    /// A custom implementation of a Task Dialog using Windows API P/Invoke.
    /// This removes the dependency on the WindowsAPICodePack library to avoid assembly conflicts in Revit.
    /// </summary>
    public static class TaskDialog
    {
        // P/Invoke declaration for the TaskDialogIndirect function from comctl32.dll
        [DllImport("comctl32.dll", CharSet = CharSet.Unicode, EntryPoint = "TaskDialogIndirect")]
        private static extern int TaskDialogIndirect(
            [In] ref TASKDIALOGCONFIG pTaskConfig,
            [Out] out int pnButton,
            [Out] out int pnRadioButton,
            [Out] out bool pfVerificationFlagChecked);

        // Main method to show the dialog
        public static void Show(string text, string instruction = "", string title = "Revit")
        {
            TASKDIALOGCONFIG config = new TASKDIALOGCONFIG
            {
                cbSize = (uint)Marshal.SizeOf(typeof(TASKDIALOGCONFIG)),
                hwndParent = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle,
                dwFlags = TASKDIALOG_FLAGS.TDF_ALLOW_DIALOG_CANCELLATION,
                dwCommonButtons = TASKDIALOG_COMMON_BUTTON_FLAGS.TDCBF_OK_BUTTON,
                pszWindowTitle = title,
                pszMainInstruction = instruction,
                pszContent = text,
                // CORRECTED: The field name is hMainIcon, not pszMainIcon, for an icon handle.
                hMainIcon = new IntPtr((int)TD_ICON.TD_INFORMATION_ICON)
            };

            TaskDialogIndirect(ref config, out _, out _, out _);
        }

        // Structures and enums that map to the Windows API for Task Dialogs

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
            public IntPtr hMainIcon; // The correct field for the main icon handle
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

        [Flags]
        private enum TASKDIALOG_FLAGS : uint
        {
            TDF_ENABLE_HYPERLINKS = 0x0001,
            TDF_USE_HICON_MAIN = 0x0002,
            TDF_USE_HICON_FOOTER = 0x0004,
            TDF_ALLOW_DIALOG_CANCELLATION = 0x0008,
            TDF_USE_COMMAND_LINKS = 0x0010,
            TDF_USE_COMMAND_LINKS_NO_ICON = 0x0020,
            TDF_EXPAND_FOOTER_AREA = 0x0040,
            TDF_EXPANDED_BY_DEFAULT = 0x0080,
            TDF_VERIFICATION_FLAG_CHECKED = 0x0100,
            TDF_SHOW_PROGRESS_BAR = 0x0200,
            TDF_SHOW_MARQUEE_PROGRESS_BAR = 0x0400,
            TDF_CALLBACK_TIMER = 0x0800,
            TDF_POSITION_RELATIVE_TO_WINDOW = 0x1000,
            TDF_RTL_LAYOUT = 0x2000,
            TDF_NO_DEFAULT_RADIO_BUTTON = 0x4000,
            TDF_CAN_BE_MINIMIZED = 0x8000
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

        private enum TD_ICON : int
        {
            TD_WARNING_ICON = -1,
            TD_ERROR_ICON = -2,
            TD_INFORMATION_ICON = -3,
            TD_SHIELD_ICON = -4
        }
    }
}
