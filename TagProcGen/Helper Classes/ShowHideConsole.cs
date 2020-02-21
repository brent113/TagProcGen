using System;
using System.Runtime.InteropServices;

namespace TagProcGen
{
    internal static class NativeMethods
    {
        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        internal const int SW_HIDE = 0;
        internal const int SW_SHOW = 5;
    }

    /// <summary>
    /// Show or Hide the application's console
    /// </summary>
    public static class ShowHideConsole
    {
        static readonly IntPtr handle = NativeMethods.GetConsoleWindow();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Visible">Sets visibility of the console</param>
        public static void SetConsoleVisibility(bool Visible)
        {
            if (Visible)
            {
                NativeMethods.ShowWindow(handle, NativeMethods.SW_SHOW);
            }
            else
            {
                NativeMethods.ShowWindow(handle, NativeMethods.SW_HIDE);
            }
        }
    }
}
