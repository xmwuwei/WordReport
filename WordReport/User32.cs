namespace WordReport
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    internal static class User32
    {
        private const string USER32 = "user32.dll";

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        internal static extern int EnumWindows(EnumWindowsEvent enumWindowsEvent, int lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr FindWindow(string className, string windowName);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        internal static extern int GetWindowText(IntPtr handle, StringBuilder text, int MaxLen);
    }
}
