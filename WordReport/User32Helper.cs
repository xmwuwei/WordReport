namespace WordReport
{
    using System;
    using System.Text;

    public static class User32Helper
    {
        public static int EnumWindows(EnumWindowsEvent enumWindowsEvent, int lParam) =>
            User32.EnumWindows(enumWindowsEvent, lParam);

        public static IntPtr FindWindow(string className, string windowName) =>
            User32.FindWindow(className, windowName);

        public static int GetWindowText(IntPtr handle, StringBuilder text, int MaxLen) =>
            User32.GetWindowText(handle, text, MaxLen);
    }
}
