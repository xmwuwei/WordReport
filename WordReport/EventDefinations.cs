namespace WordReport
{
    using System;
    using System.Runtime.CompilerServices;

    public delegate bool EnumWindowsEvent(IntPtr hWnd, int y);
    public delegate void ShowProgressEvent(ReportStatus reportStatus, int nProgressStep);

}
