using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AIFileSorterShellExtension
{
    /// <summary>
    /// Manages cursor in Windows Explorer context
    /// </summary>
    public static class CursorManager
    {
        // Constants for Win32 API
        private const int OCR_NORMAL = 32512;
        private const int OCR_WAIT = 32514;
        private const uint SPI_SETCURSORS = 0x0057;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDCHANGE = 0x02;

        [DllImport("user32.dll")]
        private static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);

        [DllImport("user32.dll")]
        private static extern IntPtr SetCursor(IntPtr hCursor);

        [DllImport("user32.dll")]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

        [DllImport("user32.dll")]
        private static extern bool GetCursorInfo(ref CURSORINFO pci);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetSystemCursor(IntPtr hcur, uint id);

        [DllImport("user32.dll")]
        private static extern IntPtr CopyIcon(IntPtr hIcon);

        [StructLayout(LayoutKind.Sequential)]
        private struct CURSORINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hCursor;
            public POINTAPI ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINTAPI
        {
            public int x;
            public int y;
        }

        /// <summary>
        /// Temporarily replaces system cursor with hourglass (wait) cursor
        /// </summary>
        /// <returns>IDisposable object for cursor restoration</returns>
        public static IDisposable ShowWaitCursor()
        {
            IntPtr waitCursor = LoadCursor(IntPtr.Zero, OCR_WAIT);
            SetSystemCursor(CopyIcon(waitCursor), OCR_NORMAL);
            SystemParametersInfo(SPI_SETCURSORS, 0, IntPtr.Zero, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
            
            return new CursorRestorer();
        }

        /// <summary>
        /// Restores system cursor to its original state
        /// </summary>
        public static void RestoreDefaultCursor()
        {
            SystemParametersInfo(SPI_SETCURSORS, 0, IntPtr.Zero, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }

        private class CursorRestorer : IDisposable
        {
            public void Dispose()
            {
                RestoreDefaultCursor();
            }
        }
    }
}
