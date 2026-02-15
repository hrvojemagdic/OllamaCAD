using SolidWorks.Interop.sldworks;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace OllamaCAD
{   
    /// <summary>
    /// Captures a screenshot of the active SOLIDWORKS window and returns it as PNG bytes.
    ///
    /// Behavior:
    /// - Retrieves the SOLIDWORKS main window handle (HWND).
    /// - First attempts capture using the Win32 PrintWindow API (preferred).
    /// - Falls back to Graphics.CopyFromScreen if PrintWindow fails.
    /// - Returns PNG-encoded byte array for use in vision model prompts.
    ///
    /// Used when "Include screenshot in prompt" is enabled.
    /// </summary>
    internal static class ScreenshotHelper
    {   
        /// <summary>
        /// Captures the SOLIDWORKS main window and returns PNG bytes.
        /// Tries PrintWindow first, then CopyFromScreen as fallback.
        /// </summary>
        public static byte[] CaptureSolidWorksWindowPngBytes(ISldWorks swApp)
        {
            int hwnd = 0;
            try { hwnd = swApp.IFrameObject().GetHWnd(); } catch { }

            if (hwnd == 0) return null;

            var bytes = TryPrintWindowPng((IntPtr)hwnd);
            if (bytes != null && bytes.Length > 0)
                return bytes;

            return TryCopyFromScreenPng((IntPtr)hwnd);
        }

        /// <summary>
        /// Attempts to capture the window using the Win32 PrintWindow API.
        /// </summary>
        private static byte[] TryPrintWindowPng(IntPtr hwnd)
        {
            try
            {
                RECT rc;
                if (!GetWindowRect(hwnd, out rc))
                    return null;

                int width = rc.Right - rc.Left;
                int height = rc.Bottom - rc.Top;
                if (width <= 0 || height <= 0)
                    return null;

                using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                using (var gfx = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = gfx.GetHdc();
                    try
                    {
                        bool ok = PrintWindow(hwnd, hdc, 0);
                        if (!ok) return null;
                    }
                    finally
                    {
                        gfx.ReleaseHdc(hdc);
                    }

                    using (var ms = new MemoryStream())
                    {
                        bmp.Save(ms, ImageFormat.Png);
                        return ms.ToArray();
                    }
                }
            }
            catch { return null; }
        }


        /// <summary>
        /// Fallback screen capture using Graphics.CopyFromScreen.
        /// </summary>
        private static byte[] TryCopyFromScreenPng(IntPtr hwnd)
        {
            try
            {
                RECT rc;
                if (!GetWindowRect(hwnd, out rc))
                    return null;

                int width = rc.Right - rc.Left;
                int height = rc.Bottom - rc.Top;
                if (width <= 0 || height <= 0)
                    return null;

                using (var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(rc.Left, rc.Top, 0, 0, new Size(width, height));
                    using (var ms = new MemoryStream())
                    {
                        bmp.Save(ms, ImageFormat.Png);
                        return ms.ToArray();
                    }
                }
            }
            catch { return null; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
    }
}
