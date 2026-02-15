using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace OllamaCAD
{   
    /// <summary>
    /// Captures the active SOLIDWORKS main window and returns
    /// the image as a Base64-encoded PNG string.
    ///
    /// - Uses the Win32 PrintWindow API to capture the window content.
    /// - Converts the bitmap to PNG format in memory.
    /// - Returns Base64 string suitable for vision-capable LLM prompts.
    ///
    /// Intended for lightweight screenshot embedding without saving to disk.
    /// </summary>
    internal static class SwScreenshot
    {
        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, int nFlags);

        /// <summary>
        /// Captures the active SOLIDWORKS window and returns it as a Base64 PNG string.
        /// Returns null if no valid window handle is available.
        /// </summary>
        public static string CaptureActiveModelViewBase64Png(ISldWorks app)
        {
            IntPtr hwnd = new IntPtr((long)app.IFrameObject().GetHWnd());
            if (hwnd == IntPtr.Zero) return null;

            using (Bitmap bmp = new Bitmap(1400, 900))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = g.GetHdc();
                    try
                    {
                        PrintWindow(hwnd, hdc, 0);
                    }
                    finally
                    {
                        g.ReleaseHdc(hdc);
                    }
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Png);
                    return Convert.ToBase64String(ms.ToArray());
                }
            }
        }
    }
}