using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Timers;

namespace tud.mci.tangram.TangramLector
{
    public class CaptureChangedEventArgs : EventArgs
    {
        /// <summary>
        /// The captured image
        /// </summary>
        /// <value>The screen capture.</value>
        public Image Img { get; private set; }
        public CaptureChangedEventArgs(Image image)
        {
            Img = image;
        }
    }

    /// <summary>
    /// A ScreenObserver can capture the screen, part of the screen ore windows as well as part of windows in an continuous updating manner.
    /// </summary>
    public class ScreenObserver
    {
        #region Members
        double refreshRate;
        Timer refreshTimer;
        //IntPtr Whnd = IntPtr.Zero;
        public IntPtr Whnd { get; private set; }
        public Object ScreenPos { private set; get; }
        Image _lastCap;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver"/> class.
        /// This Observer generates screen shots of the main screen desktop window.
        /// </summary>
        /// <param name="interval">The capturing interval in milliseconds.</param>
        public ScreenObserver(double interval)
        {
            if (interval <= 0) throw new ArgumentException("Timer interval has to be higher than 0", "interval");

            refreshRate = interval;
            refreshTimer = new Timer(interval);
            refreshTimer.Elapsed += new ElapsedEventHandler(refreshTimer_Elapsed);
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver"/> class.
        /// This Observer generates screen shots of the given application/window.
        /// </summary>
        /// <param name="interval">The capturing interval in milliseconds.</param>
        /// <param name="wHnd">The window handle of the window/application to observe.</param>
        public ScreenObserver(double interval, IntPtr wHnd) : this(interval) { Whnd = wHnd; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver"/> class.
        /// This Observer generates screen shots of the screen position.
        /// </summary>
        /// <param name="interval">The capturing interval in milliseconds.</param>
        /// <param name="screenPosition">The position on the screen.</param>
        public ScreenObserver(double interval, Rectangle screenPosition) : this(interval) { ScreenPos = screenPosition; }
        #endregion

        #region Event
        public delegate void CaptureChangedEventHandler(object sender, CaptureChangedEventArgs e);
        public event CaptureChangedEventHandler Changed;
        #endregion

        #region Timer
        void refreshTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (Changed != null)
            {
                using (Image sc = capture())
                {
                    if (sc != null && !sc.Equals(_lastCap))
                    {
                        _lastCap = sc;

                        try
                        {
                            using (Bitmap img = sc.Clone() as Bitmap)
                            {
                                Changed.Invoke(this, new CaptureChangedEventArgs(img));
                            }
                        }
                        catch { }
                        finally { 
                           /* FIXME: prevent Memory Exceptions when no one is taking over 
                            * the produced image. The GC is never called at all until the 
                            * memory runs out!
                            * So we have to force the GC to do his job
                            * */
                            GC.Collect(); 
                        }
                    }
                    else
                    {
                        //TODO: what happens if the Image is null?
                    }
                }
            }
        }

        /// <summary>
        /// Starts this observation.
        /// </summary>
        /// <returns>true if the observer could be started otherwise false.</returns>
        public bool Start()
        {
            if (refreshRate > 0 && refreshTimer != null)
            {
                refreshTimer.Start();
                return refreshTimer.Enabled = true;
            }
            return false;
        }

        /// <summary>
        /// Stops the observation.
        /// </summary>
        public void Stop() { if (refreshTimer != null) refreshTimer.Enabled = false; refreshTimer.Stop(); }

        /// <summary>
        /// Sets the interval of the observer.
        /// </summary>
        /// <param name="interval">The interval (must be largen then 0).</param>
        /// <returns>true if the interval could be changed otherwise false.</returns>
        public bool SetObserverInterval(double interval)
        {
            if (interval <= 0) throw new ArgumentException("Timer interval has to be higher than 0", "interval");
            if (refreshTimer != null)
            {
                refreshTimer.Interval = interval;
                return true;
            }
            return false;
        }
        #endregion

        #region Screencapturing
        /// <summary>
        /// Sets a window handle to capture.
        /// </summary>
        /// <param name="wHnd">The window handle.</param>
        public void SetWhnd(IntPtr wHnd)
        {
            if (tud.mci.tangram.TangramLector.ScreenCapture.User32.IsWindow(wHnd))
            {
                ScreenPos = null;
                Whnd = wHnd;
            }
        }
        /// <summary>
        /// Sets the screen position to capture.
        /// </summary>
        /// <param name="screenPosition">The screen position.</param>
        public void SetScreenPos(Rectangle screenPosition) { Whnd = IntPtr.Zero; ScreenPos = screenPosition; }
        /// <summary>
        /// Sets the part of window handle to capture.
        /// </summary>
        /// <param name="position">The screen position to capture.</param>
        /// <param name="wHnd">The corresponding window handle to caputure.</param>
        public void SetPartOfWhnd(Rectangle position, IntPtr wHnd)
        {
            if (tud.mci.tangram.TangramLector.ScreenCapture.User32.IsWindow(wHnd))
            { ScreenPos = position; Whnd = wHnd; }
        }
        /// <summary>
        /// Captures the first screen (positive parts).
        /// </summary>
        public void ObserveScreen() { Stop(); Whnd = tud.mci.tangram.TangramLector.ScreenCapture.User32.GetDesktopWindow(); ScreenPos = null; Start(); }

        Image capture()
        {
            Image result = null;
            try
            {
                if (Whnd == IntPtr.Zero)
                {
                    if (ScreenPos != null)
                    {
                        if (ScreenPos is Rectangle)
                        {
                            result = ScreenCapture.CaptureScreenPos((Rectangle)ScreenPos);
                        }
                    }
                    else
                    {
                        result = ScreenCapture.CaptureScreen();
                    }
                }
                else
                {
                    if (ScreenPos != null && ScreenPos is Rectangle && ((Rectangle)ScreenPos).Width > 0 && ((Rectangle)ScreenPos).Height > 0)
                    {
                        Rectangle sp = (Rectangle)ScreenPos;
                        result = ScreenCapture.CaptureWindowPartAtScreenpos(Whnd, sp.Height, sp.Width, sp.X, sp.Y);
                    }
                    else
                    {
                        if (tud.mci.tangram.TangramLector.ScreenCapture.User32.IsWindow(Whnd))
                        {
                            result = ScreenCapture.CaptureWindow(Whnd);
                        }
                        else
                        {
                            SetWhnd(IntPtr.Zero);
                            result = ScreenCapture.CaptureScreen();
                        }
                    }
                }
            }
            catch (Exception) { }
            return result;
        }

        #endregion
    }


    //////////////////////////////////////////////////////////////////////////
    // http://www.developerfusion.com/code/4630/capture-a-screen-shot/
    // Capture a Screen Shot
    // By James Crowley, published on 13 Apr 2004 
    //////////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Provides functions to capture the entire screen, or a particular window, and save it to a file.
    /// </summary>
    public static class ScreenCapture
    {
        /// <summary>
        /// Saves an image to a file.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <param name="img">The img to save.</param>
        /// <param name="format">The file format.</param>
        public static void SaveImageToFile(String filename, Image img, ImageFormat format = null)
        {
            if (img != null && !String.IsNullOrWhiteSpace(filename))
            {
                img.Save(filename, format != null ? format : ImageFormat.Bmp);
            }
        }

        /// <summary>
        /// Captures the screen at a specific position.
        /// </summary>
        /// <param name="capturingBounds">The capturing bounds.</param>
        /// <returns>the resulting image or NULL</returns>
        public static Image CaptureScreenPos(System.Drawing.Rectangle capturingBounds)
        {
            Bitmap image = null;

            if ((capturingBounds.Width > 0) && (capturingBounds.Height > 0))
            {
                image = new Bitmap(capturingBounds.Width, capturingBounds.Height, PixelFormat.Format32bppArgb);
                Graphics gr = Graphics.FromImage(image);
                gr.CopyFromScreen(capturingBounds.X, capturingBounds.Y, 0, 0, new System.Drawing.Size(capturingBounds.Width, capturingBounds.Height), CopyPixelOperation.SourceCopy);
                gr.Dispose();
                return image as Image;
            }
            return null;
        }
        /// <summary>
        /// Creates an Image object containing a screen shot of the entire desktop
        /// </summary>
        /// <returns></returns>
        public static Image CaptureScreen()
        {
            return CaptureWindow(User32.GetDesktopWindow());
        }

        /// <summary>
        /// Creates an Image object containing a screen shot of a specific window
        /// </summary>
        /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        /// <returns></returns>
        public static Image CaptureWindowPartAtScreenpos(IntPtr handle, int height, int width, int nXSrc, int nYSrc)
        {
            // get the hDC of the target window
            //IntPtr hdcSrc = User32.GetWindowDC(handle);
            // get the size
            User32.RECT windowRect = new User32.RECT();
            User32.GetWindowRect(handle, ref windowRect);

            int top = nYSrc - windowRect.top;
            int left = nXSrc - windowRect.left;

            int wWidth = windowRect.right - windowRect.left;
            int wHeight = windowRect.bottom - windowRect.top;

            return CaptureWindow(handle, height, width, left, top);
        }
        /// <summary>
        /// Captures the only the client area of a window.
        /// </summary>
        /// <param name="handle">The window handle.</param>
        /// <returns>the image of the client area.</returns>
        public static Image CaptureWindowClientArea(IntPtr handle)
        {
            // get the hDC of the target window
            //IntPtr hdcSrc = User32.GetWindowDC(handle);
            // get the size
            User32.RECT windowRect = new User32.RECT();
            User32.GetWindowRect(handle, ref windowRect);

            User32.RECT windowClRect = new User32.RECT();
            User32.GetClientRect(handle, ref windowClRect);

            int top = 0;
            int left = 0;

            System.Drawing.Point cPoint = new System.Drawing.Point();
            User32.ClientToScreen(handle, ref cPoint);
            top = cPoint.Y - windowRect.top;
            left = cPoint.X - windowRect.left;

            int wWidth = windowRect.right - windowRect.left;
            int wHeight = windowRect.bottom - windowRect.top;

            return CaptureWindow(handle, windowClRect.bottom, windowClRect.right, left, top);
        }

        /// <summary>
        /// Creates an Image object containing a screen shot of a specific window
        /// </summary>
        /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        /// <returns></returns>
        public static void CaptureWindowPartAtScreenpoToFile(IntPtr handle, string filename, int height, int width, int nXSrc, int nYSrc)
        {
            using (Image img = CaptureWindowPartAtScreenpos(handle, height, width, nXSrc, nYSrc))
            {
                img.Save(filename, ImageFormat.Bmp);
            }
        }

        /// <summary>
        /// Creates an Image object containing a screen shot of a specific window
        /// </summary>
        /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        /// <returns></returns>
        public static Image CaptureWindow(IntPtr handle)
        {
            // get the hDC of the target window
            //IntPtr hdcSrc = User32.GetWindowDC(handle);
            // get the size
            User32.RECT windowRect = new User32.RECT();
            User32.GetWindowRect(handle, ref windowRect);
            int width = windowRect.right - windowRect.left;
            int height = windowRect.bottom - windowRect.top;
            return CaptureWindow(handle, height, width);
        }
        /// <summary>
        /// Creates an Image object containing a screen shot of a specific window
        /// </summary>
        /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        /// <param name="height">The height.</param>
        /// <param name="width">The width.</param>
        /// <param name="nXDest">The x-coordinate, in logical units, of the upper-left corner of the destination rectangle.</param>
        /// <param name="nYDest">The y-coordinate, in logical units, of the upper-left corner of the destination rectangle.</param>
        /// <param name="nXSrc">The x-coordinate, in logical units, of the upper-left corner of the source rectangle.</param>
        /// <param name="nYSrc">The y-coordinate, in logical units, of the upper-left corner of the source rectangle.</param>
        /// <returns></returns>
        public static Image CaptureWindow(IntPtr handle, int height, int width, int nXSrc = 0, int nYSrc = 0, int nXDest = 0, int nYDest = 0)
        {
            if (!tud.mci.tangram.TangramLector.ScreenCapture.User32.IsWindow(handle))
            {
                //TODO: how to handle this
                return new Bitmap(1, 1);
            }

            // get the hDC of the target window
            IntPtr hdcSrc = User32.GetWindowDC(handle);
            // create a device context we can copy to
            IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
            // create a bitmap we can copy it to,
            // using GetDeviceCaps to get the width/height
            IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc, width, height);
            // select the bitmap object
            IntPtr hOld = GDI32.SelectObject(hdcDest, hBitmap);
            // bitblt over
            GDI32.BitBlt(hdcDest, nXDest, nYDest, width, height, hdcSrc, nXSrc, nYSrc, GDI32.SRCCOPY);
            // restore selection
            GDI32.SelectObject(hdcDest, hOld);
            // clean up
            GDI32.DeleteDC(hdcDest);
            User32.ReleaseDC(handle, hdcSrc);
            // get a .NET image object for it
            Image img = null;
            try
            {
                if (hBitmap != null && hBitmap != IntPtr.Zero)
                {
                    img = Image.FromHbitmap(hBitmap);
                }
                else { }
            }
            catch { }
            finally
            {
                // free up the Bitmap object
                GDI32.DeleteObject(hBitmap);
                GDI32.DeleteObject(hdcDest);
                GDI32.DeleteObject(hOld);
                GDI32.DeleteObject(hdcSrc);
            }
            return img;
        }

        /// <summary>
        /// Captures a screen shot of a specific window, and saves it to a file
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="filename"></param>
        public static void CaptureWindowToFile(IntPtr handle, string filename)
        {
            CaptureWindowToFile(handle, filename, ImageFormat.Bmp);
        }
        /// <summary>
        /// Captures a screen shot of a specific window, and saves it to a file
        /// </summary>
        /// <param name="handle">The window handle</param>
        /// <param name="filename">the file name to save</param>
        /// <param name="height">The height.</param>
        /// <param name="width">The width.</param>
        /// <param name="nXSrc">The start X in the sorce.</param>
        /// <param name="nYSrc">The start Y in the source.</param>
        /// <param name="nXDest">The start X in the destination.</param>
        /// <param name="nYDest">The start Y in the destination.</param>
        public static void CaptureWindowToFile(IntPtr handle, string filename, int height, int width, int nXSrc = 0, int nYSrc = 0, int nXDest = 0, int nYDest = 0)
        {
            using (Image img = CaptureWindow(handle, height, width, nXSrc, nXSrc, nXDest, nYDest))
            {
                img.Save(filename, ImageFormat.Bmp);
            }
        }
        /// <summary>
        /// Captures a screen shot of a specific window, and saves it to a file
        /// </summary>
        /// <param name="handle">The window handle</param>
        /// <param name="filename">the file name to save</param>
        /// <param name="format">the image format</param>
        public static void CaptureWindowToFile(IntPtr handle, string filename, ImageFormat format)
        {
            using (Image img = CaptureWindow(handle))
            {
                img.Save(filename, format);
            }
        }
        /// <summary>
        /// Captures a screen shot of a specific window, and saves it to a file
        /// </summary>
        /// <param name="handle">The window handle</param>
        /// <param name="filename">the file name to save</param>
        /// <param name="format">The target file format.</param>
        /// <param name="height">The height.</param>
        /// <param name="width">The width.</param>
        /// <param name="nXSrc">The start X in the sorce.</param>
        /// <param name="nYSrc">The start Y in the source.</param>
        /// <param name="nXDest">The start X in the destination.</param>
        /// <param name="nYDest">The start Y in the destination.</param>
        public static void CaptureWindowToFile(IntPtr handle, string filename, ImageFormat format, int height, int width, int nXSrc = 0, int nYSrc = 0, int nXDest = 0, int nYDest = 0)
        {
            using (Image img = CaptureWindow(handle, height, width, nXSrc, nXSrc, nXDest, nYDest))
            {
                img.Save(filename, format);
            }
        }

        /// <summary>
        /// Captures a screen shot of the entire desktop, and saves it to a file
        /// </summary>
        /// <param name="filename">the file name to save</param>
        public static void CaptureScreenToFile(string filename)
        {
            CaptureScreenToFile(filename, ImageFormat.Bmp);
        }
        /// <summary>
        /// Captures a screen shot of the entire desktop, and saves it to a file
        /// </summary>
        /// <param name="filename">the file name to save</param>
        /// <param name="format">The target file format.</param>
        public static void CaptureScreenToFile(string filename, ImageFormat format)
        {
            using (Image img = CaptureScreen())
            {
                img.Save(filename, format);
            }
        }

        /// <summary>
        /// Helper class containing Gdi32 API functions
        /// </summary>
        private static class GDI32
        {
            public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter
            [DllImport("gdi32.dll")]
            public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest,
                int nWidth, int nHeight, IntPtr hObjectSource,
                int nXSrc, int nYSrc, int dwRop);
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth,
                int nHeight);
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
            [DllImport("gdi32.dll")]
            public static extern bool DeleteDC(IntPtr hDC);
            [DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);
            [DllImport("gdi32.dll")]
            public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
        }

        /// <summary>
        /// Helper class containing User32 API functions
        /// </summary>
        public static class User32
        {
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsWindow(IntPtr hWnd);

            [StructLayout(LayoutKind.Sequential)]
            public struct RECT
            {
                public int left;
                public int top;
                public int right;
                public int bottom;
            }
            [DllImport("user32.dll")]
            public static extern IntPtr GetDesktopWindow();
            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowDC(IntPtr hWnd);
            [DllImport("user32.dll")]
            public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
            [DllImport("user32.dll")]
            public static extern IntPtr GetClientRect(IntPtr hWnd, ref RECT rect);
            [DllImport("user32.dll")]
            public static extern bool ClientToScreen(IntPtr hWnd, ref System.Drawing.Point lpPoint);

        }
    }
}