using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Timers;

namespace tud.mci.tangram.TangramLector
{
    /// <summary>
    /// Event arguments for a captured screen shot
    /// </summary>
    /// <seealso cref="System.EventArgs" />
    public class CaptureChangedEventArgs : EventArgs
    {
        /// <summary>
        /// The captured image
        /// </summary>
        /// <value>The screen capture.</value>
        public Image Img { get; private set; }
        /// <summary>
        /// Initializes a new instance of the <see cref="CaptureChangedEventArgs"/> class.
        /// </summary>
        /// <param name="image">The image.</param>
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
        //double refreshRate;
        Timer refreshTimer;
        /// <summary>
        /// Gets the Window handle of the window to capture.
        /// </summary>
        /// <value>
        /// The window handle.
        /// </value>
        private IntPtr _whnd; // { get; private set; }
        /// <summary>
        /// Gets the Window handle of the window to capture.
        /// </summary>
        /// <value>
        /// The window handle.
        /// </value>
        public IntPtr Whnd
        {
            get { return _whnd; }
            private set
            {
                _whnd = value;
            }
        }
        /// <summary>
        /// Gets or sets the screen position to capture.
        /// </summary>
        /// <value>
        /// The screen position.
        /// </value>
        public Object ScreenPos { private set; get; }
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
            //refreshRate = interval;

            refreshTimer = new Timer(interval);
            refreshTimer.Elapsed += new ElapsedEventHandler(refreshTimer_Elapsed);
        }

        int tickMod = 1;
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver" /> class.
        /// This Observer generates screen shots of the main screen desktop window.
        /// </summary>
        /// <param name="timer">The timer to listen to.</param>
        /// <param name="tickMod">The tick modulo. Modulo value for which the timer event should be 
        /// handled if the result is 0. Default is 1.</param>
        /// <exception cref="System.ArgumentNullException">Timer has to be not NULL</exception>
        public ScreenObserver(Timer timer, int tickMod = 1)
        {
            if (timer == null) throw new ArgumentNullException("Timer has to be not NULL", "timer");

            this.tickMod = tickMod;
            refreshTimer = timer;
            refreshTimer.Elapsed += new ElapsedEventHandler(refreshTimer_Elapsed);
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver" /> class.
        /// This Observer generates screen shots of the given application/window.
        /// </summary>
        /// <param name="timer">The timer to listen to.</param>
        /// <param name="wHnd">The window handle of the window/application to observe.</param>
        /// <param name="tickMod">The tick modulo. Modulo value for which the timer event should be
        /// handled if the result is 0. Default is 1.</param>
        public ScreenObserver(Timer timer, IntPtr wHnd, int tickMod = 1) : this(timer, tickMod) { _whnd = wHnd; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver" /> class.
        /// This Observer generates screen shots of the screen position.
        /// </summary>
        /// <param name="timer">The timer to listen to.</param>
        /// <param name="screenPosition">The position on the screen.</param>
        /// <param name="tickMod">The tick modulo. Modulo value for which the timer event should be
        /// handled if the result is 0. Default is 1.</param>
        public ScreenObserver(Timer timer, Rectangle screenPosition, int tickMod = 1) : this(timer, tickMod) { ScreenPos = screenPosition; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver"/> class.
        /// This Observer generates screen shots of the given application/window.
        /// </summary>
        /// <param name="interval">The capturing interval in milliseconds.</param>
        /// <param name="wHnd">The window handle of the window/application to observe.</param>
        public ScreenObserver(double interval, IntPtr wHnd) : this(interval) { _whnd = wHnd; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ScreenObserver"/> class.
        /// This Observer generates screen shots of the screen position.
        /// </summary>
        /// <param name="interval">The capturing interval in milliseconds.</param>
        /// <param name="screenPosition">The position on the screen.</param>
        public ScreenObserver(double interval, Rectangle screenPosition) : this(interval) { ScreenPos = screenPosition; }

        #endregion

        #region Event
        /// <summary>
        /// Handler for changing screen shots.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="CaptureChangedEventArgs"/> instance containing the event data.</param>
        public delegate void CaptureChangedEventHandler(object sender, CaptureChangedEventArgs e);
        /// <summary>
        /// Occurs when a new screen shot was produces.
        /// </summary>
        public event CaptureChangedEventHandler Changed;
        #endregion

        int _runs = -1;
        //int captureCount = 0;
        #region Timer
        void refreshTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            _runs++;

            if (Changed != null && (_runs % (tickMod) == 0))
            {
                if (_runs >= tickMod)
                {
                    _runs = 0;
                }
                using (Image sc = capture())
                {
                    if (sc != null)
                    {
                        try
                        {
                            Changed.Invoke(this, new CaptureChangedEventArgs(sc));
                        }
                        catch { }
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
            if (refreshTimer != null)
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
            else
            {
                Whnd = IntPtr.Zero;
            }
        }
        /// <summary>
        /// Sets the screen position to capture.
        /// </summary>
        /// <param name="screenPosition">The screen position.</param>
        public void SetScreenPos(Rectangle screenPosition) { SetWhnd(IntPtr.Zero); ScreenPos = screenPosition; }
        /// <summary>
        /// Sets the part of window handle to capture.
        /// </summary>
        /// <param name="position">The screen position to capture.</param>
        /// <param name="wHnd">The corresponding window handle to caputure.</param>
        public void SetPartOfWhnd(Rectangle position, IntPtr wHnd)
        {
            if (tud.mci.tangram.TangramLector.ScreenCapture.User32.IsWindow(wHnd))
            { ScreenPos = position; Whnd = wHnd; }
            else
            {
                Whnd = IntPtr.Zero;
                ScreenPos = null;
            }
        }
        /// <summary>
        /// Captures the first screen (positive parts).
        /// </summary>
        public void ObserveScreen()
        {
            Stop();
            Whnd = tud.mci.tangram.TangramLector.ScreenCapture.User32.GetDesktopWindow();
            ScreenPos = null;
            Start();
        }

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

            // get the device context (DC) for the entire target window
            IntPtr hdcSrc = User32.GetWindowDC(handle);
            // create a device context we can copy to
            IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
            // create a bitmap we can copy it to,
            // using GetDeviceCaps to get the width/height
            IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc, width, height);
            // select the bitmap object
            IntPtr hOld = GDI32.SelectObject(hdcDest, hBitmap);
            // copy the bit blocks of the window bitmap to the save bitmap 
            GDI32.BitBlt(hdcDest, nXDest, nYDest, width, height, hdcSrc, nXSrc, nYSrc, GDI32.SRCCOPY);
            
            //// restore selection
            //GDI32.SelectObject(hdcDest, hOld);
            
            
            // clean up
            User32.ReleaseDC(handle, hdcSrc);
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
                //_capCount++;
                // free up the Bitmap object
                GDI32.DeleteObject(hBitmap);
                // GDI32.DeleteObject(hdcDest);
                GDI32.DeleteObject(hOld);
                // GDI32.DeleteObject(hdcSrc);
                //forceGCFree();
            }
            return img;
        }

        //static volatile int _capCount = 0;
        //private static void forceGCFree()
        //{
        //    if (_capCount > 10)
        //    {
        //        var rs = GC.WaitForFullGCComplete(10);
        //        if (rs.HasFlag(GCNotificationStatus.Succeeded))
        //        {
        //            _capCount = 0;
        //        }
        //    }
        //}

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

            /// <summary>
            /// The BitBlt function performs a bit-block transfer of the color data corresponding 
            /// to a rectangle of pixels from the specified source device context into a destination 
            /// device context.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd183370(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hObject">A handle to the destination device context.</param>
            /// <param name="nXDest">The x-coordinate, in logical units, of the upper-left corner 
            /// of the destination rectangle.</param>
            /// <param name="nYDest">The y-coordinate, in logical units, of the upper-left corner 
            /// of the destination rectangle.</param>
            /// <param name="nWidth">The width, in logical units, of the source and destination 
            /// rectangles.</param>
            /// <param name="nHeight">The height, in logical units, of the source and the 
            /// destination rectangles.</param>
            /// <param name="hObjectSource">A handle to the source device context.</param>
            /// <param name="nXSrc">The x-coordinate, in logical units, of the upper-left corner 
            /// of the source rectangle.</param>
            /// <param name="nYSrc">The y-coordinate, in logical units, of the upper-left corner 
            /// of the source rectangle.</param>
            /// <param name="dwRop">A raster-operation code. These codes define how the color data 
            /// for the source rectangle is to be combined with the color data for the destination 
            /// rectangle to achieve the final color.</param>
            /// <returns>If the function succeeds, the return value is nonzero. 
            /// If the function fails, the return value is zero.</returns>
            [DllImport("gdi32.dll")]
            public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest,
                int nWidth, int nHeight, IntPtr hObjectSource,
                int nXSrc, int nYSrc, int dwRop);
            
            /// <summary>
            /// The CreateCompatibleBitmap function creates a bitmap compatible with the device that is associated with the specified device context.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd183488(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hDC">A handle to a device context.</param>
            /// <param name="nWidth">The bitmap width, in pixels.</param>
            /// <param name="nHeight">The bitmap height, in pixels.</param>
            /// <returns>If the function succeeds, the return value is a handle to the compatible bitmap (DDB).
            /// If the function fails, the return value is NULL.</returns>
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth,
                int nHeight);

            /// <summary>
            /// The CreateCompatibleDC function creates a memory device context (DC) compatible with the specified device.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd183489(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hDC">A handle to an existing DC. If this handle is NULL, the function creates a memory DC compatible with the application's current screen.</param>
            /// <returns>If the function succeeds, the return value is the handle to a memory DC.
            /// If the function fails, the return value is NULL.</returns>
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleDC(IntPtr hDC);

            /// <summary>
            /// The DeleteDC function deletes the specified device context (DC).
            /// An application must not delete a DC whose handle was obtained by calling the GetDC function. Instead, it must call the ReleaseDC function to free the DC.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd183533(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hDC">A handle to the device context.</param>
            /// <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.</returns>
            [DllImport("gdi32.dll")]
            public static extern bool DeleteDC(IntPtr hDC);

            /// <summary>
            /// The DeleteObject function deletes a logical pen, brush, font, bitmap, region, or palette, freeing all system resources associated with the object. After the object is deleted, the specified handle is no longer valid.
            /// Do not delete a drawing object (pen or brush) while it is still selected into a DC.
            /// When a pattern brush is deleted, the bitmap associated with the brush is not deleted. The bitmap must be deleted independently.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd183539(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hObject">A handle to a logical pen, brush, font, bitmap, region, or palette.</param>
            /// <returns>If the function succeeds, the return value is nonzero.
            /// If the specified handle is not valid or is currently selected into a DC, the return value is zero.</returns>
            [DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);

            /// <summary>
            /// The SelectObject function selects an object into the specified device context (DC). 
            /// The new object replaces the previous object of the same type.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd162957(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hDC">A handle to the DC.</param>
            /// <param name="hObject">A handle to the object to be selected..</param>
            /// <returns>If an error occurs and the selected object is not a region, the return value is NULL. Otherwise, it is HGDI_ERROR.</returns>
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

            /// <summary>
            /// The GetWindowDC function retrieves the device context (DC) for the entire window, 
            /// including title bar, menus, and scroll bars. A window device context permits 
            /// painting anywhere in a window, because the origin of the device context is the 
            /// upper-left corner of the window instead of the client area.
            /// GetWindowDC assigns default attributes to the window device context each time 
            /// it retrieves the device context. Previous attributes are lost.
            /// After painting is complete, the ReleaseDC function must be called to release 
            /// the device context. Not releasing the window device context has serious effects 
            /// on painting requested by applications.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd144947(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hWnd">A handle to the window with a device context that is to be 
            /// retrieved. If this value is NULL, GetWindowDC retrieves the device context for 
            /// the entire screen.
            /// If this parameter is NULL, GetWindowDC retrieves the device context for the 
            /// primary display monitor. </param>
            /// <returns>If the function succeeds, the return value is a handle to a device context
            /// for the specified window. If the function fails, the return value is NULL, 
            /// indicating an error or an invalid hWnd parameter.</returns>
            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowDC(IntPtr hWnd);

            /// <summary>
            /// The ReleaseDC function releases a device context (DC), freeing it for use by other applications. The effect of the ReleaseDC function depends on the type of DC. It frees only common and window DCs. It has no effect on class or private DCs.
            /// <seealso cref="https://msdn.microsoft.com/de-de/library/windows/desktop/dd162920(v=vs.85).aspx"/>
            /// </summary>
            /// <param name="hWnd">A handle to the window whose DC is to be released.</param>
            /// <param name="hDC">A handle to the DC to be released.</param>
            /// <returns>The return value indicates whether the DC was released. If the DC was released, the return value is 1. If the DC was not released, the return value is zero.</returns>
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