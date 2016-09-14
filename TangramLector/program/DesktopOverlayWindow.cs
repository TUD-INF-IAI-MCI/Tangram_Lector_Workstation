using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace tud.mci.tangram.TangramLector
{
    public partial class DesktopOverlayWindow : Form
    {
        private static readonly DesktopOverlayWindow _instance = new DesktopOverlayWindow();
        public static DesktopOverlayWindow Instance { get { return _instance; } }

        // import user32 functions for extended window styling
        // Retrieve window attribue
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        // Changes an attribute of the specified window (hWnd is the WindowHandle, index GWL_EXSTYLE (-20) is for extended window style.
        // see https://msdn.microsoft.com/en-us/library/windows/desktop/ff700543(v=vs.85).aspx
        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        static extern bool SetWindowPos(
             int hWnd,             // Window handle
             int hWndInsertAfter,  // Placement-order handle
             int X,                // Horizontal position
             int Y,                // Vertical position
             int cx,               // Width
             int cy,               // Height
             uint uFlags);         // Window positioning flags

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static void ShowInactiveTopmost(Form frm, int posX=Int32.MinValue, int posY = Int32.MinValue)
        {
            ShowWindow(frm.Handle, 4 /*SW_SHOWNOACTIVATE*/);
            if (posX == Int32.MinValue) posX = frm.Left;
            if (posY == Int32.MinValue) posY = frm.Top;
            SetWindowPos(frm.Handle.ToInt32(), -1/*HWND_TOPMOST*/,frm.Left, frm.Top, frm.Width, frm.Height,0x0010/*SWP_NOACTIVATE*/);
        }

        private DesktopOverlayWindow()
        {
            InitializeComponent();
            // make window click transparent
            // GetWindowLongPtr(Handle, -20/*GWL_EXSTYLE*/): for retrieving already set GWL_EXSTYLE flags
            // WS_EX_LAYERED: system automatically composes and repaints layered windows and the windows of underlying applications. As a result, layered windows are rendered smoothly, without the flickering typical of complex window regions. In addition, layered windows can be partially translucent, that is, alpha-blended.
            // WS_EX_TRANSPARENT: make window (click) transparent
            // WS_EX_TOOLWINDOW: hide from alt+tab
            SetWindowLong(this.Handle, -20/*GWL_EXSTYLE*/, 
                Convert.ToInt32(GetWindowLong(Handle, -20/*GWL_EXSTYLE*/)) |
                    0x00080000/*WS_EX_LAYERED*/ | 0x00000020/*WS_EX_TRANSPARENT*/ | 0x00000080/*WS_EX_TOOLWINDOW*/ | 0x00000008/*WS_EX_TOPMOST*/ | 0x08000000/*WS_EX_NOACTIVATE*/ | 0x40000000/*WS_CHILD*/); 
        }

        // override ShowWithoutActivation to not steal focus from other windows when showing
        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        // delegate type declaration
        private delegate void ShowDesktopOverlayWindowDelegate(ref byte[] pngImageByteArray, Rectangle boundingBox, double opacity, int blinkTimeInMs, int[] blinkPattern);
        // internal method to be called as delegate to be invoked on the forms thread
        private void _blinkDesktopOverlayWindow(ref byte[] pngImageByteArray, Rectangle boundingBox, double opacity, int blinkTimeInMs, int[] blinkPattern)
        {
            onOffStepTimer.Stop();
            totalDisplayTimer.Stop();
            this.Location = new Point(boundingBox.X, boundingBox.Y);
            this.Size = new Size(boundingBox.Width, boundingBox.Height);
            this.Opacity = opacity;
            loadBitmap(ref pngImageByteArray);

            ShowInactiveTopmost(this);
            //this.Visible=true;

            totalDisplayTimer.Interval = blinkTimeInMs;
            totalDisplayTimer.Start();
            onOffStepIntervalls.Clear();
            if (blinkPattern.Length < 2)
            {
                onOffStepIntervalls.Add(500);
                onOffStepIntervalls.Add(500);
            }
            else foreach (int d in blinkPattern)
                {
                    onOffStepIntervalls.Add(d);
                }
            onOffStepIndex = 0;
            onOffStepTimer.Interval = onOffStepIntervalls[onOffStepIndex];
            onOffStepTimer.Start();
            switchBlinking();
        }

        /// <summary>
        /// Starts to flash the window on the screen with the given parameters. 
        /// This method creates a callback and invokes it on the forms thread automatically.
        /// </summary>
        /// <param name="pngImageByteArray">The image to flash.</param>
        /// <param name="boundingBox">The bounding box (absolute desktop pixel positions)</param>
        /// <param name="opacity">The opacity (0.0 = transparent .. 1.0 = opaque)</param>
        /// <param name="blinkTimeInMs">The total blinking time in ms until blinking stops and the window is hidden again.</param>
        /// <param name="blinkPattern">The pattern of milliseconds to be on|off|..., 
        /// e.g. [250,150,250,150,250,150,800,1000] for three short flashes (on for 250ms) 
        /// with gaps (of 150ms) and one long on time (800ms) followd by a pause(1s).
        /// The pattern is repeated until the total blinking time is over
        /// If less than 2 values are given, the default of [500, 500] is used!
        /// </param>
        public void initBlinking(ref byte[] pngImageByteArray, Rectangle boundingBox, double opacity, int blinkTimeInMs, int[] blinkPattern)
        {
            ShowDesktopOverlayWindowDelegate ShowDesktopOverlayWindowCallback = _blinkDesktopOverlayWindow;
            this.BeginInvoke(ShowDesktopOverlayWindowCallback, new object[5] { pngImageByteArray, boundingBox, opacity, blinkTimeInMs, blinkPattern });
        }

        private delegate void RefreshBoundsDelegate(Rectangle boundingBox, ref byte[] pngImageByteArray);

        [DllImport("user32.dll", EntryPoint = "MoveWindow")]
        static extern bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);

        private void _refreshBounds(Rectangle boundingBox, ref byte[] pngImageByteArray)
        {
            if (totalDisplayTimer.Enabled)
            { 
                MoveWindow(this.Handle, boundingBox.X, boundingBox.Y, boundingBox.Width, boundingBox.Height, true);
                if (pngImageByteArray.Length>0)
                {
                    loadBitmap(ref pngImageByteArray);
                }
            }
        }

        /// <summary>
        /// If the window is still blinking its bounding box may be refreshed. Also if pngInageByteArray is not empty, a new image may be load.
        /// This may be used to follow a shape moving around, scaled or resized or changed of zoom or scrolling in openoffice.
        /// </summary>
        /// <param name="boundingBox">Absolute screen bounding box.</param>
        /// <param name="pngImageByteArray"></param>
        public void refreshBounds(Rectangle boundingBox, ref byte[] pngImageByteArray)
        {
            RefreshBoundsDelegate RefreshBoundsCallback = _refreshBounds;
            this.BeginInvoke(RefreshBoundsCallback, new object[2] { boundingBox, pngImageByteArray });
        }

        private void loadBitmap(ref byte[] byteArray)
        {
            Bitmap foregroundBitmap = new Bitmap(Size.Width, Size.Height);
            Bitmap backgroundBitmap = new Bitmap(Size.Width, Size.Height);
            MemoryStream stream = new MemoryStream(byteArray);
            try
            {
                Image png = Image.FromStream(stream);
                //get a graphics object from the new image
                Graphics g = Graphics.FromImage(foregroundBitmap);
                Graphics gBg = Graphics.FromImage(backgroundBitmap);
                // create the negative color matrix
                ColorMatrix colorMatrixRed = new ColorMatrix(new float[][]
                {
                    // more red
                    new float[] {2f, 0, 0, 0, 0},
                    new float[] {0, -1f, 0, 0, 0},
                    new float[] {0, 0, -1f, 0, 0},
                    new float[] {0, 0, 0, 1f, 0},
                    new float[] {.2f, 0f, 0f, 0, 1f}
                });
                ColorMatrix colorMatrixInverse = new ColorMatrix(new float[][]
                {
                    // invert
                    new float[] {-1f, 0, 0, 0, 0},
                    new float[] {0, -1f, 0, 0, 0},
                    new float[] {0, 0, -1f, 0, 0},
                    new float[] {0, 0, 0, 1f, 0},
                    new float[] {1f, 1f, 1f, 0, 1f}
                });
                // create some image attributes
                ImageAttributes attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrixRed);
                ImageAttributes attributesBg = new ImageAttributes();
                attributesBg.SetColorMatrix(colorMatrixInverse);
                // draw with color matrix
                g.DrawImage(png, new Rectangle(0, 0, Size.Width, Size.Height), 0, 0, Size.Width, Size.Height, GraphicsUnit.Pixel, attributes);
                gBg.DrawImage(png, new Rectangle(0, 0, Size.Width, Size.Height), 0, 0, Size.Width, Size.Height, GraphicsUnit.Pixel, attributesBg);

                //dispose the Graphics object
                g.Dispose();
                gBg.Dispose();
                this.BackgroundImage = backgroundBitmap;
                transparentOverlayPictureBox1.Image = foregroundBitmap;
            }
            catch (ArgumentException ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, ex.Source, ex);
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            totalDisplayTimer.Stop();
            this.Visible = false;
        }


        private List<int> onOffStepIntervalls = new List<int>();
        private int onOffStepIndex=0;
        private bool onState = false;

        private void onOffStepTimer_Tick(object sender, EventArgs e)
        {
            switchBlinking();
            // resume with next timer interval
            onOffStepIndex++;
            // resume with first index if overflow
            if (onOffStepIndex >= onOffStepIntervalls.Count) onOffStepIndex=0;
            // adopt timing
            onOffStepTimer.Interval = onOffStepIntervalls[onOffStepIndex];
        }

        private void switchBlinking()
        {
            // switch blinking state
            onState = !onState;
            try
            {
                if (onState)
                {
                    transparentOverlayPictureBox1.Visible = true;
                    //ShowInactiveTopmost(this);
                }
                else
                {
                    transparentOverlayPictureBox1.Visible = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, ex); 
            }
        }
    }
}
