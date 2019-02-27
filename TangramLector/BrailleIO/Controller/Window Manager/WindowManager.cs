using BrailleIO;
using BrailleIO.Interface;
using BrailleIO.Renderer;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;
using tud.mci.tangram.TangramLector.BrailleIO.Model;
using tud.mci.tangram.TangramLector.BrailleIO.View;
using tud.mci.tangram.TangramLector.Window_Manager;

namespace tud.mci.tangram.TangramLector
{
    /// <summary>
    /// Main class for handling views, screens and foci's.
    /// </summary>
    /// <seealso cref="tud.mci.tangram.TangramLector.SpecializedFunctionProxies.AbstractSpecializedFunctionProxyBase" />
    /// <seealso cref="tud.mci.LanguageLocalization.ILocalizable" />
    /// <seealso cref="System.IDisposable" />
    public partial class WindowManager : ILocalizable, IDisposable
    {
        #region Members

        #region static
        /// <summary>
        /// Global timer triggering the blinkTimer_Tick event for blinking pins.
        /// </summary>
        public readonly static BlinkTimer blinkTimer = BlinkTimer.Instance;
        static readonly LL LL = new LL(Properties.Resources.Language);
        #endregion

        #region constants
        /// <summary>
        /// The standard brightness threshold for image renderer
        /// </summary>
        const int STANDARD_CONTRAST_THRESHOLD = 210;
        /// <summary>
        /// The step size for changing the brightness threshold of image renderer
        /// </summary>
        private const int THRESHOLD_STEP = 10;
        /// <summary>
        /// The name of the main screen containing the default drawing application views. 
        /// </summary>
        public const String BS_MAIN_NAME = "Mainscreen";
        /// <summary>
        /// The name of the screen containing the full screen view to the drawing application. 
        /// </summary>
        public const String BS_FULLSCREEN_NAME = "Fullscreen";
        /// <summary>
        /// The name of the main screen containing the minimap view to the drawing application. 
        /// </summary>
        public const String BS_MINIMAP_NAME = "Minimap";
        /// <summary>
        /// Name of the second large content region.
        /// </summary>
        public const String VR_CENTER_2_NAME = "ContentRegion2";
        /// <summary>
        /// Name of the primary large content region
        /// </summary>
        public const String VR_CENTER_NAME = "ContentRegion";
        /// <summary>
        /// Name of the top title region
        /// </summary>
        public const String VR_TOP_NAME = "Titleregion";
        /// <summary>
        /// Name of the top right status indexing region
        /// </summary>
        public const String VR_STATUS_NAME = "Statusregion";
        /// <summary>
        /// Name of the detail information region at the bottom of the screen.
        /// </summary>
        public const String VR_DETAIL_NAME = "Detailregion";
        // TODO: maybe necessary to have a title for each screen (e.g. when multiple windows were filtered)
        /// <summary>
        /// The title of the main screen - title region content
        /// </summary>
        public const String MAINSCREEN_TITLE = "Tangram Lector"; // TODO: Name der Datei mit übergeben
        /// <summary>
        /// The factor between the 10 dpi pin device resolution and the 96 dpi resolution of the DRAW application.
        /// </summary>
        private const float _PRINT_ZOOM_FACTOR = 0.10561666418313964f;
        #endregion

        #region private
        InteractionManager InteractionManager = InteractionManager.Instance;
        AudioRenderer audioRenderer = AudioRenderer.Instance;
        BrailleIOMediator io = BrailleIOMediator.Instance;
        Size deviceSize;
        /// <summary>
        /// Screen that is shown in minimap mode.
        /// </summary>
        //BrailleIOScreen screenForMinimap;
        /// <summary>
        /// Current view mode of center region (has to be a value from LectorViews).
        /// </summary>
        LectorView currentView = LectorView.Drawing;
        /// <summary>
        /// Screen that was visible before minimap was activated.
        /// </summary>
        BrailleIOScreen screenBeforeMinimap = null;
        #endregion

        #region public
        /// <summary>
        /// Timer-based value telling if blinking pins should be up or down.
        /// </summary>
        public bool BlinkPinsUp = true;

        /// <summary>
        /// Gets the DRAW application model.
        /// </summary>
        /// <value>
        /// The DRAW application model.
        /// </value>
        public OoDrawModel DrawAppModel { get; private set; }
        DrawRenderer drawRenderer_Fullscreen = new DrawRenderer();
        DrawRenderer drawRenderer_Minimap = new DrawRenderer();
        DrawRenderer drawRenderer_Default = new DrawRenderer();
        MinimapRendererHook minimapHook = new MinimapRendererHook();
        GridRendererHook gridHook = new GridRendererHook();



        private FollowFocusModes _focusMode = FollowFocusModes.NONE;
        /// <summary>
        /// Gets or sets the focus mode of the application.
        /// </summary>
        /// <value>
        /// The focus mode.
        /// </value>
        public FollowFocusModes FocusMode
        {
            get { return _focusMode; }
            set
            {
                var old = _focusMode;
                _focusMode = value;
                updateStatusRegionContent();
                sayInteractionModeChange(old, _focusMode);
                fireFocusModeChangeEvent(old, _focusMode);
            }
        }

        private ConcurrentBag<string> detailRegionHistory = new ConcurrentBag<string>();
        /// <summary>
        /// Get a list of all consistent messages shown in detail region
        /// </summary>
        /// <returns>list of all consistent messages</returns>
        public string[] GetDetailRegionHistory()
        {
            return detailRegionHistory.ToArray<String>();
        }
        #endregion

        #region Singleton
        private static WindowManager instance = new WindowManager();
        /// <summary>
        /// Gets the singleton instance of the window manager.
        /// </summary>
        /// <value>The instance.</value>
        public static WindowManager Instance
        {
            get
            {
                if (instance == null) instance = new WindowManager();
                return instance;
            }
            set { instance = null; }
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="WindowManager"/> class from being created.
        /// </summary>
        private WindowManager()
        {
            DrawAppModel = new OoDrawModel();
            registerForEvents();
            initBlinkTimer();
            BuildScreens();
        }
        #endregion

        #endregion

        #region screen handling

        internal void BuildScreens()
        {
            if (io != null && io.AdapterManager != null && io.AdapterManager.ActiveAdapter != null && io.AdapterManager.ActiveAdapter.Device != null)
            {
                deviceSize = new Size(io.AdapterManager.ActiveAdapter.Device.DeviceSizeX, io.AdapterManager.ActiveAdapter.Device.DeviceSizeY);
            }
            else
            {
                //showOffAdapter size 
                deviceSize = new Size(120, 60);
            }
            BrailleIOScreen mainScreen = buildMainScreen(deviceSize.Width, deviceSize.Height, currentView);
            buildFullScreen(deviceSize.Width, deviceSize.Height);
            buildMinimapScreen(deviceSize.Width, deviceSize.Height);
        }

        internal void CleanScreen()
        {
            try
            {
                if (io != null)
                {
                    io.RemoveView(BS_MAIN_NAME);
                    io.RemoveView(BS_FULLSCREEN_NAME);
                    io.RemoveView(BS_MINIMAP_NAME);
                }
            }
            catch { }
        }

        internal void UpdateScreens()
        {
            try
            {

                if (io != null && io.AdapterManager != null && io.AdapterManager.ActiveAdapter != null && io.AdapterManager.ActiveAdapter.Device != null)
                {
                    deviceSize = new Size(io.AdapterManager.ActiveAdapter.Device.DeviceSizeX, io.AdapterManager.ActiveAdapter.Device.DeviceSizeY);
                }
                else
                {
                    //showOffAdapter size 
                    deviceSize = new Size(120, 60);
                }

                updateMainScreen(deviceSize.Width, deviceSize.Height);
                updateFullScreen(deviceSize.Width, deviceSize.Height);
                updateMinimapScreen(deviceSize.Width, deviceSize.Height);
            }
            catch { }
        }

        #endregion

        #region init methods

        /// <summary>
        /// Init the blinking timer object.
        /// </summary>
        void initBlinkTimer()
        {
            blinkTimer.ThreeQuarterTick += new EventHandler<EventArgs>(blinkTimer_Tick);
        }

        /// <summary>
        /// Generates the main screen consisting of title (top region), detail region and center region.
        /// </summary>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <param name="view">The view of the main region.</param>
        /// <returns></returns>
        BrailleIOScreen buildMainScreen(int width, int height, LectorView view)
        {
            BrailleIOScreen mainScreen = new BrailleIOScreen(BS_MAIN_NAME);
            mainScreen.SetHeight(height);
            mainScreen.SetWidth(width);

            var center = getMainScreenCenterRegion(0, 0, width, height);
            center.ShowScrollbars = true;
            var center2 = getMainScreenCenter2Region(0, 0, width, height);
            center2.ShowScrollbars = true;
            center2.SetVisibility(false);
            center2.SetZoom(1);

            var top = getMainTopRegion(0, 0, width, 7);
            var status = getMainStatusRegion(width - 12, 0, 12, 5);
            status.SetText(LectorStateNames.STANDARD_MODE);

            var detail = getMainDetailRegion(0, height - 7, width, 7);
            detail.ShowScrollbars = true; // make the region scrollable
            // make the BrailleRenderer to ignore the last line space
            detail.SetText(LL.GetTrans("tangram.lector.wm.no_element_selected")); // set text to enable the BrailleRenderer
            var renderer = detail.ContentRender;
            if (renderer != null && renderer is MatrixBrailleRenderer)
            {
                ((MatrixBrailleRenderer)renderer).RenderingProperties |= RenderingProperties.IGNORE_LAST_LINESPACE;
            }

            mainScreen.AddViewRange(VR_CENTER_2_NAME, center2);
            mainScreen.AddViewRange(VR_CENTER_NAME, center);
            mainScreen.AddViewRange(VR_TOP_NAME, top);
            mainScreen.AddViewRange(VR_STATUS_NAME, status);
            mainScreen.AddViewRange(VR_DETAIL_NAME, detail);

            setRegionContent(mainScreen, VR_TOP_NAME, MAINSCREEN_TITLE);
            setRegionContent(mainScreen, VR_DETAIL_NAME, LL.GetTrans("tangram.lector.wm.no_element_selected"));

            currentView = view;
            fillMainCenterContent(mainScreen);

            if (io != null)
            {
                io.AddView(BS_MAIN_NAME, mainScreen);
                io.ShowView(BS_MAIN_NAME);
                io.RefreshDisplay(true);
            }
            return mainScreen;
        }

        bool updateMainScreen(int width, int height)
        {
            if (io != null)
            {
                // main screen
                BrailleIOScreen mainScreen = io.GetView(BS_MAIN_NAME) as BrailleIOScreen;
                if (mainScreen != null)
                {
                    mainScreen.SetHeight(height);
                    mainScreen.SetWidth(width);

                    // center
                    BrailleIOViewRange center = mainScreen.GetViewRange(VR_CENTER_NAME);
                    if (center != null)
                    {
                        center.SetHeight(height);
                        center.SetWidth(width);
                    }

                    // center 2
                    BrailleIOViewRange center2 = mainScreen.GetViewRange(VR_CENTER_2_NAME);
                    if (center2 != null)
                    {
                        center2.SetHeight(height);
                        center2.SetWidth(width);
                    }

                    //top
                    BrailleIOViewRange top = mainScreen.GetViewRange(VR_TOP_NAME);
                    if (top != null)
                    {
                        top.SetWidth((int)(width * 1.4)); // only that region content is not doing line breaks
                    }

                    // status
                    BrailleIOViewRange status = mainScreen.GetViewRange(VR_STATUS_NAME);
                    if (status != null)
                    {
                        status.SetLeft(width - 12);
                    }

                    // detail
                    BrailleIOViewRange detail = mainScreen.GetViewRange(VR_DETAIL_NAME);
                    if (detail != null)
                    {
                        detail.SetTop(height - 7);
                        detail.SetWidth(width);
                    }

                    io.RefreshDisplay(true);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Shows the center region in full screen (other regions are set invisible).
        /// </summary>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns></returns>
        BrailleIOScreen buildFullScreen(int width, int height)
        {
            BrailleIOScreen fullScreen = new BrailleIOScreen(BS_FULLSCREEN_NAME);
            BrailleIOViewRange center = new BrailleIOViewRange(0, 0, width, height);
            center.SetZoom(-1);
            center.SetBorder(0);
            center.SetContrastThreshold(STANDARD_CONTRAST_THRESHOLD);
            center.ShowScrollbars = true;
            center.SetOtherContent(DrawAppModel, drawRenderer_Fullscreen);

            // register grid hook
            if (drawRenderer_Fullscreen != null && gridHook != null)
            {
                drawRenderer_Fullscreen.UnregisterHook(gridHook);
                drawRenderer_Fullscreen.RegisterHook(gridHook);
            }

            fullScreen.AddViewRange(VR_CENTER_NAME, center);
            if (io != null) io.AddView(BS_FULLSCREEN_NAME, fullScreen);
            fullScreen.SetVisibility(false);

            return fullScreen;
        }

        bool updateFullScreen(int width, int height)
        {
            if (io != null)
            {
                // main screen
                BrailleIOScreen fullScreen = io.GetView(BS_FULLSCREEN_NAME) as BrailleIOScreen;
                if (fullScreen != null)
                {
                    fullScreen.SetHeight(height);
                    fullScreen.SetWidth(width);

                    // center
                    BrailleIOViewRange center = fullScreen.GetViewRange(VR_CENTER_NAME);
                    if (center != null)
                    {
                        center.SetHeight(height);
                        center.SetWidth(width);
                    }

                    io.RefreshDisplay(true);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Shows the center region in full screen (other regions are set invisible).
        /// </summary>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns></returns>
        BrailleIOScreen buildMinimapScreen(int width, int height)
        {
            BrailleIOScreen minimapScreen = new BrailleIOScreen(BS_MINIMAP_NAME);

            var center = getMainScreenCenterRegion(0, 0, width, height);
            var top = getMainTopRegion(0, 0, width, 7);
            var detail = getMainDetailRegion(0, height - 7, width, 7);

            minimapScreen.AddViewRange(VR_CENTER_NAME, center); // TODO auch center2 nötig?
            minimapScreen.AddViewRange(VR_TOP_NAME, top);
            minimapScreen.AddViewRange(VR_DETAIL_NAME, detail);


            // register minimap view frame renderer hook
            if (drawRenderer_Minimap != null)
            {
                drawRenderer_Minimap.UnregisterHook(minimapHook);
                drawRenderer_Minimap.RegisterHook(minimapHook);
            }

            center.SetOtherContent(DrawAppModel, drawRenderer_Minimap);
            setRegionContent(minimapScreen, VR_TOP_NAME, "Minimap");
            setRegionContent(minimapScreen, VR_DETAIL_NAME, LL.GetTrans("tangram.lector.wm.minimap.vr_detail"));

            if (io != null) io.AddView(BS_MINIMAP_NAME, minimapScreen);
            minimapScreen.SetVisibility(false);

            return minimapScreen;
        }

        bool updateMinimapScreen(int width, int height)
        {
            if (io != null)
            {
                // main screen
                BrailleIOScreen minimapScreen = io.GetView(BS_MINIMAP_NAME) as BrailleIOScreen;
                if (minimapScreen != null)
                {
                    minimapScreen.SetHeight(height);
                    minimapScreen.SetWidth(width);

                    // center
                    BrailleIOViewRange center = minimapScreen.GetViewRange(VR_CENTER_NAME);
                    if (center != null)
                    {
                        center.SetHeight(height);
                        center.SetWidth(width);
                    }

                    //top
                    BrailleIOViewRange top = minimapScreen.GetViewRange(VR_TOP_NAME);
                    if (top != null)
                    {
                        top.SetWidth((int)(width * 1.4)); // only that region content is not doing line breaks
                    }

                    // status
                    BrailleIOViewRange status = minimapScreen.GetViewRange(VR_STATUS_NAME);
                    if (status != null)
                    {
                        status.SetLeft(width - 12);
                    }

                    // center 2
                    BrailleIOViewRange detail = minimapScreen.GetViewRange(VR_DETAIL_NAME);
                    if (detail != null)
                    {
                        detail.SetTop(height - 7);
                        detail.SetWidth(width);
                    }

                    io.RefreshDisplay(true);
                    return true;
                }
            }
            return false;
        }

        BrailleIOViewRange getMainScreenCenterRegion(int left, int top, int width, int height)
        {
            BrailleIOViewRange center = new BrailleIOViewRange(left, top, width, height);
            center.SetZoom(-1);
            center.SetBorder(0);
            center.SetContrastThreshold(STANDARD_CONTRAST_THRESHOLD);
            center.SetMargin(7, 0);
            return center;
        }

        BrailleIOViewRange getMainScreenCenter2Region(int left, int top, int width, int height)
        {
            BrailleIOViewRange center2 = new BrailleIOViewRange(left, top, width, height);
            center2.SetZoom(-1);
            center2.SetBorder(0);
            center2.SetContrastThreshold(STANDARD_CONTRAST_THRESHOLD);
            center2.SetMargin(7, 0);
            return center2;
        }

        BrailleIOViewRange getMainTopRegion(int left, int top, int width, int height)
        {
            BrailleIOViewRange _top = new BrailleIOViewRange(left, top, width, height);
            _top.SetBorder(0, 0, 1);
            _top.SetMargin(0, 0, 1);
            _top.SetPadding(0, 11, 1, 0);
            _top.SetWidth((int)(width * 1.4)); // only that region content is not doing line breaks
            return _top;
        }

        BrailleIOViewRange getMainDetailRegion(int left, int top, int width, int height)
        {
            BrailleIOViewRange bottom = new BrailleIOViewRange(left, top, width, height);
            bottom.SetBorder(1, 0, 0);
            bottom.SetMargin(1, 0, 0);
            bottom.SetPadding(1, 0, 0);
            return bottom;
        }

        BrailleIOViewRange getMainStatusRegion(int left, int top, int width, int height)
        {
            BrailleIOViewRange _status = new BrailleIOViewRange(left, top, width, height);
            _status.SetBorder(0, 0, 0, 1);
            _status.SetMargin(0, 0, 0, 2);
            _status.SetPadding(0, 0, 0, 1);
            return _status;
        }

        #endregion

        #region setter and getter for region and views

        /// <summary>
        /// Sets the content of the region.
        /// </summary>
        /// <param name="screen">The screen.</param>
        /// <param name="vrName">Name of the region (viewRange).</param>
        /// <param name="content">The content.</param>
        /// <returns>true if the content could be set, otherwise false</returns>
        private bool setRegionContent(BrailleIOScreen screen, String vrName, Object content)
        {
            if (screen != null)
            {
                BrailleIOViewRange vr = screen.GetViewRange(vrName);
                if (vr != null)
                {
                    if (content != null)
                    {
                        if (content is bool[,]) { vr.SetMatrix(content as bool[,]); }
                        else if (content is Bitmap) { vr.SetBitmap(content as Bitmap); }
                        else if (content is String) { vr.SetText(content.ToString()); }
                        else { return false; }
                        return true;
                    }
                    else { vr.SetMatrix(null); } //TODO: check if this works
                }
            }
            return false;
        }

        /// <summary>
        /// Sets the content of the region.
        /// </summary>
        /// <param name="screen">The screen.</param>
        /// <param name="vrName">Name of the region (viewRange).</param>
        /// <param name="content">The content.</param>
        /// <param name="renderer">The renderer to use to transform the content into a matrix.</param>
        /// <returns>true if the content could be set, otherwise false</returns>
        public bool SetRegionContent(BrailleIOScreen screen, String vrName, Object content, IBrailleIOContentRenderer renderer)
        {
            if (screen != null)
            {
                BrailleIOViewRange vr = screen.GetViewRange(vrName);
                if (vr != null)
                {
                    if (content != null)
                    {
                        vr.SetOtherContent(content, renderer);
                    }
                    else { vr.SetMatrix(null); }
                }
            }
            return false;
        }

        /// <summary>
        /// Set consistent content of the current detail region and saves it in history.
        /// For system messages, that should only be visible temporarely, use ShowTemporaryMessageInDetailRegion() function.
        /// </summary>
        /// <param name="content">content to render</param>
        public void SetDetailRegionContent(String content)
        {
            detailRegionHistory.Add(content);
            tempMessageShown = false;
            setDetailRegionContent(content);
        }

        /// <summary>
        /// set content of current detail region (without adding to history)
        /// </summary>
        /// <param name="content">content to render</param>
        private void setDetailRegionContent(String content)
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (vs != null)
            {
                setRegionContent(vs, VR_DETAIL_NAME, content);
                BrailleIOViewRange vr = vs.GetViewRange(VR_DETAIL_NAME);
                if (vr != null) scrollViewRangeTo(vr, 0);
            }
        }

        /// <summary>
        /// Set the content of the current top region.
        /// </summary>
        /// <param name="content">content to render</param>
        public void SetTopRegionContent(String content)
        {
            BrailleIOScreen vs = GetVisibleScreen();
            setRegionContent(vs, VR_TOP_NAME, content);
        }

        /// <summary>
        /// Sets the content of the status region.
        /// </summary>
        /// <param name="status">Should be a string of the LectorStateNames class.</param>
        public void SetStatusRegionContent(String status)
        {
            BrailleIOScreen vs = io.GetView(BS_MAIN_NAME) as BrailleIOScreen;
            if (vs != null)
            {
                int width = vs.GetViewRange(VR_STATUS_NAME).GetWidth();
                int maxCharacters = width / 3 - 1; // one Braille letter needs 3 pins
                status = status.Substring(0, maxCharacters);
                setRegionContent(vs, VR_STATUS_NAME, status);
            }
        }

        /// <summary>
        /// Updates the content of status region.
        /// </summary>
        private void updateStatusRegionContent()
        {
            BrailleIOScreen vs = io.GetView(BS_MAIN_NAME) as BrailleIOScreen;
            if (vs != null)
            {
                if (InteractionManager != null)
                {
                    if (InteractionManager.Mode.Equals(InteractionMode.Braille))
                    {
                        setRegionContent(vs, VR_STATUS_NAME, LectorStateNames.BRAILLE_MODE);
                        return;
                    }
                }
                FollowFocusModes mode = FocusMode;
                switch (mode)
                {
                    case FollowFocusModes.FOLLOW_MOUSE_FOCUS:
                        setRegionContent(vs, VR_STATUS_NAME, LectorStateNames.FOLLOW_MOUSE_FOCUS_MODE);
                        break;
                    case FollowFocusModes.FOLLOW_BRAILLE_FOCUS:
                        setRegionContent(vs, VR_STATUS_NAME, LectorStateNames.FOLLOW_BRAILLE_FOCUS_MODE);
                        break;
                    default:
                        setRegionContent(vs, VR_STATUS_NAME, LectorStateNames.STANDARD_MODE);
                        break;
                }
            }
        }

        /// <summary>
        /// Fills the center region of the main screen with the content depending on the currentView.
        /// </summary>
        private void fillMainCenterContent()
        {
            if (io != null)
            {
                BrailleIOScreen screen = GetVisibleScreen();
                if (screen != null)
                {
                    fillMainCenterContent(screen);
                }
            }
        }

        /// <summary>
        /// Fills the center region of the main screen with the content depending on the currentView.
        /// </summary>
        private void fillMainCenterContent(BrailleIOScreen screen)
        {
            if (screen != null)
            {
                BrailleIOViewRange center = screen.GetViewRange(VR_CENTER_NAME);
                BrailleIOViewRange center2 = screen.GetViewRange(VR_CENTER_2_NAME);
                if (center != null)
                {
                    switch (currentView)
                    {
                        case LectorView.Drawing:
                            center.SetOtherContent(DrawAppModel, drawRenderer_Default);

                            // register grid hook
                            if (drawRenderer_Default != null && gridHook != null)
                            {
                                drawRenderer_Default.UnregisterHook(gridHook);
                                drawRenderer_Default.RegisterHook(gridHook);
                            }

                            setCaptureArea();
                            if (center2 != null)
                                center2.SetVisibility(false);


                            break;
                        case LectorView.Braille:
                            String content = "Hallo Welt"; // TODO: set real content
                            setRegionContent(screen, VR_CENTER_2_NAME, content);
                            break;
                        default:
                            center.SetOtherContent(DrawAppModel, drawRenderer_Default);

                            // register grid hook
                            if (drawRenderer_Default != null && gridHook != null)
                            {
                                drawRenderer_Default.UnregisterHook(gridHook);
                                drawRenderer_Default.RegisterHook(gridHook);
                            }

                            setCaptureArea();
                            break;
                    }
                }
            }
        }

        ///// <summary>
        ///// Set the content in minimap mode. Thereby the screen capturing bitmap is shown with blinking frame.
        ///// </summary>
        ///// <param name="vr">ViewRange in which minimap content should be shown.</param>
        ///// <param name="e">EventArgs of the screen capturing event.</param>
        //void setMinimapContent(BrailleIOViewRange vr, CaptureChangedEventArgs e)
        //{
        //    int width = vr.ContentBox.Width;
        //    int height = vr.ContentBox.Height;
        //    Bitmap bmp = new Bitmap(width, height);
        //    Graphics graph = Graphics.FromImage(bmp);

        //    if (screenBeforeMinimap != null)
        //    {
        //        BrailleIOViewRange imgvr = screenBeforeMinimap.GetViewRange(VR_CENTER_NAME);
        //        if (imgvr == null) return;

        //        int xoffset = imgvr.GetXOffset();
        //        int yoffset = imgvr.GetYOffset();
        //        int imgWidth = imgvr.ContentWidth;
        //        int imgHeight = imgvr.ContentHeight;

        //        if (imgWidth > 0 && imgHeight > 0)
        //        {
        //            // calculate image size for minimap (complete image has to fit into view range)
        //            int imgScaledWidth = imgWidth;
        //            int imgScaledHeight = imgHeight;
        //            double scaleFactorX = 1;
        //            double scaleFactorY = 1;

        //            if (width != 0 && imgScaledWidth > width)
        //            {
        //                scaleFactorX = (double)imgWidth / (double)width;
        //                if (scaleFactorX != 0)
        //                {
        //                    imgScaledWidth = width;
        //                    imgScaledHeight = (int)(imgHeight / scaleFactorX);
        //                }
        //            }
        //            if (height != 0 && imgScaledHeight > height)
        //            {
        //                scaleFactorY = (double)imgScaledHeight / (double)height;
        //                if (scaleFactorY != 0)
        //                {
        //                    imgScaledHeight = height;
        //                    imgScaledWidth = (int)(imgScaledWidth / scaleFactorY);
        //                }
        //            }

        //            // calculate scaling factor from original image to minimap image size
        //            MinimapScalingFactor = 1 / (scaleFactorX * scaleFactorY);
        //            double zoom = imgvr.GetZoom();
        //            if (zoom > 0) MinimapScalingFactor = MinimapScalingFactor * zoom;

        //            // calculate position and size of the blinking frame
        //            int x = Math.Abs(xoffset) * imgScaledWidth / imgWidth;
        //            int y = Math.Abs(yoffset) * imgScaledHeight / imgHeight;
        //            int frameWidth = width * imgScaledWidth / imgWidth;
        //            int frameHeigth = height * imgScaledHeight / imgHeight;

        //            // draw scaled image and blinking frame
        //            graph.DrawImage(e.Img, 0, 0, imgScaledWidth, imgScaledHeight);
        //            Color frameColor = Color.Black;
        //            if (!blinkPinsUp) frameColor = Color.White;
        //            graph.DrawRectangle(new Pen(frameColor, 2), x, y, frameWidth, frameHeigth);
        //            vr.SetBitmap(bmp);
        //        }
        //    }
        //}

        /// <summary>
        /// Get the visible screen.
        /// </summary>
        /// <returns>currently visible screen</returns>
        public BrailleIOScreen GetVisibleScreen()
        {
            var views = io.GetViews();
            if (views != null && views.Count > 0)
            {
                try
                {
                    var vs = views.First(x => (x is IViewable && ((IViewable)x).IsVisible()));
                    if (vs != null && vs is BrailleIOScreen)
                    {
                        return vs as BrailleIOScreen;
                    }
                }
                catch (InvalidOperationException) { } //Happens if no view could been found in the listing
            }
            return null;
        }

        /// <summary>
        /// Get the current lector view mode.
        /// </summary>
        /// <returns>current lector view</returns>
        public LectorView GetCurrentLectorView()
        {
            return currentView;
        }

        /// <summary>
        /// Get the text content of the region.
        /// </summary>
        /// <param name="screen">The screen.</param>
        /// <param name="vrName">Name of the region (viewRange).</param>
        /// <returns>Text content or String.Empty if there is no valid data (e.g. content is not text).</returns>
        public String GetTextContentOfRegion(BrailleIOScreen screen, String vrName)
        {
            String content = String.Empty;
            if (screen != null)
            {
                BrailleIOViewRange vr = screen.GetViewRange(vrName);
                if (vr != null)
                {
                    content = vr.GetText();
                }
            }
            return content;
        }

        /// <summary>
        /// Get the content offset position of the region.
        /// </summary>
        /// <param name="screen">The screen.</param>
        /// <param name="vrName">Name of the region (viewRange).</param>
        /// <returns>Offset position or Point.Empty if there is no valid data.</returns>
        public Point GetContentOffsetOfRegion(BrailleIOScreen screen, String vrName)
        {
            Point offset = Point.Empty;
            if (screen != null)
            {
                BrailleIOViewRange vr = screen.GetViewRange(vrName);
                if (vr != null)
                {
                    offset = vr.OffsetPosition;
                }
            }
            return offset;
        }

        #endregion

        #region screen capture methods

        /// <summary>
        /// Get the view range where screen capturing should be shown (center view range of the visible screen).
        /// </summary>
        /// <returns>view range</returns>
        BrailleIOViewRange getTargetViewRangeForScreenCapturing()
        {
            BrailleIOScreen s = GetVisibleScreen();

            if (s != null)
            {
                return s.GetViewRange(VR_CENTER_NAME);
            }
            return io.GetView(VR_CENTER_NAME) as BrailleIOViewRange;
        }

        /// <summary>
        /// Starts the screen capturing and register for the event.
        /// </summary>
        void setCaptureArea()
        {
            if (DrawAppModel.ScreenObserver == null)
            {
                DrawAppModel.ScreenObserver = new ScreenObserver(100);

                // so_Changed event handles the rendering of the bitmap
                try { DrawAppModel.ScreenObserver.Changed -= new ScreenObserver.CaptureChangedEventHandler(so_Changed); }
                catch (Exception) { }
                DrawAppModel.ScreenObserver.Changed += new ScreenObserver.CaptureChangedEventHandler(so_Changed);
            }
            DrawAppModel.ScreenObserver.Start();
        }

        /// <summary>
        /// Handles the Changed event of the screen observer control and send it to view range that is defined by getTargetViewRangeForScreenCapturing().
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="tud.mci.tangram.TangramLector.CaptureChangedEventArgs"/> instance containing the event data.</param>
        void so_Changed(object sender, CaptureChangedEventArgs e)
        {
            Task t = new Task(() =>
            {
                // add some delay so the other listeners can do their job (e.g. prerender)
                Thread.Sleep(5);
                io.RefreshDisplay(true);
            });
            t.Start();
        }

        #endregion

        #region helping methods

        /// <summary>
        /// Synchronize the zoom and x/y offset of the given view range with that of the center view range of the visible screen.
        /// </summary>
        /// <param name="vr">view range that should be syncronized with visible center</param>
        private void syncViewRangeWithVisibleCenterViewRange(BrailleIOViewRange vr)
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (!vs.Name.Equals(BS_MINIMAP_NAME))
            {
                BrailleIOViewRange visibleCenterVR = vs.GetViewRange(VR_CENTER_NAME);
                double zoom = visibleCenterVR.GetZoom();
                int x = visibleCenterVR.GetXOffset();
                int y = visibleCenterVR.GetYOffset();

                if (visibleCenterVR.ContentHeight <= vr.ContentBox.Height && visibleCenterVR.ContentWidth <= vr.ContentBox.Width)
                {
                    vr.SetZoom(-1); // für Anpassung bei kleinster Zoomstufe
                }
                else vr.SetZoom(zoom);
                vr.SetXOffset(x);
                vr.SetYOffset(y);
            }
        }

        /// <summary>
        /// Synchronize the contrastThreshold and InvertImage value.
        /// </summary>
        /// <param name="oldvr">ViewRange whose values should be used in the new ViewRange.</param>
        /// <param name="newvr">ViewRange that should get the values of the old ViewRange.</param>
        private void syncContrastSettings(BrailleIOViewRange oldvr, BrailleIOViewRange newvr)
        {
            if (oldvr != null && newvr != null)
            {
                newvr.InvertImage = oldvr.InvertImage;
                newvr.SetContrastThreshold(oldvr.GetContrastThreshold());
            }
        }

        /// <summary>
        /// Get the touched view range.
        /// </summary>
        /// <param name="x">x value of the tap (on pin device)</param>
        /// <param name="y">y value of the tap (on pin device)</param>
        /// <returns>touched view range</returns>
        public BrailleIOViewRange GetTouchedViewRange(double x, double y)
        {
            return io != null ? io.GetViewAtPosition((int)x, (int)y) : null;
        }
        /// <summary>
        /// Get the touched view range.
        /// </summary>
        /// <param name="x">x value of the tap (on pin device)</param>
        /// <param name="y">y value of the tap (on pin device)</param>
        /// <param name="s">visible screen</param>
        /// <returns>touched view range</returns>
        public BrailleIOViewRange GetTouchedViewRange(double x, double y, BrailleIOScreen s)
        {
            if (s != null)
            {
                OrderedDictionary viewRanges = s.GetViewRanges();
                if (viewRanges.Count > 0)
                {
                    object[] keys = new object[viewRanges.Keys.Count];
                    viewRanges.Keys.CopyTo(keys, 0);
                    for (int i = keys.Length - 1; i >= 0; i--)
                    {
                        BrailleIOViewRange vr = viewRanges[keys[i]] as BrailleIOViewRange;
                        if (vr != null && vr.IsVisible())
                        {
                            if (x >= vr.GetLeft() && x <= (vr.GetLeft() + vr.GetWidth()))
                            {
                                if (y >= vr.GetTop() && y <= (vr.GetTop() + vr.GetHeight()))
                                {
                                    return vr;
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Get position of tap within the content
        /// </summary>
        /// <param name="x">x value of tap (on pin device)</param>
        /// <param name="y">y value of tap (on pin device)</param>
        /// <param name="vr">touched view range</param>
        /// <returns>position within content</returns>
        public Point GetTapPositionInContent(double x, double y, BrailleIOViewRange vr)
        {
            Point p = new Point();
            if (vr != null)
            {
                double contentX = x - vr.OffsetPosition.X - vr.Padding.Left - vr.Margin.Left - vr.Border.Left;
                double contentY = y - vr.OffsetPosition.Y - vr.Padding.Top - vr.Margin.Top - vr.Border.Top;
                p = new Point((int)Math.Round(contentX), (int)Math.Round(contentY));
                System.Diagnostics.Debug.WriteLine("tap content position: " + p.ToString());
            }
            return p;
        }

        /// <summary>
        /// Get position of tap on the screen
        /// </summary>
        /// <param name="x">x value of tap (on pin device)</param>
        /// <param name="y">y value of tap (on pin device)</param>
        /// <param name="vr">touched view range</param>
        /// <returns>position on the screen</returns>
        public Point GetTapPositionOnScreen(double x, double y, BrailleIOViewRange vr)
        {
            Point p = GetTapPositionInContent(x, y, vr);
            if (vr != null && p != null)
            {
                double zoom = vr.GetZoom();
                if (zoom != 0)
                {
                    int x_old = p.X;
                    int y_old = p.Y;
                    p = new Point((int)Math.Round(x_old / zoom), (int)Math.Round(y_old / zoom));

                    if (DrawAppModel.ScreenObserver != null && DrawAppModel.ScreenObserver.ScreenPos is Rectangle)
                    {
                        Rectangle sp = (Rectangle)DrawAppModel.ScreenObserver.ScreenPos;
                        p.X += sp.X;
                        p.Y += sp.Y;
                    }
                    System.Diagnostics.Debug.WriteLine("tap screen position: " + p.ToString());
                }
            }
            return p;
        }

        #endregion

        #region temporary messages

        private const int tempMessageTime = 20; // a value of 20 is around 10 seconds
        private List<string> detailRegionTempHistory = new List<string>();
        /// <summary>
        /// Get history of detail region temporary messages.
        /// </summary>
        /// <returns>list with all temp messages shown in detail region</returns>
        public List<string> GetDetailRegionTempMessageHistory() { return detailRegionTempHistory; }
        private volatile int messageTimerCount = 0;
        private volatile bool tempMessageShown = false;

        /// <summary>
        /// Show a message in detail region that will disappear after a given time (see tempMessageTime).
        /// </summary>
        /// <param name="message">message that should be shown in detail region</param>
        public void ShowTemporaryMessageInDetailRegion(string message)
        {
            messageTimerCount = tempMessageTime;
            tempMessageShown = true;
            detailRegionTempHistory.Add(message);
            setDetailRegionContent(message);
        }

        /// <summary>
        /// Get the last temporary message shown in detail region.
        /// </summary>
        /// <returns>last temporary message</returns>
        public string GetLastTemporaryMessage()
        {
            if (detailRegionTempHistory.Count > 1)
            {
                System.Diagnostics.Debug.WriteLine("[DETAIL REGION TEMP HISTORY] last entry: " + detailRegionTempHistory.ElementAt(1));
                return detailRegionTempHistory.ElementAt(1);
            }
            else return String.Empty;
        }

        /// <summary>
        /// Get the current temporary message shown in detail region.
        /// </summary>
        /// <returns>current temporary message</returns>
        public string GetCurrentTemporaryMessage()
        {
            if (detailRegionTempHistory.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine("[DETAIL REGION TEMP HISTORY] current entry: " + detailRegionTempHistory.ElementAt(0));
                return detailRegionTempHistory.ElementAt(0);
            }
            else return String.Empty;
        }

        /// <summary>
        /// Get the last consistent detail region content.
        /// </summary>
        /// <returns>last consistent detail region content</returns>
        public string GetLastDetailRegionContent()
        {
            if (detailRegionHistory.Count > 1)
            {
                System.Diagnostics.Debug.WriteLine("[DETAIL REGION HISTORY] last entry: " + detailRegionHistory.ElementAt(1));
                return detailRegionHistory.ElementAt(1);
            }
            else return String.Empty;
        }

        /// <summary>
        /// Get the current consistent detail region content.
        /// </summary>
        /// <returns>current consistent detail region content</returns>
        public string GetCurrentDetailRegionContent()
        {
            if (detailRegionHistory.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine("[DETAIL REGION HISTORY] current entry: " + detailRegionHistory.ElementAt(0));
                return detailRegionHistory.ElementAt(0);
            }
            else return String.Empty;
        }

        #endregion

        #region events

        /// <summary>
        /// Event of the blinking timer.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void blinkTimer_Tick(object sender, EventArgs e)
        {
            BlinkPinsUp = !BlinkPinsUp;

            // Timer handling for temporary detail area messages
            if (tempMessageShown)
            {
                if (messageTimerCount > 0) messageTimerCount--;
                else
                {
                    tempMessageShown = false;
                    setDetailRegionContent(GetCurrentDetailRegionContent());
                }
            }

            Task t = new Task(() => { Thread.Sleep(5); io.RefreshDisplay(true); });
            t.Start();

        }

        /// <summary>
        /// Occurs when follow focus mode is changed.
        /// </summary>
        public event EventHandler<FocusModeChangedEventArgs> FollowFocusModeChange;

        private void fireFocusModeChangeEvent(FollowFocusModes oldValue, FollowFocusModes newValue)
        {
            if (FollowFocusModeChange != null)
            {
                try
                {
                    FollowFocusModeChange.Invoke(this, new FocusModeChangedEventArgs(oldValue, newValue));
                }
                catch (Exception) { }
            }
        }

        #endregion

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

        #region IDisposable

        public bool Disposed { get; private set; }

        public void Dispose()
        {
            Disposed = true;
            Instance = null;
            fireDisposingEvent();
        }

        /// <summary>
        /// Occurs when this instance of the WindowManager is disposing.
        /// </summary>
        public event EventHandler Disposing;

        void fireDisposingEvent()
        {
            if (Disposing != null)
            {
                try
                {
                    Disposing.DynamicInvoke(this, null);
                }
                catch { }
            }
        }

        #endregion

        #region Enums
        /// <summary>
        /// View modes for the lector application.
        /// Braille mode is for reading and writing Braille.
        /// Drawing mode is for pixel based output.
        /// </summary>
        public enum LectorView
        {
            Braille = 1,
            Drawing = 2
        }
        #endregion
    }

    /// <summary>
    /// Defines strings for showing in state area.
    /// </summary>
    public static class LectorStateNames
    {
        /// <summary>
        /// Used when system is in normal interaction mode (default).
        /// </summary>
        public const string STANDARD_MODE = "std";
        /// <summary>
        /// Used when inputs are interpreted as Braille keyboard.
        /// </summary>
        public const string BRAILLE_MODE = "brm";
        /// <summary>
        /// Used when focus following mode is set to mouse following.
        /// </summary>
        public const string FOLLOW_MOUSE_FOCUS_MODE = "mff";
        /// <summary>
        /// Used when focus following mode is set to Braille focus following.
        /// </summary>
        public const string FOLLOW_BRAILLE_FOCUS_MODE = "bff";
        /// <summary>
        /// Used while gestures or taps create drawing objects within the canvas.
        /// </summary>
        public const string DIRECT_DRAWING_MODE = "drw";
    }

    /// <summary>
    /// Enum of all focus following modes.
    /// </summary>
    public enum FollowFocusModes
    {
        /// <summary>
        /// Used when no focus following is activated.
        /// </summary>
        NONE,
        /// <summary>
        /// Used when mouse focus following is activated.
        /// </summary>
        FOLLOW_MOUSE_FOCUS,
        /// <summary>
        /// Used when Braille focus following is activated.
        /// </summary>
        FOLLOW_BRAILLE_FOCUS
    }

    /// <summary>
    /// Event args for focus mode change.
    /// </summary>
    public class FocusModeChangedEventArgs : EventArgs
    {
        /// <summary>
        /// follow focus mode before change
        /// </summary>
        public readonly FollowFocusModes OldValue;
        /// <summary>
        /// current follow focus mode
        /// </summary>
        public readonly FollowFocusModes NewValue;

        public FocusModeChangedEventArgs(FollowFocusModes oldValue, FollowFocusModes newValue)
        {
            OldValue = oldValue;
            NewValue = newValue;
        }
    }
}