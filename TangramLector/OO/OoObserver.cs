using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using BrailleIO;
using BrailleIO.Interface;
using TangramLector.OO;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.controller;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.TangramLector.OO
{
    public partial class OoObserver : AbstractSpecializedFunctionProxyBase, IOoDrawConnection, ILocalizable
    {
        #region Members

        AudioRenderer audioRenderer = AudioRenderer.Instance;
        WindowManager windowManager = WindowManager.Instance;
        static LL LL = new LL(Properties.Resources.Language);
        #region Hooks

        internal readonly DrawDocumentRendererHook DocumentBorderHook = new DrawDocumentRendererHook();
        internal readonly FocusRendererHook BrailleDomFocusRenderer = new FocusRendererHook();
        internal readonly FocusRendererHook DrawSelectFocusRenderer = new FocusRendererHook(true);
        /// <summary>
        /// A renderer hook that overlays elements with text with a Braille label.
        /// </summary>
        public readonly BrailleTextView TextRendererHook = new BrailleTextView();

        readonly ImageData imgDataFunctionProxy = ImageData.Instance;

        #endregion

        readonly OpenOfficeDrawShapeManipulator shapeManipulatorFunctionProxy = null;

        /// <summary>
        /// Bounding box of all selected items.
        /// </summary>
        /// <returns></returns>
        public Rectangle SelectedBoundingBox { get; private set; }

        /// <summary>
        /// Get first item of list which is selected on GUI (Mouse focus).
        /// </summary>
        /// <returns></returns>
        public OoShapeObserver SelectedItem { get; private set; }

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="OoObserver"/> class.
        /// Listen to an OoDrawAccessibilityObserver and handles the events.
        /// Basic interpreter for events from an OpenOffice Draw application. 
        /// </summary>
        public OoObserver()
            : base(20)
        {
            shapeManipulatorFunctionProxy = new OpenOfficeDrawShapeManipulator(this, WindowManager.Instance as IZoomProvider, WindowManager.Instance as IFeedbackReceiver);

            WindowManager.blinkTimer.ThreeQuarterTick += new EventHandler<EventArgs>(blinkTimer_Tick);
            OnShapeBoundRectChange = new OoShapeObserver.BoundRectChangeEventHandler(ShapeBoundRectChangeHandler);
            OnViewOrZoomChange = new OoDrawPagesObserver.ViewOrZoomChangeEventHandler(ShapeBoundRectChangeHandler);

            initFunctionProxy();
            registerAsSpecializedFunctionProxy();
            OoSelectionObserver.Instance.SelectionChanged += new EventHandler<OoSelectionChandedEventArgs>(OoDrawAccessibilityObserver_SelectionChanged);

            OoDrawAccessibilityObserver.Instance.DrawWindowOpend += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowOpend);
            OoDrawAccessibilityObserver.Instance.DrawWindowPropertyChange += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowPropertyChange);
            OoDrawAccessibilityObserver.Instance.DrawWindowMinimized += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowMinimized);
            OoDrawAccessibilityObserver.Instance.DrawWindowClosed += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowClosed);
            OoDrawAccessibilityObserver.Instance.DrawWindowActivated += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowActivated);
            OoDrawAccessibilityObserver.Instance.DrawSelectionChanged += new EventHandler<OoAccessibilitySelectionEventArgs>(Instance_DrawSelectionChanged);

            shapeManipulatorFunctionProxy.SelectedShapeChanged += new EventHandler<EventArgs>(shapeManipulatorFunctionProxy_SelectedShapeChanged);
            shapeManipulatorFunctionProxy.PolygonPointSelected += shapeManipulatorFunctionProxy_PolygonPointSelected;

            // make the observes aware of the currently open DRAW docs (still opened before started this instance)
            List<OoAccessibleDocWnd> docList = OoDrawAccessibilityObserver.Instance.GetDrawDocs();
            foreach (var doc in docList)
            {
                OoDrawAccessibilityObserver_DrawWindowActivated(null, new OoWindowEventArgs(doc));
            }

            #region Renderer Hook

            BrailleIOViewRange vr = getDisplayViewRange(WindowManager.BS_MAIN_NAME);
            if (vr != null)
            {
                vr.RendererChanged += new EventHandler<EventArgs>(vr_RendererChanged);
                vr_RendererChanged(vr, null);
            }

            BrailleIOViewRange vr_minimap = getDisplayViewRange(WindowManager.BS_MINIMAP_NAME);
            if (vr_minimap != null)
            {
                vr_minimap.RendererChanged += new EventHandler<EventArgs>(vr_RendererChanged);
                vr_RendererChanged(vr_minimap, null);
            }

            BrailleIOViewRange vr_fullscreen = getDisplayViewRange(WindowManager.BS_FULLSCREEN_NAME);
            if (vr_fullscreen != null)
            {
                vr_fullscreen.RendererChanged += new EventHandler<EventArgs>(vr_RendererChanged);
                vr_RendererChanged(vr_fullscreen, null);
            }

            #endregion
        }

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

        private void initFunctionProxy()
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                // register to the activation and deactivation events
                shapeManipulatorFunctionProxy.Activated += new EventHandler(shapeManipulatorFunctionProxy_Activated);
                shapeManipulatorFunctionProxy.Deactivated += new EventHandler(shapeManipulatorFunctionProxy_Deactivated);

                if (ScriptFunctionProxy.Instance != null)
                {
                    ScriptFunctionProxy.Instance.AddProxy(shapeManipulatorFunctionProxy);
                    shapeManipulatorFunctionProxy.Active = true;
                }
            }
        }

        void shapeManipulatorFunctionProxy_Deactivated(object sender, EventArgs e)
        {
            if (sender == this.shapeManipulatorFunctionProxy)
            {
                if (BrailleDomFocusRenderer != null) { BrailleDomFocusRenderer.Active = false; }
                if (DrawSelectFocusRenderer != null) { DrawSelectFocusRenderer.Active = false; }
                if (TextRendererHook != null) { TextRendererHook.Active = false; }
            }
        }

        void shapeManipulatorFunctionProxy_Activated(object sender, EventArgs e)
        {
            if (sender == this.shapeManipulatorFunctionProxy)
            {
                if (BrailleDomFocusRenderer != null) { BrailleDomFocusRenderer.Active = true; }
                if (DrawSelectFocusRenderer != null) { DrawSelectFocusRenderer.Active = true; }
            }
        }

        void shapeManipulatorFunctionProxy_SelectedShapeChanged(object sender, EventArgs e)
        {
            if (sender != null && sender is OpenOfficeDrawShapeManipulator)
            {
                OoShapeObserver _shape = ((OpenOfficeDrawShapeManipulator)sender).LastSelectedShape;
                stopFocusHighlightModes();

                try
                {
                    if (_shape != null)
                    {
                        // memorize original properties to be restored after blinking
                        InitBrailleDomFocusHighlightMode();

                        WindowManager.Instance.SetDetailRegionContent(OoElementSpeaker.GetElementAudioText(_shape));
                        if (WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_BRAILLE_FOCUS) jumpToDomFocus();
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] new DOM focus " + _shape.Name + " " + _shape.Text);
                    }
                }
                catch (System.Exception ex)
                {
                    _shape = null;
                    Logger.Instance.Log(LogPriority.OFTEN, this, "[DRAW INTERACTION] new DOM focus error --> shape is null", ex);
                }
                if (BrailleDomFocusRenderer != null)
                {
                    BrailleDomFocusRenderer.SetCurrentBoundingBoxByShape(_shape);
                }
            }
        }


        void shapeManipulatorFunctionProxy_PolygonPointSelected(object sender, PolygonPointSelectedEventArgs e)
        {

            if (sender != null && sender is OpenOfficeDrawShapeManipulator)
            {
                //OoShapeObserver _shape = ((OpenOfficeDrawShapeManipulator)sender).LastSelectedShape;
                stopFocusHighlightModes();

                // memorize original properties to be restored after blinking
                if (((OpenOfficeDrawShapeManipulator)sender).LastSelectedShape != null)
                    InitBrailleDomFocusHighlightMode();
            }

            if (BrailleDomFocusRenderer != null)
            {
                if (e != null && e.PolygonPoints != null)
                {
                    Point p = e.PolygonPoints.TransformPointCoordinatesIntoScreenCoordinates(e.Point);
                    BrailleDomFocusRenderer.CurrentPoint = p;
                    //System.Diagnostics.Debug.WriteLine(" [P] ----- Polypoint selected Event: " + e.Point.ToString() + "   Iterator: " + e.PolygonPoints.GetIteratorIndex() + " of " + e.PolygonPoints.Count);
                }
                else
                {    
                    BrailleDomFocusRenderer.CurrentPoint = new Point(-1, -1);
                    //System.Diagnostics.Debug.WriteLine(" [P] ----- Polypoint reset Event");
                }
            }


        }

        #region Renderer Hook Registration and Update

        private void initRendererHook()
        {
            WindowManager wm = WindowManager.Instance;
            if (wm != null)
            {
                BrailleIOMediator io = BrailleIOMediator.Instance;
                if (io != null)
                {
                    var fsScreen = io.GetView(WindowManager.BS_FULLSCREEN_NAME);
                    var nsScreen = io.GetView(WindowManager.BS_MAIN_NAME);

                    addRendererHookToScreen(fsScreen as BrailleIOScreen);
                    addRendererHookToScreen(nsScreen as BrailleIOScreen);
                }
            }
        }

        /// <summary>
        /// Adds the renderer hook to a screen that has a viewRangen named WindowManager.VR_CENTER_NAME.
        /// </summary>
        /// <param name="screen">The screen.</param>
        private void addRendererHookToScreen(BrailleIOScreen screen)
        {
            if (screen != null)
            {
                //try to get the main view range
                if (screen.HasViewRange(WindowManager.VR_CENTER_NAME))
                {
                    BrailleIOViewRange vr = screen.GetViewRange(WindowManager.VR_CENTER_NAME);
                    if (vr != null)
                    {
                        if (vr.IsImage() && vr.ContentRender != null && vr.ContentRender is IBrailleIOHookableRenderer)
                        {
                            ((IBrailleIOHookableRenderer)vr.ContentRender).RegisterHook(BrailleDomFocusRenderer);
                            ((IBrailleIOHookableRenderer)vr.ContentRender).RegisterHook(DrawSelectFocusRenderer);
                            ((IBrailleIOHookableRenderer)vr.ContentRender).RegisterHook(TextRendererHook);
                        }
                        else
                        {
                            vr.RendererChanged += new EventHandler<EventArgs>(vr_RendererChanged);
                        }
                    }
                }
            }
        }

        #endregion

        #region Events

        #region Public Events

        /// <summary>
        /// Occurs when a Draw application window was activated.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> WindowActivated;
        /// <summary>
        /// Occurs when a Draw application window was closed.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> WindowClosed;

        void fireWindowActivatedEvent(OoWindowEventArgs e)
        {
            if (WindowActivated != null)
            {
                try { WindowActivated.Invoke(this, e); }
                catch { }
            }
        }

        void fireWindowClosedEvent(OoWindowEventArgs e)
        {
            if (WindowClosed != null)
            {
                try { WindowClosed.Invoke(this, e); }
                catch { }
            }
        }

        #endregion

        #region Oo Events

        /// <summary>
        /// Starts the draw select focus highlighting mode. 
        /// The current (GUI ?) selections will be observed and highlighted 
        /// with a blinking frame on the non-visual output devices.
        /// </summary>
        public void StartDrawSelectFocusHighlightingMode()
        {
            // look for current selection
            // first, if there is none try getting from last selected args
            if (SelectedBoundingBox.IsEmpty && OoDrawAccessibilityObserver.Instance.LastSelection != null) SelectedBoundingBox = OoDrawAccessibilityObserver.Instance.LastSelection.SelectionBounds;
            Rectangle bb = SelectedBoundingBox;
            Rectangle pageBounds = WindowManager.Instance.ScreenObserver.ScreenPos != null ? (Rectangle)WindowManager.Instance.ScreenObserver.ScreenPos : new Rectangle();
            this.DrawSelectFocusRenderer.DoRenderBoundingBox = true;
            this.DrawSelectFocusHighlightMode = true;
            if (!bb.IsEmpty && pageBounds.Width > 0 && pageBounds.Height > 0)
            {
                System.Drawing.Rectangle absBBox = new System.Drawing.Rectangle(bb.X + pageBounds.X, bb.Y + pageBounds.Y, bb.Width, bb.Height);
                this.DrawSelectFocusRenderer.CurrentBoundingBox = absBBox;
            }
            else
            {
                this.DrawSelectFocusRenderer.CurrentBoundingBox = new System.Drawing.Rectangle(-1, -1, 0, 0);
            }
        }

        private List<OoShapeObserver> currentSelection = new List<OoShapeObserver>();
        void Instance_DrawSelectionChanged(object sender, OoAccessibilitySelectionEventArgs e)
        {
            if (e != null && e.SelectedItems != null)
            {
                if (e.SelectedItems.Count > 0 && !e.SelectionBounds.IsEmpty)
                {
                    if (!e.Silent) Logger.Instance.Log(LogPriority.MIDDLE, this, "[GUI INTERACTION] selection changed, count: " + e.SelectedItems.Count.ToString());
                    SelectedBoundingBox = e.SelectionBounds;
                }
                else
                {
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[GUI INTERACTION] deselection");
                    SelectedItem = null;
                    SelectedBoundingBox = new Rectangle(-1, -1, 0, 0);
                    this.DrawSelectFocusRenderer.CurrentBoundingBox = new System.Drawing.Rectangle(-1, -1, 0, 0);
                }

                if (windowManager.FocusMode == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
                {
                    if (e.SelectedItems.Count > 0)
                    {
                        SelectedBoundingBox = e.SelectionBounds;
                        if (e.Source != null && e.Source.DrawPagesObs != null)
                        {
                            SelectedItem = e.SelectedItems[0];
                        }
                        if (shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.Active && this.DrawSelectFocusRenderer != null)
                        {
                            StartDrawSelectFocusHighlightingMode();
                        }
                    }

                    if (e != null && !e.Silent)
                    {
                        if (e.SelectedItems.Count > 0)
                        {
                            windowManager.MoveToObject(e.SelectionBounds); // sync view box on pin device
                        }
                        if ((e.SelectedItems.Count > 0) &&
                            (e.SelectedItems.Count > 1 || currentSelection.Count > 1 ||
                            currentSelection.Count == 0 ||
                            (currentSelection.Count > 0 && !e.SelectedItems[0].Equals(currentSelection[0])))
                            )
                        {
                            CommunicateSelection(e.SelectedItems);
                        }
                    }
                }
            }
            currentSelection = e != null && e.SelectedItems != null ? e.SelectedItems : new List<OoShapeObserver>();
        }

        /// <summary>
        /// Communicates the selection to the user as detail-region-message and audio output.
        /// </summary>
        /// <param name="e">The e.</param>
        /// <param name="immediately">if set to <c>true</c> immediately.</param>
        public void CommunicateSelection(List<OoShapeObserver> e, bool immediately = true)
        {
            if (e.Count > 1)
            {
                String audioText = LL.GetTrans("tangram.lector.oo_observer.selected_element", e.Count);
                if (immediately) audioRenderer.PlaySoundImmediately(audioText);
                else audioRenderer.PlaySound(audioText);
                windowManager.SetDetailRegionContent(audioText);
            }
            else if (e.Count == 1)
            {
                String elementText = OoElementSpeaker.GetElementAudioText(e[0]);
                String audioText = LL.GetTrans("tangram.lector.oo_observer.selected", elementText);
                if (immediately) audioRenderer.PlaySoundImmediately(audioText);
                else audioRenderer.PlaySound(audioText);
                windowManager.SetDetailRegionContent(elementText);
            }
        }

        /// <summary>
        /// Communicates the last selection to the user as detail-region-message and audio output.
        /// </summary>
        /// <param name="immediately">if set to <c>true</c> immediately.</param>
        public void CommunicateLastSelection(bool immediately = true)
        {
            if (OoDrawAccessibilityObserver.Instance != null)
            {
                var args = OoDrawAccessibilityObserver.Instance.LastSelection;
                if (args != null && args.SelectedItems != null)
                {
                    CommunicateSelection(args.SelectedItems, immediately);
                }
            }
        }

        void OoDrawAccessibilityObserver_DrawWindowActivated(object sender, OoWindowEventArgs e)
        {
            tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
            try
            {
                DocumentBorderHook.Active = true;
                addWindowEvent(e);
                Logger.Instance.Log(LogPriority.DEBUG, this, "ooDraw wnd activated " + e.Window.ToString());

                if (e != null && e.Window != null)
                {
                    fireWindowActivatedEvent(e);
                    if (windowManager != null)
                    {
                        windowManager.SetTopRegionContent(e.Window.Title);
                    }
                }

            }
            catch (System.Exception ex)
            {
                wm.ScreenObserver.ObserveScreen();
                Logger.Instance.Log(LogPriority.DEBUG, ex);
            }
        }

        void OoDrawAccessibilityObserver_DrawWindowClosed(object sender, OoWindowEventArgs e)
        {
            try
            {
                tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
                if (e != null && e.Window != null && wm != null && wm.ScreenObserver != null && wm.ScreenObserver.Whnd == e.Window.Whnd)
                {
                    fireWindowClosedEvent(e);
                    if (windowManager != null)
                    {
                        windowManager.SetTopRegionContent(WindowManager.MAINSCREEN_TITLE);
                    }

                    //TODO: check for other windows
                    wm.ScreenObserver.ObserveScreen();
                }
                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] ooDraw wnd closed" + e.Window.ToString());
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, ex);
            }
            audioRenderer.PlaySound(LL.GetTrans("tangram.lector.oo_observer.window_closed"));
        }

        void OoDrawAccessibilityObserver_DrawWindowMinimized(object sender, OoWindowEventArgs e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "ooDraw wnd minimized " + e.Window.ToString());
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, ex);
            }
            //tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
            ////TODO: check why the screenobserver does not capture the screen
            //wm.ScreenObserver.ObserveScreen();

            addWindowEvent(e);

            audioRenderer.PlaySound(LL.GetTrans("tangram.lector.oo_observer.window_minimized", e.Window.Title));
        }

        void OoDrawAccessibilityObserver_DrawWindowPropertyChange(object sender, OoWindowEventArgs e)
        {
            addWindowEvent(e);
        }

        #region Window Property Changeing

        readonly ConcurrentStack<OoWindowEventArgs> stack = new ConcurrentStack<OoWindowEventArgs>();
        Thread windowEventThread;

        private void addWindowEvent(OoWindowEventArgs e)
        {
            if (e != null)
            {
                stack.Push(e);
            }

            if (windowEventThread == null || !windowEventThread.IsAlive)
            {
                windowEventThread = new Thread(new ThreadStart(handleWindowEvents));
                int trys = 0;
                while (trys++ < 10)
                {
                    try
                    {
                        windowEventThread.Start();
                        break;
                    }
                    catch
                    {
                        if (windowEventThread != null) Thread.Sleep(10);
                        else break;
                    }
                }
                if (trys > 9)
                {
                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] cannot start window event handling thread");
                }

            }
            else { }
        }

        private readonly ConcurrentDictionary<IntPtr, Rectangle> boundsCache = new ConcurrentDictionary<IntPtr, Rectangle>();

        void handleWindowEvents()
        {
            while (!stack.IsEmpty)
            {
                OoWindowEventArgs e;
                bool result = stack.TryPop(out e);
                if (!result) continue;
                stack.Clear();
                if (e != null && e.Window != null && !e.Window.Disposed)
                {
                    if (!e.Window.IsVisible() || e.Type.HasFlag(WindowEventType.CLOSED) || e.Type.HasFlag(WindowEventType.MINIMIZED))
                    {
                        resetScreenObserver(windowManager);
                    }
                    else
                    {
                        if (windowManager != null)
                        {
                            windowManager.SetTopRegionContent(e.Window.Title);
                        }

                        Thread.Sleep(50); // gives the element the chance to finish the changes before requesting the new properties. Should prevent a hang on. 

                        tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
                        try
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "ooDraw wnd property changed " + e.Window.ToString());

                            var bounds = e.Window.DocumentComponent.ScreenBounds;
                            var whndl = e.Window.Whnd;

                            if (bounds.Height < 1 || bounds.Width < 1)
                            {
                                //TODO: check if minimized
                                if (boundsCache.ContainsKey(whndl)) { bounds = boundsCache[whndl]; }
                            }
                            else
                            {
                                boundsCache.AddOrUpdate(whndl, bounds, (key, existingVal) => { return bounds; });
                            }

                            DocumentBorderHook.Wnd = e.Window;
                            wm.ScreenObserver.SetPartOfWhnd(bounds, whndl);

                        }
                        catch (System.Exception ex)
                        {
                            resetScreenObserver(wm);
                            Logger.Instance.Log(LogPriority.DEBUG, ex);
                        }
                    }
                }
            }

            windowEventThread = null;
            return;
        }

        private static void resetScreenObserver(tud.mci.tangram.TangramLector.WindowManager wm)
        {

            if (wm != null)
            {
                if (wm.ScreenObserver != null) wm.ScreenObserver.ObserveScreen();
                wm.SetTopRegionContent("");
            }
        }

        #endregion

        void OoDrawAccessibilityObserver_DrawWindowOpend(object sender, OoWindowEventArgs e)
        {
            tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "ooDraw wnd opened " + e.Window.ToString());
                wm.ScreenObserver.SetPartOfWhnd(e.Window.DocumentComponent.ScreenBounds, e.Window.Whnd);
                DocumentBorderHook.Update();
            }
            catch (System.Exception ex)
            {
                wm.ScreenObserver.ObserveScreen();
                Logger.Instance.Log(LogPriority.DEBUG, ex);
            }
        }

        void OoDrawAccessibilityObserver_SelectionChanged(object sender, OoSelectionChandedEventArgs e)
        {
            //System.Diagnostics.Debug.WriteLine("Instance_SelectionChanged");
        }

        #endregion

        #endregion

        #region Functions

        public void Reset()
        {
            OoSelectionObserver.Instance.Reset();
            OoDrawAccessibilityObserver.Instance.Reset();
        }

        /// <summary>
        /// Determines if the Observer knows the given window handle as a observed window.
        /// </summary>
        /// <param name="wHndl">The window handle.</param>
        /// <returns>The observed OoAccessibleDocWnd or <c>null</c> if the window handle is unknown</returns>
        public OoAccessibleDocWnd ObservesWHndl(int wHndl)
        {
            List<OoAccessibleDocWnd> wnds = OoDrawAccessibilityObserver.Instance.GetDrawDocs();
            foreach (var item in wnds)
            {
                try
                {
                    if (item.Whnd.ToInt32() == wHndl) return item;
                }
                catch { }
            }
            return null;
        }

        /// <summary>
        /// Gets the shape for modification.
        /// </summary>
        /// <param name="c">The c.</param>
        /// <param name="observed">The observed.</param>
        public OoShapeObserver GetShapeForModification(OoAccComponent c, OoAccessibleDocWnd observed)
        {
            if (observed != null && c.Role != AccessibleRole.INVALID)
            {
                //TODO: prepare name:

                //var shape = observed.DrawPagesObs.GetRegisteredShapeObserver(c.AccComp); 
                OoShapeObserver shape = observed.GetRegisteredShapeObserver(c);
                if (shape != null)
                {
                    if (shapeManipulatorFunctionProxy != null && !ImageData.Instance.Active)
                    {
                        OoElementSpeaker.PlayElementImmediately(shape, LL.GetTrans("tangram.lector.oo_observer.selected", String.Empty));
                        //audioRenderer.PlaySound("Form kann manipuliert werden");
                        shapeManipulatorFunctionProxy.LastSelectedShape = shape;
                        return shape;
                    }
                    else // title+desc dialog handling
                    {
                        ImageData.Instance.NewSelectionHandling(shape);
                    }
                }
                else
                {
                    // disable the pageShapes
                    if (c.Name.StartsWith("PageShape:"))
                    {
                        return null;
                    }

                    OoElementSpeaker.PlayElementImmediately(c, LL.GetTrans("tangram.lector.oo_observer.selected"));
                    audioRenderer.PlaySound(LL.GetTrans("tangram.lector.oo_observer.selected_element.locked"));
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the last selected shape (DOM / Braille focus).
        /// </summary>
        /// <returns>The OoShapeObserever for the DRAW-object or <c>null</c></returns>
        public OoShapeObserver GetLastSelectedShape()
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                return shapeManipulatorFunctionProxy.LastSelectedShape;
            }
            return null;
        }

        /// <summary>
        /// Gets the current GUI selection.
        /// </summary>
        /// <returns></returns>
        public OoShapeObserver GetCurrentSelection()
        {
            if (OoDrawAccessibilityObserver.Instance != null && OoDrawAccessibilityObserver.Instance.LastSelection != null && OoDrawAccessibilityObserver.Instance.LastSelection.SelectedItems != null && OoDrawAccessibilityObserver.Instance.LastSelection.SelectedItems.Count > 0)
            {
                this.SelectedItem = OoDrawAccessibilityObserver.Instance.LastSelection.SelectedItems[0];
                this.SelectedBoundingBox = OoDrawAccessibilityObserver.Instance.LastSelection.SelectionBounds;
                return this.SelectedItem;
            }
            return null;
        }

        /// <summary>
        /// Try to get the active OpenOffice document window.
        /// Check the ScreenObserver for his window otherwise uses the first
        /// of all observed documents.
        /// </summary>
        /// <returns></returns>
        public OoAccessibleDocWnd GetActiveDocument()
        {
            // check the screen observer
            // return the document that is currently observed and therefor shown on the output device.
            if (WindowManager.Instance != null)
            {
                ScreenObserver so = WindowManager.Instance.ScreenObserver;
                if (so != null)
                {
                    IntPtr whndl = so.Whnd;
                    if (whndl != null && whndl != IntPtr.Zero)
                    {
                        OoObserver obs = OoConnector.Instance.Observer;
                        if (obs != null)
                        {
                            OoAccessibleDocWnd doc = obs.ObservesWHndl(whndl.ToInt32());
                            if (doc != null)
                            {
                                return doc;
                            }
                        }
                    }
                }
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] An instance of the WindowManager singleton could not been achieved!");
            }


            // get all observed documents -- choose the first
            var docs = OoDrawAccessibilityObserver.Instance.GetDrawDocs();
            if (docs.Count > 0)
            {
                if (docs.Count == 1) return docs[0];
                else
                {
                    // can not decide which one is the one to use ...
                    //foreach (var doc in docs)
                    //{
                    //    Debug.GetAllInterfacesOfObject(doc.DocumentWindow);
                    //    var set = tud.mci.tangram.Accessibility.OoAccessibility.GetAccessibleStates(doc.MainWondow.getAccessibleContext().getAccessibleStateSet());
                    //}

                    return docs[0];
                }
            }
            return null;
        }

        /// <summary>
        /// Highlights the current Braille focus on the screen (triggers blinking on the GUI).
        /// </summary>
        internal void HighlightBrailleFocusOnScreen()
        {
            InitBrailleDomFocusHighlightMode();
        }

        #endregion

        #region IOoDrawConnection

        bool IOoDrawConnection.ISConnected()
        {
            return true;
        }

        Object IOoDrawConnection.GetActiveDrawDocument()
        {
            return GetActiveDocument();
        }

        Object IOoDrawConnection.GetCurrentDrawSelection()
        {

            OoShapeObserver shape = SelectedItem;
            if (shape == null)
            {
                shape = GetCurrentSelection();
            }
            return shape;
        }

        #endregion

        #region Renderer Hook

        BrailleIOViewRange getDisplayViewRange(String screenName)
        {
            var ms = BrailleIOMediator.Instance.GetView(screenName);
            if (ms != null && ms is BrailleIOScreen)
            {
                var vr = ((BrailleIOScreen)ms).GetViewRange(WindowManager.VR_CENTER_NAME);
                if (vr != null)
                {
                    return vr;
                }
            }
            return null;
        }

        void vr_RendererChanged(object sender, EventArgs e)
        {
            if (sender != null && sender is BrailleIOViewRange)
            {
                var renderer = ((BrailleIOViewRange)sender).ContentRender;
                if (renderer != null && renderer is IBrailleIOHookableRenderer)
                {
                    ((IBrailleIOHookableRenderer)renderer).UnregisterHook(DocumentBorderHook);
                    ((IBrailleIOHookableRenderer)renderer).RegisterHook(DocumentBorderHook);

                    ((IBrailleIOHookableRenderer)renderer).UnregisterHook(BrailleDomFocusRenderer);
                    ((IBrailleIOHookableRenderer)renderer).RegisterHook(BrailleDomFocusRenderer);

                    ((IBrailleIOHookableRenderer)renderer).UnregisterHook(DrawSelectFocusRenderer);
                    ((IBrailleIOHookableRenderer)renderer).RegisterHook(DrawSelectFocusRenderer);

                    ((IBrailleIOHookableRenderer)renderer).UnregisterHook(TextRendererHook);
                    ((IBrailleIOHookableRenderer)renderer).RegisterHook(TextRendererHook);

                }
            }
        }

        #endregion

        #region Title and Description Dialog

        public const string TITLE_DESC_VIEW_NAME = "imageData";

        /// <summary>
        /// Opens Title and Description dialog for current selected shape in Detail area.
        /// </summary>
        private void openTitleDescDialog()
        {
            if (shapeManipulatorFunctionProxy == null || shapeManipulatorFunctionProxy.LastSelectedShape == null)
            {
                return;
            }

            // quit full screen to allow for showing dialog in detail region
            BrailleIOScreen vs = WindowManager.Instance.GetVisibleScreen();
            if (vs != null && vs.Name.Equals(WindowManager.BS_FULLSCREEN_NAME))
            {
                WindowManager.Instance.QuitFullscreen();
            }

            string dialogTitle = LL.GetTrans("tangram.lector.oo_observer.dialog.title_desc.dialogtitle", shapeManipulatorFunctionProxy.LastSelectedShape.Name);
            WindowManager.Instance.SetTopRegionContent(dialogTitle);
            AudioRenderer.Instance.PlaySoundImmediately(dialogTitle);

            //TODO class GUI-Dialog for Image
            //ImageData imageDataView = new ImageData(InteractionManager.BKI,io);
            InteractionManager.Instance.ChangeMode(InteractionMode.Braille);

            if (imgDataFunctionProxy != null && ScriptFunctionProxy.Instance != null)
            {
                ScriptFunctionProxy.Instance.AddProxy(imgDataFunctionProxy);
                imgDataFunctionProxy.Active = true;
                imgDataFunctionProxy.ShowImageDetailView(TITLE_DESC_VIEW_NAME, 15, 0, Property.Title, true);
            }
        }

        /// <summary>
        /// Speak description of focused element.
        /// </summary>
        private void speakDescription()
        {
            if (shapeManipulatorFunctionProxy == null || shapeManipulatorFunctionProxy.LastSelectedShape == null)
            {
                return;
            }
            string desc = "";
            if (String.IsNullOrWhiteSpace(shapeManipulatorFunctionProxy.LastSelectedShape.Description))
            {
                desc = LL.GetTrans("tangram.lector.oo_observer.dialog.title_desc.desc.no");
            }
            else desc = LL.GetTrans("tangram.lector.oo_observer.dialog.title_desc.desc", shapeManipulatorFunctionProxy.LastSelectedShape.Name, shapeManipulatorFunctionProxy.LastSelectedShape.Description);
            WindowManager.Instance.SetDetailRegionContent(desc);
            AudioRenderer.Instance.PlaySoundImmediately(desc);
        }

        #endregion
    }
}