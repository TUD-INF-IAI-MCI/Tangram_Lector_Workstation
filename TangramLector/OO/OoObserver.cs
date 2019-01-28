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
        internal readonly SelectionRendererHook DrawSelectFocusRenderer = new SelectionRendererHook(true);
        /// <summary>
        /// A renderer hook that overlays elements with text with a Braille label.
        /// </summary>
        public readonly BrailleTextView TextRendererHook = new BrailleTextView();

        readonly ImageData imgDataFunctionProxy = ImageData.Instance;

        #endregion

        readonly OpenOfficeDrawShapeManipulator shapeManipulatorFunctionProxy = null;
        public OpenOfficeDrawShapeManipulator OoManipulator { get { return shapeManipulatorFunctionProxy; } }
        #endregion

        #region Constructor

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
            //OoSelectionObserver.Instance.SelectionChanged += new EventHandler<OoSelectionChandedEventArgs>(OoDrawAccessibilityObserver_SelectionChanged);

            OoDrawAccessibilityObserver.Instance.DrawWindowOpend += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowOpend);
            OoDrawAccessibilityObserver.Instance.DrawWindowPropertyChange += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowPropertyChange);
            OoDrawAccessibilityObserver.Instance.DrawWindowMinimized += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowMinimized);
            OoDrawAccessibilityObserver.Instance.DrawWindowClosed += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowClosed);
            OoDrawAccessibilityObserver.Instance.DrawWindowActivated += new EventHandler<OoWindowEventArgs>(OoDrawAccessibilityObserver_DrawWindowActivated);
            OoDrawAccessibilityObserver.Instance.DrawSelectionChanged += new EventHandler<OoAccessibilitySelectionChangedEventArgs>(Instance_DrawSelectionChanged);

            shapeManipulatorFunctionProxy.SelectedShapeChanged += new EventHandler<SelectedShapeChangedEventArgs>(shapeManipulatorFunctionProxy_SelectedShapeChanged);
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

        #endregion

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

        #region Function Proxy Events

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

        void shapeManipulatorFunctionProxy_SelectedShapeChanged(object sender, SelectedShapeChangedEventArgs e)
        {
            if (sender != null && sender is OpenOfficeDrawShapeManipulator)
            {
                StopFocusHighlightModes();
                try
                {
                    if (e.Reason != ChangeReson.Property) ((OpenOfficeDrawShapeManipulator)sender).SayLastSelectedShape(false);

                    if (((OpenOfficeDrawShapeManipulator)sender).IsShapeSelected)
                    {
                        if (InteractionManager.Instance != null)
                            InteractionManager.Instance.ChangeMode(
                                InteractionManager.Instance.Mode | InteractionMode.Manipulation);

                        OoShapeObserver _shape = ((OpenOfficeDrawShapeManipulator)sender).LastSelectedShape;
                        if (_shape != null)
                        {
                            if (e.Reason != ChangeReson.Property) // shape changed 
                            {
                                BrailleDomFocusRenderer.CurrentPolyPoint = null;
                                focusHighlightPaused = false;
                            }
                            InitBrailleDomFocusHighlightMode(null, e.Reason != ChangeReson.Property);

                            if (e.Reason != ChangeReson.Property && ((OpenOfficeDrawShapeManipulator)sender).LastSelectedShapePolygonPoints == null) // in other cases the detailed infos for the  point is sent by the manipulator itself
                                WindowManager.Instance.SetDetailRegionContent(OoElementSpeaker.GetElementAudioText(_shape));
                            
                            if (WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_BRAILLE_FOCUS) 
                                jumpToDomFocus();
                            
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] new DOM focus " + _shape.Name + " " + _shape.Text);
                        }
                    }
                    else
                    {
                        if (InteractionManager.Instance != null)
                            InteractionManager.Instance.ChangeMode(
                                InteractionManager.Instance.Mode & ~InteractionMode.Manipulation);
                        WindowManager.Instance.SetDetailRegionContent(LL.GetTrans("tangram.lector.oo_observer.selected_no"));
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Log(LogPriority.OFTEN, this, "[DRAW INTERACTION] new DOM focus error --> shape is null", ex);
                }
            }
        }

        void shapeManipulatorFunctionProxy_PolygonPointSelected(object sender, PolygonPointSelectedEventArgs e)
        {
            if (sender != null && sender is OpenOfficeDrawShapeManipulator)
            {
                if (((OpenOfficeDrawShapeManipulator)sender).IsShapeSelected/*LastSelectedShape != null*/)
                {
                    //PauseFocusHighlightModes();
                    InitBrailleDomFocusHighlightMode();
                }
            }

            if (BrailleDomFocusRenderer != null)
            {
                if (e != null && e.PolygonPoints != null)
                {
                    BrailleDomFocusRenderer.CurrentPolyPoint = e.PolygonPoints;
                }
                else if (shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.IsShapeSelected && shapeManipulatorFunctionProxy.LastSelectedShapePolygonPoints != null)
                {
                    BrailleDomFocusRenderer.CurrentPolyPoint = shapeManipulatorFunctionProxy.LastSelectedShapePolygonPoints;
                }
                else
                {
                    BrailleDomFocusRenderer.CurrentPolyPoint = null;
                }
            }
        }

        #endregion

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

        #region Own Events

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
        public void StartDrawSelectFocusHighlightingMode(OoAccessibilitySelection selection = null)
        {
            if (OoDrawAccessibilityObserver.Instance != null)
            {
                if (selection == null)
                {
                    bool success = OoDrawAccessibilityObserver.Instance.TryGetSelection(GetActiveDocument(), out selection);
                }
            }

            this.DrawSelectFocusRenderer.DoRenderBoundingBox = true;
            this.DrawSelectFocusHighlightMode = true;
            this.DrawSelectFocusRenderer.SetSelection(selection);

            // TODO: center selection on the pin-matrix device
            //if (selection != null)
            //{
            //    System.Drawing.Rectangle bb = new Rectangle();
            //    bb = selection.SelectionBounds; // TODO: SelectionBounds are always 0
            //    if (windowManager != null && bb != null && bb.Width > 0 && bb.Height > 0) windowManager.MoveToObject(bb);
            //}
        }


        OoAccessibilitySelection _lastRegisterdSelection = null;

        //private List<OoShapeObserver> currentSelection = new List<OoShapeObserver>();
        void Instance_DrawSelectionChanged(object sender, OoAccessibilitySelectionChangedEventArgs e)
        {
            if (e != null /*&& e.SelectedItems != null*/)
            {
                if (windowManager.FocusMode == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
                {
                    if (e != null && sender != null && e.Source != null)
                    {
                        OoAccessibilitySelection selection;

                        bool success = OoDrawAccessibilityObserver.Instance.TryGetSelection(e.Source, out selection);
                        if (success)
                        {
                            if (selection != null && selection.Count > 0)
                            {
                                if (shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.Active
                                   && this.DrawSelectFocusRenderer != null)
                                {
                                    StartDrawSelectFocusHighlightingMode(selection);
                                }

                                if (
                                    (!e.Silent && (selection.Count > 0))
                                    || _lastRegisterdSelection == null
                                    || selection != _lastRegisterdSelection
                                    )
                                {
                                    CommunicateSelection(selection.SelectedItems);
                                }

                                _lastRegisterdSelection = selection;
                            }
                            else
                            {
                                _lastRegisterdSelection = null;
                                if (this.DrawSelectFocusRenderer != null)
                                    this.DrawSelectFocusRenderer.SetSelection(null);
                            }
                        }
                    }
                }
            }
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
                String audioText = LL.GetTrans("tangram.lector.oo_observer.selected_elements", e.Count);
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
                OoAccessibilitySelection selection = null;
                bool success = OoDrawAccessibilityObserver.Instance.TryGetSelection(GetActiveDocument(), out selection);
                if (success && selection != null)
                {
                    CommunicateSelection(selection.SelectedItems, immediately);
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
                    setTitelregionToDocTitle(e.Window);
                }
            }
            catch (System.Exception ex)
            {
                wm.ScreenObserver.ObserveScreen();
                Logger.Instance.Log(LogPriority.DEBUG, ex);
            }
        }

        private void setTitelregionToDocTitle(OoAccessibleDocWnd wnd)
        {
            if (windowManager != null && wnd != null)
            {
                string appTitle = wnd.Title;
                appTitle += getPageNumInfosOfWindow(wnd);
                windowManager.SetTopRegionContent(appTitle);
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
                        Thread.Sleep(50); // gives the element the chance to finish the changes before requesting the new properties. Should prevent a hang on. 

                        tud.mci.tangram.TangramLector.WindowManager wm = tud.mci.tangram.TangramLector.WindowManager.Instance;
                        try
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "ooDraw wnd property changed " + e.Window.ToString());

                            if (windowManager != null && e.Type.HasFlag(WindowEventType.ACTIVATED))
                            {
                                setTitelregionToDocTitle(e.Window);
                            }

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

        //void OoDrawAccessibilityObserver_SelectionChanged(object sender, OoSelectionChandedEventArgs e)
        //{
        //    //System.Diagnostics.Debug.WriteLine("Instance_SelectionChanged");
        //}

        #endregion

        #endregion

        #region Public Functions

        /// <summary>
        /// Resets this instance.
        /// </summary>
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
        /// Sets the shape for modification to the given observer.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="silent">if set to <c>true</c> no audio feedback about the selection is given.</param>
        /// <param name="immediatly">if set to <c>true</c> the audio feedback is immediately given and all other audio feedback is aborted.</param>
        /// <returns>
        /// the currently selected Shape observer
        /// </returns>
        public bool SetShapeForModification(OoShapeObserver shape, bool silent = true, bool immediately = true)
        {
            if (shape != null)
            {
                if (InteractionManager.Instance != null) InteractionManager.Instance.ChangeMode(InteractionManager.Instance.Mode | InteractionMode.Manipulation);

                if (shapeManipulatorFunctionProxy != null && !ImageData.Instance.Active)
                {
                    if (!silent)
                    {
                        if (immediately) OoElementSpeaker.PlayElementImmediately(shape, LL.GetTrans("tangram.lector.oo_observer.selected", String.Empty));
                        else OoElementSpeaker.PlayElement(shape, LL.GetTrans("tangram.lector.oo_observer.selected", String.Empty));
                    }
                    //audioRenderer.PlaySound("Form kann manipuliert werden");
                    shapeManipulatorFunctionProxy.LastSelectedShape = shape;
                    return shapeManipulatorFunctionProxy.LastSelectedShape == shape;
                }
                else // title+desc dialog handling
                {
                    ImageData.Instance.NewSelectionHandling(shape);
                }
            }
            else
            {
                if (InteractionManager.Instance != null)
                    InteractionManager.Instance.ChangeMode(
                        InteractionManager.Instance.Mode & ~InteractionMode.Manipulation);
            }

            return false;
        }

        /// <summary>
        /// Sets the shape for modification to the given observer.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="observed">The observed.</param>
        /// <returns>the currently selected Shape observer</returns>
        public bool SetPolypointForModification(OoPolygonPointsObserver points, OoShapeObserver shape)
        {
            if (shape != null && points != null && points.Shape == shape)
            {
                if (shapeManipulatorFunctionProxy != null && !ImageData.Instance.Active)
                {
                    // OoElementSpeaker.PlayElementImmediately(shape, LL.GetTrans("tangram.lector.oo_observer.selected", String.Empty));
                    if (shapeManipulatorFunctionProxy.LastSelectedShape != shape)
                    {
                        shapeManipulatorFunctionProxy.LastSelectedShape = shape;
                    }
                    shapeManipulatorFunctionProxy.LastSelectedShapePolygonPoints = points;
                    // shapeManipulatorFunctionProxy.SelectLastPolygonPoint();
                    shapeManipulatorFunctionProxy.SelectPolygonPoint();
                }
                else // title+desc dialog handling
                {
                    ImageData.Instance.NewSelectionHandling(shape);
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the shape for modification from the given Draw page child object.
        /// </summary>
        /// <param name="c">The component to search its shape observer for [<see cref="OoAccComponent"/>], [<see cref="OoShapeObserver"/>], [XShape].</param>
        /// <param name="observed">The observed window / draw document.</param>
        /// <param name="silent">if set to <c>true</c> no audio feedback about the selection is given.</param>
        /// <param name="immediatly">if set to <c>true</c> the audio feedback is immediately given and all other audio feedback is aborted.</param>
        /// <returns>The corresponding registered shape observer</returns>
        public OoShapeObserver GetShapeForModification(Object c, OoAccessibleDocWnd observed, OoDrawPageObserver page = null, bool silent = false, bool immediately = true)
        {
            if (c is OoAccComponent)
            {
                return GetShapeForModification(c as OoAccComponent, observed, silent, immediately);
            }
            else if (c is OoShapeObserver)
            {
                return GetShapeForModification(c as OoShapeObserver, observed, silent, immediately);
            }
            else if (observed != null)
            {
                return observed.GetRegisteredShapeObserver(c, page != null ? page : observed.GetActivePage());
            }

            return null;
        }


        /// <summary>
        /// Sets the shape for modification to the given observer.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="observed">The observed.</param>
        /// <param name="silent">if set to <c>true</c> no audio feedback about the selection is given.</param>
        /// <param name="immediatly">if set to <c>true</c> the audio feedback is immediately given and all other audio feedback is aborted.</param>
        /// <returns>the currently selected Shape observer</returns>
        public OoShapeObserver GetShapeForModification(OoShapeObserver shape, OoAccessibleDocWnd observed = null, bool silent = true, bool immediately = true)
        {
            return SetShapeForModification(shape, silent, immediately) ? shape : null;
        }

        /// <summary>
        /// Gets the shape observer for modification.
        /// </summary>
        /// <param name="c">The component to get the corresponding observer to.</param>
        /// <param name="observed">The observed window / draw document.</param>
        /// <param name="silent">if set to <c>true</c> no audio feedback about the selection is given.</param>
        /// <param name="immediatly">if set to <c>true</c> the audio feedback is immediately given and all other audio feedback is aborted.</param>
        /// <returns>The corresponding observer to the shape in the given Draw document.</returns>
        public OoShapeObserver GetShapeForModification(OoAccComponent c, OoAccessibleDocWnd observed, bool silent = true, bool immediately = true)
        {
            if (observed != null && c.Role != AccessibleRole.INVALID)
            {
                //TODO: prepare name:

                OoShapeObserver shape = observed.GetRegisteredShapeObserver(c);
                if (shape != null)
                {
                    return GetShapeForModification(shape, observed, silent, immediately);

                    //if (shapeManipulatorFunctionProxy != null && !ImageData.Instance.Active)
                    //{
                    //    OoElementSpeaker.PlayElementImmediately(shape, LL.GetTrans("tangram.lector.oo_observer.selected", String.Empty));
                    //    //audioRenderer.PlaySound("Form kann manipuliert werden");
                    //    shapeManipulatorFunctionProxy.LastSelectedShape = shape;
                    //    return shape;
                    //}
                    //else // title+desc dialog handling
                    //{
                    //    ImageData.Instance.NewSelectionHandling(shape);
                    //}
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
        /// Check if a shape is selected for manipulation (DOM / Braille focus).
        /// </summary>
        /// <returns><c>true</c> if a shape was selected for manipulation.</returns>
        public bool IsShapeSelected()
        {
            return shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.IsShapeSelected;
        }

        /// <summary>
        /// Gets the last selected shape (DOM / Braille focus).
        /// You should call <see cref="IsShapeSelected"/> if you only want to know if a shape is selected.
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
            if (OoDrawAccessibilityObserver.Instance != null)
            {
                OoAccessibilitySelection selection = null;
                bool success = OoDrawAccessibilityObserver.Instance.TryGetSelection(GetActiveDocument(), out selection);

                if (success && selection != null && selection.Count > 0)
                {
                    return selection.SelectedItems[0];
                }
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

        /// <summary>
        /// Gets the infos about the amount of pages and the current page number of a window.
        /// </summary>
        /// <param name="win">The DRAW doc window.</param>
        /// <returns>a String of type ' X/Y', where X is the current page and Y the amount of pages; otherwise the empty string.</returns>
        private static string getPageNumInfosOfWindow(OoAccessibleDocWnd win)
        {
            String result = String.Empty;
            if (win != null)
            {
                int pC = win.GetPageCount();
                if (pC > 1)
                {
                    var aPObs = win.GetActivePage();
                    if (aPObs != null)
                    {
                        int cP = aPObs.GetPageNum();
                        if (cP > 0)
                        {
                            result += " " + cP + "/" + pC;
                        }
                    }
                }
            }
            return result;
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

            OoShapeObserver shape = null; // FIXME: check that // SelectedItem;
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
            if (shapeManipulatorFunctionProxy == null || !shapeManipulatorFunctionProxy.IsShapeSelected /*.LastSelectedShape == null*/)
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
            if (shapeManipulatorFunctionProxy == null || shapeManipulatorFunctionProxy.IsShapeSelected /*.LastSelectedShape == null*/)
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