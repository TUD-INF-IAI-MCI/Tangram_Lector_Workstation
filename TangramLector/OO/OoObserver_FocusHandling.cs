using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BrailleIO;
using System.Drawing;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.audio;
using System.Threading;

namespace tud.mci.tangram.TangramLector.OO
{
    public partial class OoObserver
    {

        #region Focus Handling

        #region Member

        private bool brailleDomFocusHighlightMode = true;
        /// <summary>
        /// brailleDomFocusRenderer
        /// </summary>
        public bool DrawSelectFocusHighlightMode = true;
        private bool blinkStateOn = false;
        /// <summary>
        /// true if braille focus is blinking, false if braille focus is not blinking
        /// </summary>
        private bool blinkFocusActive = false;

        // original attributes of selected element to be restored on focus change / end of highlight
        // TODO: have to be changed live when last observed shape attributes are changed by draw!
        private int oLineTransparence = 0;
        private int oLineColor = 0;
        private int oFillTransparence = 0;
        private int oFillColor = 0;

        private tud.mci.tangram.controller.observer.OoShapeObserver.BoundRectChangeEventHandler OnShapeBoundRectChange;
        private tud.mci.tangram.controller.observer.OoDrawPagesObserver.ViewOrZoomChangeEventHandler OnViewOrZoomChange;

        #endregion

        /// <summary>
        /// Sets the shape or point to blink and starts the blinking frame around the last selected shape.
        /// </summary>
        internal void InitBrailleDomFocusHighlightMode(OoShapeObserver _shape = null)
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                if(_shape == null) _shape = shapeManipulatorFunctionProxy.LastSelectedShape;

                if (_shape != null && _shape.IsValid(false))
                {
                    StartFocusHighlightModes();

                    // inform sighted user 
                    byte[] pngData;
                    if (_shape.GetShapeAsPng(out pngData) > 0)
                    {
                        DesktopOverlayWindow.Instance.initBlinking(
                            ref pngData,                    // png as byte array
                            _shape.GetAbsoluteScreenBoundsByDom(),   // bounding box
                            0.75,                           // opacity (0.0 (transparent) .. 1.0)
                            1500,                           // The total blinking time in ms until blinking stops and the window is hidden again
                            //The pattern of milliseconds to be on|off|..., 
                            // e.g. [250,150,250,150,250,150,800,1000] for three short flashes (on for 250ms) 
                            // with gaps (of 150ms) and one long on time (800ms) followed by a pause(1s).
                            // The pattern is repeated until the total blinking time is over.
                            // If less than 2 values are given, the default of [500, 500] is used!
                            new int[8] { 250, 150, 250, 150, 250, 150, 500, 150 }
                            );
                    }
                }
            }
        }

        /// <summary>
        /// Starts the focus highlight modes.
        /// </summary>
        internal void StartFocusHighlightModes()
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                OoShapeObserver _shape = shapeManipulatorFunctionProxy.LastSelectedShape;

                if (_shape != null)
                {
                    _shape.BoundRectChangeEventHandlers += OnShapeBoundRectChange;
                    _shape.Page.PagesObserver.ViewOrZoomChangeEventHandlers += OnViewOrZoomChange;
                    // TODO: listen to changes by oo draw!
                    //oFillTransparence = _shape.FillTransparence;
                    //oLineTransparence = _shape.LineTransparence;
                    //oFillColor = _shape.FillColor;
                    //oLineColor = _shape.LineColor;
                    // start highlighting focus (may be switched off later, e.g. by timer or key press

                    BrailleDomFocusRenderer.SetCurrentBoundingBoxByShape(_shape);
                    
                    brailleDomFocusHighlightMode = true;
                    blinkFocusActive = true;
                }
            }
        }

        /// <summary>
        /// Stops / pauses the focus highlight modes.
        /// </summary>
        internal void StopFocusHighlightModes()
        {

            if (shapeManipulatorFunctionProxy != null)
            {
                OoShapeObserver _shape = shapeManipulatorFunctionProxy.LastSelectedShape;
                var point = shapeManipulatorFunctionProxy.LastSelectedShapePolygonPoints;

                //if (point != null && !point.IsEmpty)
                //{
                //    // make a timeout ;-) and start again
                //    pausePointHighligting();
                //}

                if (_shape != null)
                {
                    if (OnShapeBoundRectChange != null)
                    {
                        _shape.BoundRectChangeEventHandlers -= OnShapeBoundRectChange;
                        _shape.Page.PagesObserver.ViewOrZoomChangeEventHandlers -= OnViewOrZoomChange;
                    }
                    //_shape.FillTransparence = (short)(oFillTransparence);
                    //_shape.LineTransparence = oLineTransparence;
                    //_shape.FillColor = oFillColor;
                    //_shape.LineColor = oLineColor;
                }
                blinkStateOn = false;
                brailleDomFocusHighlightMode = false;
                BrailleDomFocusRenderer.DoRenderBoundingBox = false;
                DrawSelectFocusHighlightMode = false;
                DrawSelectFocusRenderer.DoRenderBoundingBox = false;
                blinkFocusActive = false;

                PauseFocusHighlightModes();

            }

        }

        #region Pause Focus Highlighting

        private const string FOCUS_HIGHLIGHT_PAUSE_CONFIG_KEY = "FocusHighlightPause";

        private static int getFocusTimerPauseTimeout()
        {
            int timeout = 5000;

            // load the timeout from the app.config
            try
            {
                var config = System.Configuration.ConfigurationManager.AppSettings;
                if (config != null && config.Count > 0)
                {
                    var value = config[FOCUS_HIGHLIGHT_PAUSE_CONFIG_KEY];
                    if (!String.IsNullOrWhiteSpace(value))
                        timeout = Convert.ToInt32(value);
                }
            }
            catch { }

            return timeout;
        }

        Timer fhpt = null;
        Timer focusHighlightPauseTimer
        {
            get
            {
                //if (fhpt == null) {
                //    fhpt = new Timer(focusHighlightPausTimerCB, null, 0, Timeout.Infinite);
                //}
                return fhpt;
            }
            set
            {
                if (fhpt != null)
                {
                    fhpt.Dispose();
                    fhpt = null;
                }
                fhpt = value;
            }
        }

        /// <summary>
        /// Pauses the focus highlight modes.
        /// </summary>
        internal void PauseFocusHighlightModes()
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                if (BrailleDomFocusRenderer != null)
                {
                    if (BrailleDomFocusRenderer.DoRenderBoundingBox == false)
                    {
                        // check if the timer is already running
                        focusHighlightPauseTimer = new Timer(focusHighlightPausTimerCB, null, getFocusTimerPauseTimeout(), Timeout.Infinite);
                    }
                }
            }
        }

        void focusHighlightPausTimerCB(object status)
        {
            focusHighlightPauseTimer = null;
            StartFocusHighlightModes();
        }

        #endregion


        private int invertRGBColor(int rbgColor)
        {
            int r = (rbgColor >> 16) & 0xFF;
            int g = (rbgColor >> 8) & 0xFF;
            int b = (rbgColor) & 0xFF;
            return (0xFF - r << 16) + (0xFF - g << 8) + (0xFF - b);
        }

        volatile bool _render = false;
        void blinkTimer_Tick(object sender, EventArgs e)
        {
            // bring is some delay to not interfere with the rendering itself
            System.Threading.Thread.Sleep(95);

            // change blink state
            blinkStateOn = !blinkStateOn;

            if (brailleDomFocusHighlightMode && shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.LastSelectedShape != null)
            {
                _render = true;
                // blink rendered bounding box
                BrailleDomFocusRenderer.DoRenderBoundingBox = blinkStateOn;

                //doFocusHighlighting(blinkStateOn);
                // BrailleIOMediator.Instance.RefreshDisplay(true);
            }
            else
                if (DrawSelectFocusHighlightMode && WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
                {
                    _render = true;
                    DrawSelectFocusRenderer.DoRenderBoundingBox = blinkStateOn;
                    // BrailleIOMediator.Instance.RefreshDisplay(true);
                }
                else
                {
                    DrawSelectFocusHighlightMode = false;
                    DrawSelectFocusRenderer.DoRenderBoundingBox = false;

                    if (_render)
                    {
                        _render = false;
                        // BrailleIOMediator.Instance.RefreshDisplay(true);
                    }
                }
        }

        /// <summary>
        /// Highlights the focused element by realizing a blinking of the shape.
        /// </summary>
        /// <param name="blinkStateOn"></param>
        private void doFocusHighlightingByOsmColorManipulation(bool blinkStateOn)
        {
            if (shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.LastSelectedShape != null)
            {
                // blink dom object (color inversion + transparency inversion)
                if (blinkStateOn)
                {
                    // original
                    shapeManipulatorFunctionProxy.LastSelectedShape.FillTransparence = (short)(oFillTransparence);
                    shapeManipulatorFunctionProxy.LastSelectedShape.LineTransparence = oLineTransparence;
                    shapeManipulatorFunctionProxy.LastSelectedShape.FillColor = oFillColor;
                    shapeManipulatorFunctionProxy.LastSelectedShape.LineColor = oLineColor;
                }
                else
                {
                    // inverted
                    shapeManipulatorFunctionProxy.LastSelectedShape.FillTransparence = (short)(100 - oFillTransparence);
                    shapeManipulatorFunctionProxy.LastSelectedShape.LineTransparence = 100 - oLineTransparence;
                    shapeManipulatorFunctionProxy.LastSelectedShape.FillColor = invertRGBColor(oFillColor);
                    shapeManipulatorFunctionProxy.LastSelectedShape.LineColor = invertRGBColor(oLineColor);
                }
            }
        }

        /// <summary>
        /// Centers the focused DOM object in the middle of the view range.
        /// </summary>
        /// <returns></returns>
        private bool jumpToDomFocus()
        {
            if (shapeManipulatorFunctionProxy == null || shapeManipulatorFunctionProxy.LastSelectedShape == null)
            {
                playError();
                return false;
            }
            WindowManager wm = WindowManager.Instance;
            if (wm != null)
            {
                var vScreen = wm.GetVisibleScreen();
                // do this only by known specializes Screen that are related to the OO-window
                if (vScreen != null && (vScreen.Name == WindowManager.BS_FULLSCREEN_NAME || vScreen.Name == WindowManager.BS_MAIN_NAME))
                {
                    //TODO: check if OO is shown
                    BrailleIOViewRange vr = null;
                    if (vScreen.HasViewRange(WindowManager.VR_CENTER_NAME)) vr = vScreen.GetViewRange(WindowManager.VR_CENTER_NAME);
                    else if (vScreen.HasViewRange(WindowManager.VR_CENTER_2_NAME)) vr = vScreen.GetViewRange(WindowManager.VR_CENTER_2_NAME);
                    if (vr != null)
                    {
                        // jump to element position
                        // try to get the accessible view to the element - it's the
                        // only way to get pixel positions on screen
                        if (shapeManipulatorFunctionProxy.LastSelectedShape.IsVisible())
                        {
                            var bounds = shapeManipulatorFunctionProxy.LastSelectedShape.GetRelativeScreenBoundsByDom();
                            if (!bounds.IsEmpty && bounds.Height > 0 && bounds.Width > 0)
                            {
                                wm.MoveToObject(bounds);
                                brailleDomFocusHighlightMode = true;
                                return true;
                            }
                            else
                            {
                                playError();
                                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] can't find a valid accessible counterpart for the dom focused element: " + shapeManipulatorFunctionProxy.LastSelectedShape);
                            }
                        }
                        else
                        {
                            playError();
                            Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] can't find the accessible counterpart to the dom focused element: " + shapeManipulatorFunctionProxy.LastSelectedShape);
                        }
                    }
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] can't find the Screen or ViewRannge to synchronize the DOM focus with");
                }
            }
            else
            {
                Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] no WindowManager instance available!!");
            }
            return false;
        }

        void ShapeBoundRectChangeHandler()
        {
            if (shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.LastSelectedShape != null)
            {
                byte[] pngData;
                shapeManipulatorFunctionProxy.LastSelectedShape.GetShapeAsPng(out pngData);
                if (pngData != null && pngData.Length > 0)
                {
                    Rectangle newBBox = shapeManipulatorFunctionProxy.LastSelectedShape.GetAbsoluteScreenBoundsByDom();
                    DesktopOverlayWindow.Instance.refreshBounds(newBBox, ref pngData);
                    BrailleDomFocusRenderer.CurrentBoundingBox = newBBox;

                    GC.AddMemoryPressure(pngData.Length);
                }
                else
                {
                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] cant get png data overlay picture for BrailleFocus");
                }
            }
        }

        #endregion

        #region Audio

        /// <summary>
        /// Plays a sound indicating an error has occurred.
        /// </summary>
        private static void playError() { AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error); }

        #endregion

    }
}
