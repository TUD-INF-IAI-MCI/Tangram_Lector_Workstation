using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BrailleIO;
using System.Drawing;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.audio;

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

        internal void InitBrailleDomFocusHighlightMode()
        {
            if (shapeManipulatorFunctionProxy != null)
            {
                OoShapeObserver _shape = shapeManipulatorFunctionProxy.LastSelectedShape;

                if (_shape != null && _shape.IsValid())
                {
                    startFocusHighlightModes();

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

        private void startFocusHighlightModes()
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
                    brailleDomFocusHighlightMode = true;
                    blinkFocusActive = true;
                }
            }
        }

        private void stopFocusHighlightModes()
        {

            if (shapeManipulatorFunctionProxy != null)
            {
                OoShapeObserver _shape = shapeManipulatorFunctionProxy.LastSelectedShape;

                if (_shape != null && _shape.IsValid())
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
            }
        }

        private int invertRGBColor(int rbgColor)
        {
            int r = (rbgColor >> 16) & 0xFF;
            int g = (rbgColor >> 8) & 0xFF;
            int b = (rbgColor) & 0xFF;
            return (0xFF - r << 16) + (0xFF - g << 8) + (0xFF - b);
        }

        void blinkTimer_Tick(object sender, EventArgs e)
        {
            // change blink state
            blinkStateOn = !blinkStateOn;

            if (brailleDomFocusHighlightMode && shapeManipulatorFunctionProxy != null && shapeManipulatorFunctionProxy.LastSelectedShape != null)
            {
                // blink rendered bounding box
                BrailleDomFocusRenderer.DoRenderBoundingBox = blinkStateOn;

                //doFocusHighlighting(blinkStateOn);
                BrailleIOMediator.Instance.RefreshDisplay();
            }
            if (DrawSelectFocusHighlightMode && WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
            {
                DrawSelectFocusRenderer.DoRenderBoundingBox = blinkStateOn;

                BrailleIOMediator.Instance.RefreshDisplay();
            }
            else
            {
                DrawSelectFocusHighlightMode = false;
                DrawSelectFocusRenderer.DoRenderBoundingBox = false;
                BrailleIOMediator.Instance.RefreshDisplay();
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
                                //play("Ich kann das Element nicht richtig auf dem Bildschirm finden.");
                                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] can't find a valid accessible counterpart for the dom focused element: " + shapeManipulatorFunctionProxy.LastSelectedShape);
                            }
                        }
                        else
                        {
                            playError();
                            //play("Ich kann das Element nicht auf dem Bildschirm finden.");
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
                Rectangle newBBox = shapeManipulatorFunctionProxy.LastSelectedShape.GetAbsoluteScreenBoundsByDom();
                DesktopOverlayWindow.Instance.refreshBounds(newBBox, ref pngData);
                BrailleDomFocusRenderer.CurrentBoundingBox = newBBox;
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
