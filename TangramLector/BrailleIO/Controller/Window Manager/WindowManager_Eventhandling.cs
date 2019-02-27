using BrailleIO;
using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Automation;
using tud.mci.tangram.audio;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.Uia;
using tud.mci.tangram.util;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager
    {

        #region Events

        void registerForEvents()
        {
            // BRAILLE IO 
            if (io != null && io.AdapterManager != null)
            {
                io.AdapterManager.ActiveAdapterChanged += AdapterManager_ActiveAdapterChanged;
                io.AdapterManager.NewAdapterRegistered += AdapterManager_NewAdapterRegistered;
            }

            // INTERACTION MANGER
            if (InteractionManager != null)
            {
                // load the button mapping
                InteractionManager.AddButton2FunctionMapping(Properties.Resources.GlobalFunctionMappings);


                registerAsSpecializedFunctionProxy();
                InteractionManager.InteractionModeChanged += new EventHandler<InteractionModeChangedEventArgs>(InteractionManager_InteractionModeChanged);
            }
            registerTopMostSpFProxy();
        }

        void AdapterManager_NewAdapterRegistered(object sender, IBrailleIOAdapterEventArgs e)
        {
            if (e != null && e.Adapter != null && e.Adapter.Device != null)
            {
                InteractionManager.AddNewDevice(e.Adapter);
                // TODO: monitor new adapter!
            }
        }

        void AdapterManager_ActiveAdapterChanged(object sender, IBrailleIOAdapterEventArgs e)
        {
            //CleanScreen();
            //BuildScreens();            
            UpdateScreens();
        }

        void InteractionManager_InteractionModeChanged(object sender, InteractionModeChangedEventArgs e)
        {
            if (e != null && (e.NewValue.Equals(InteractionMode.Braille) || e.OldValue.Equals(InteractionMode.Braille)))
            {
                updateStatusRegionContent();
            }
            else if (e != null && (e.NewValue.Equals(InteractionMode.Gesture) || e.OldValue.Equals(InteractionMode.Gesture)))
            {
                updateStatusRegionContent();
            }
            else if (e != null && (e.NewValue.Equals(InteractionMode.Manipulation) || e.OldValue.Equals(InteractionMode.Manipulation)))
            {
                updateStatusRegionContent();
            }
        }

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        { }

        protected override void im_FunctionCall(object sender, FunctionCallInteractionEventArgs e)
        {
            if (e != null && !string.IsNullOrEmpty(e.Function))
            {
                BrailleIOScreen vs = GetVisibleScreen();

                #region General Commands

                switch (e.Function)
                {
                    case "abortSpeechOutput": // abort speech output
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "stop audio");
                        AbortAudio();
                        e.Cancel = true;
                        e.Handled = true;
                        return;

                    case "recalibrate":
                        // calibrate BrailleDis //
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "recalibrate devices");
                        if (io != null)
                        {
                            // this.ScreenObserver.Stop();
                            audioRenderer.PlaySound(LL.GetTrans("tangram.lector.wm.recalibrate"));
                            // bool res = io.Recalibrate();
                            bool res = io.RecalibrateAll();
                            // this.ScreenObserver.Start();
                            audioRenderer.PlaySound(LL.GetTrans("tangram.lector.wm.recalibrate.success", (res ? "" : LL.GetTrans("tangram.lector.not") + " ")));
                        }
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] recalibrate the device");
                        e.Cancel = true;
                        e.Handled = true;
                        return;

                    case "followGUIFocus":
                        // follow GUI focus mode //
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "start follow GUI focus mode");
                        bool communicateSelection = false;
                        if (this.FocusMode == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
                        {
                            this.FocusMode = FollowFocusModes.NONE;
                        }
                        else
                        {
                            this.FocusMode = FollowFocusModes.FOLLOW_MOUSE_FOCUS;

                            // jump to selected element
                            OoConnector ooc = OoConnector.Instance;
                            if (ooc != null && ooc.Observer != null)
                            {
                                System.Drawing.Rectangle bb = new Rectangle();
                                if (OoDrawAccessibilityObserver.Instance != null)
                                {
                                    OoAccessibilitySelection selection = null;
                                    bool success = OoDrawAccessibilityObserver.Instance.TryGetSelection(ooc.Observer.GetActiveDocument(), out selection);
                                    if (success && selection != null)
                                    {
                                        bb = selection.SelectionBounds;
                                    }
                                }

                                ooc.Observer.StartDrawSelectFocusHighlightingMode();
                                if (bb != null && bb.Width > 0 && bb.Height > 0) MoveToObject(bb);
                                communicateSelection = true;
                            }
                        }
                        if (communicateSelection
                            && OoConnector.Instance != null
                            && OoConnector.Instance.Observer != null)
                        {
                            OoConnector.Instance.Observer.CommunicateLastSelection(false);
                        }
                        e.Handled = true;
                        e.Cancel = true;
                        return;
                }

                #endregion

                #region minimap commands
                if (vs != null && vs.Name.Equals(BS_MINIMAP_NAME))
                {
                    switch (e.Function)
                    {
                        case "toggleMinimap":
                        //case "exitMinimap":
                            break;
                        case "invert":
                            break;
                        case "pannPageLeft":
                        case "pannStepLeft":
                        case "pannPageRight":
                        case "pannStepRight":
                        case "pannPageUp":
                        case "pannStepUp":
                        case "pannPageDown":
                        case "pannStepDown":
                        case "pannJumpToLeft":
                        case "pannJumpToRight":
                        case "pannJumpToTop":
                        case "pannJumpToBottom":
                            vs = screenBeforeMinimap;
                            break;
                        default:
                            vs = screenBeforeMinimap;
                            audioRenderer.PlayWaveImmediately(StandardSounds.Error);
                            e.Cancel = true;
                            return;
                    }
                }



                #endregion

                #region global button commands

                BrailleIOViewRange center = GetActiveCenterView(vs);
                BrailleIOViewRange detail = null;
                if (vs != null) detail = vs.GetViewRange(VR_DETAIL_NAME);

                switch (e.Function)
                {
                    #region panning operations
                    case "pannPageLeft":// links blättern
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large panning left");
                        if (vs != null && center != null)
                            if (moveHorizontal(vs.Name, center.Name, center.ContentBox.Width))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.left_big"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll left");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannStepLeft":// links verschieben
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "panning left");
                        if (vs != null && center != null)
                            if (moveHorizontal(vs.Name, center.Name, 5))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.left"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll left");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannStepRight":// rechts verschieben
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "panning right");
                        if (vs != null && center != null)
                            if (moveHorizontal(vs.Name, center.Name, -5))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.right"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll right");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannPageRight":// rechts blättern
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large panning right");
                        if (vs != null && center != null)
                            if (moveHorizontal(vs.Name, center.Name, -center.ContentBox.Width))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.right_big"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll right");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannPageUp":// hoch blättern
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large panning up");
                        if (vs != null && center != null)
                            if (moveVertical(vs.Name, center.Name, center.ContentBox.Height))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.up_big"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll up");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannStepUp":// hoch verschieben
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "panning up");
                        if (vs != null && center != null)
                            if (moveVertical(vs.Name, center.Name, 5))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.up"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll up");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannStepDown":// runter verschieben
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "panning down");
                        if (vs != null && center != null)
                            if (moveVertical(vs.Name, center.Name, -5))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.down"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll down");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannPageDown":// runter blättern
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large panning down");
                        if (vs != null && center != null)
                            if (moveVertical(vs.Name, center.Name, -center.ContentBox.Height))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.down_big"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll down");
                            }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannJumpToTop":// Sprung nach oben
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "jump to top");
                        jumpToTopBorder(center);
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannJumpToLeft":// Sprung nach links
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "jump to left");
                        jumpToLeftBorder(center);
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannJumpToRight":// Sprung nach rechts
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "jump to right");
                        jumpToRightBorder(center);
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "pannJumpToBottom":// Sprung nach unten
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "jump to bottom");
                        jumpToBottomBorder(center);
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    #endregion

                    #region zooming operations

                    case "zoomIncrease": // kleiner Zoom in
                        if (vs != null && center != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "zoom in");
                            if (zoomWithFactor(vs.Name, center.Name, 1.5))
                            {
                                var zInProzent = GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name);
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.wm.zooming.in", zInProzent));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small zoom in");
                            }
                            else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                        }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "zoomDecrease": // kleiner Zoom out
                        if (vs != null && center != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "zoom out");
                            if (zoomWithFactor(vs.Name, center.Name, 1 / 1.5))
                            {
                                var zInProzent = GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name);
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.wm.zooming.out", zInProzent));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small zoom out");
                            }
                            else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        return;

                    case "zoomIncreaseLarge": // großer Zoom in
                        if (vs != null && center != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large zoom in");
                            if (zoomWithFactor(vs.Name, center.Name, 3))
                            {
                                var zInProzent = GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name);
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.wm.zooming.in_big", zInProzent));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big zoom in");
                            }
                            else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                        }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    case "zoomDecreaseLarge": // großer Zoom out
                        if (vs != null && center != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "large zoom out");
                            if (zoomWithFactor(vs.Name, center.Name, (double)1 / 3))
                            {
                                var zInProzent = GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name);
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.wm.zooming.out_big", zInProzent));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big zoom out");
                            }
                            else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                        }
                        e.Handled = true;
                        e.Cancel = true;
                        return;

                    #endregion

                    #region detail area scrolling

                    case "pannDetailUp": // scroll detail region up
                        if (detail != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "detail area panning up");
                            if (scrollViewRange(detail, 5))
                            {
                                e.Handled = true;
                                e.Cancel = true;
                                break;
                            }
                        }
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        break;

                    case "pannDetailLeft": // no function yet
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        break;

                    case "pannDetailRight": // no function yet
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        break;

                    case "pannDetailDown": // scroll detail region down
                        if (detail != null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "detail area panning down");
                            if (scrollViewRange(detail, -5))
                            {
                                e.Handled = true;
                                e.Cancel = true;
                                break;
                            }
                        }
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        break;

                    #endregion
                    default:
                        break;

                }
                #endregion

                #region Drawing mode button commands

                // button commands that are not working in Braille mode
                if (!InteractionManager.Mode.Equals(InteractionMode.Braille))
                {
                    switch (e.Function)
                    {
                        case "zoomTo100": // Druck-Zoom

                            if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "zoom to print zoom");
                                if (ZoomTo(vs.Name, VR_CENTER_NAME, GetPrintZoomLevel()))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.to_print"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] print zoom");
                                }
                                else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "toggleMinimap": // Minimap
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "toggle minimap");
                            toggleMinimap();
                            BrailleIOScreen nvs = GetVisibleScreen();
                            if (nvs != null && nvs.Name.Equals(BS_MINIMAP_NAME))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.minimap.show"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] start minimap");
                            }
                            else if (nvs != null)
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.minimap.hide"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] stop minimap");
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "zoomToFit": // Breite/Höhe einpassen
                            if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "zoom to fit");
                                if (ZoomTo(vs.Name, VR_CENTER_NAME, -1))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.fit_height"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] zoom fit to height");
                                }
                                else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "zoomTo1to1": // 1-zu-1 Zoom
                            if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[NAVIGATION]\t" + "zoom to 1:1");
                                if (ZoomTo(vs.Name, VR_CENTER_NAME, 1))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.1-1"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] zoom 1:1");
                                }
                                else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "invert": // Invertieren
                            if (vs != null)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "invert view");
                                invertImage(vs.Name, VR_CENTER_NAME);
                                BrailleIOViewRange vr = vs.GetViewRange(VR_CENTER_NAME);
                                if (vr != null && !vr.InvertImage)
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.non_invert"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] not invert output");
                                }
                                else
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.invert"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] invert output");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "decreaseBrightnessThreshold": // Schwellwert minus
                            if (vs != null)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "decrease threshold");
                                if (updateContrast(vs.Name, VR_CENTER_NAME, -THRESHOLD_STEP))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.threshold_down"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] reduce threshold");
                                }
                                else
                                {
                                    audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] already mininum threshold");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "increaseBrightnessThreshold": // Schwellwert plus
                            if (vs != null)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "increase threshold");
                                if (updateContrast(vs.Name, VR_CENTER_NAME, THRESHOLD_STEP))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.threshold_up"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] increase threshold");
                                }
                                else
                                {
                                    audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] already maximum threshold");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        // Standard-Schwellwert
                        case "resetBrightnessThreshold":
                            if (vs != null)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "reset threshold");
                                setContrast(vs.Name, VR_CENTER_NAME, STANDARD_CONTRAST_THRESHOLD);
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.threshold_reset"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] reset threshold");
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        // Kopfbereich ein/aus
                        case "toggleHeaderRegion":
                            if (vs != null && vs.GetViewRange(VR_TOP_NAME) != null)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "toggle header region");
                                if (vs.GetViewRange(VR_TOP_NAME).IsVisible())
                                {
                                    if (changeViewVisibility(vs.Name, VR_TOP_NAME, false))
                                    {
                                        changeViewVisibility(vs.Name, VR_STATUS_NAME, false);
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.views.region.hide",
                                            LL.GetTrans("tangram.lector.wm.views.region.VR_TOP_NAME")
                                            ));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] hide top region");
                                    }
                                }
                                else if (changeViewVisibility(vs.Name, VR_TOP_NAME, true))
                                {
                                    changeViewVisibility(vs.Name, VR_STATUS_NAME, true);
                                    audioRenderer.PlaySoundImmediately(
                                        LL.GetTrans("tangram.lector.wm.views.region.show",
                                            LL.GetTrans("tangram.lector.wm.views.region.VR_TOP_NAME")
                                            ));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] show top region");
                                }
                            }
                            else if (vs != null && vs.Name.Equals(BS_FULLSCREEN_NAME))
                            {
                                // exit fullscreen mode and show top region
                                toggleFullscreen();
                                BrailleIOScreen __nvs = GetVisibleScreen();
                                if (__nvs != null)
                                {
                                    BrailleIOViewRange top = __nvs.GetViewRange(VR_TOP_NAME);
                                    if (top != null && !top.IsVisible())
                                    {
                                        changeViewVisibility(__nvs.Name, VR_TOP_NAME, true);
                                    }
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        // Detailbereich ein/aus
                        case "toggleDetailRegion":
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "toggle detail region");
                            if (vs != null && vs.GetViewRange(VR_DETAIL_NAME) != null)
                            {
                                if (vs.GetViewRange(VR_DETAIL_NAME).IsVisible())
                                {
                                    if (changeViewVisibility(vs.Name, VR_DETAIL_NAME, false))
                                    {
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.views.region.hide",
                                            LL.GetTrans("tangram.lector.wm.views.region.VR_DETAIL_NAME")
                                            ));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] hide detail region");
                                    }
                                }
                                else if (changeViewVisibility(vs.Name, VR_DETAIL_NAME, true))
                                {
                                    audioRenderer.PlaySoundImmediately(
                                        LL.GetTrans("tangram.lector.wm.views.region.show",
                                            LL.GetTrans("tangram.lector.wm.views.region.VR_DETAIL_NAME")
                                            ));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] show detail region");
                                }
                            }

                            else if (vs != null && vs.Name.Equals(BS_FULLSCREEN_NAME))
                            {
                                // exit fullscreen mode and show detail region
                                toggleFullscreen();
                                BrailleIOScreen __nvs = GetVisibleScreen();
                                if (__nvs != null)
                                {
                                    var _detail = __nvs.GetViewRange(VR_DETAIL_NAME);
                                    if (_detail != null && !detail.IsVisible())
                                    {
                                        changeViewVisibility(__nvs.Name, VR_DETAIL_NAME, true);
                                    }
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        // Vollbildmodus
                        case "toggleFullScreen":
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "toggle full screen");
                            toggleFullscreen();
                            BrailleIOScreen _nvs = GetVisibleScreen();
                            if (_nvs != null && _nvs.Name.Equals(BS_FULLSCREEN_NAME))
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_start"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] start full screen");
                            }
                            else if (_nvs != null)
                            {
                                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        // Zoomstufe abfragen
                        case "returnZoomLevel":
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[CONTROL]\t" + "request current zoom level");
                            if (vs != null)
                            {
                                BrailleIOViewRange vr = vs.GetViewRange(VR_CENTER_NAME);
                                if (vr != null)
                                {
                                    double zl = GetZoomPercentageBasedOnPrintZoom(vs.Name, vr.Name);
                                    string message = LL.GetTrans("tangram.lector.wm.zooming.current", zl.ToString());

                                    audioRenderer.PlaySoundImmediately(message);
                                    ShowTemporaryMessageInDetailRegion(message);
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] ask for current zoom level");
                                }
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            return;

                        case "toggleGrid":
                            if(gridHook != null)
                            {
                                if(gridHook.Mode == BrailleIO.View.GridMode.Grid)
                                {
                                    gridHook.Mode = BrailleIO.View.GridMode.None;
                                    gridHook.Active = false;
                                }
                                else
                                {
                                    gridHook.Mode = BrailleIO.View.GridMode.Grid;
                                    gridHook.Active = true;
                                }
                                io.RenderDisplay();
                            }
                            e.Cancel = true;
                            e.Handled = true;
                            break;

                        default:
                            break;
                    }
                }
                #endregion

                #region Braille input mode commands
                // button commands that only work in Braille writing mode
                if (InteractionManager.Mode.Equals(InteractionMode.Braille))
                {
                    switch (e.Function)
                    {
                        case "maximizeEditArea": // maximize detail area
                                                 //toggleFullDetailScreen();
                            audioRenderer.PlayWaveImmediately(StandardSounds.Error);
                            e.Cancel = true;
                            e.Handled = true;
                            return;
                        default:
                            break;
                    }
                }
                #endregion
            }
        }


        private const String OO_DOC_WND_CLASS_NAME = "SALFRAME";
        protected override void im_GesturePerformed(object sender, GestureEventArgs e)
        {
            if (e != null)
            {
                if (e.Gesture != null && e.Gesture.Name.Equals("tap"))
                {

                    Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE]\t[INTERACTION]\t[GESTURE]\t" + "tap");

                    BrailleIOScreen vs = GetVisibleScreen();
                    if (vs != null)
                    {
                        BrailleIOViewRange vr = GetTouchedViewRange(e.Gesture.NodeParameters[0].X, e.Gesture.NodeParameters[0].Y, vs);
                        if (vr != null)
                        {

                            checkForTouchedBrailleText(vr, e);

                            String vrName = vr.Name;
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[INTERACTION] tap in " + vrName);

                            // handling for top and detail region
                            if (vrName.Equals(VR_TOP_NAME) || vrName.Equals(VR_DETAIL_NAME))
                            {
                                String text = GetTextContentOfRegion(vs, vrName);
                                audioRenderer.PlaySoundImmediately(text);
                                return;
                            }

                            // handling for status region
                            else if (vrName.Equals(VR_STATUS_NAME))
                            {
                                String text = GetTextContentOfRegion(vs, vrName);
                                switch (text)
                                {
                                    case LectorStateNames.BRAILLE_MODE:
                                        text = LL.GetTrans("tangram.lector.wm.modes.BRAILLE_MODE");
                                        break;

                                    case LectorStateNames.FOLLOW_BRAILLE_FOCUS_MODE:
                                        text = LL.GetTrans("tangram.lector.wm.modes.FOLLOW_BRAILLE_FOCUS_MODE");
                                        break;

                                    case LectorStateNames.FOLLOW_MOUSE_FOCUS_MODE:
                                        text = LL.GetTrans("tangram.lector.wm.modes.FOLLOW_MOUSE_FOCUS_MODE");
                                        break;

                                    case LectorStateNames.STANDARD_MODE:
                                        text = LL.GetTrans("tangram.lector.wm.modes.STANDARD_MODE");
                                        break;

                                    default:
                                        break;
                                }
                                audioRenderer.PlaySoundImmediately(text);
                                return;
                            }

                            Point p = GetTapPositionOnScreen(e.Gesture.NodeParameters[0].X, e.Gesture.NodeParameters[0].Y, vr);

                            //check if a OpenOffice Window is presented
                            if (DrawAppModel.ScreenObserver != null && DrawAppModel.ScreenObserver.Whnd != null)
                            {
                                var observed = isObservedOpebnOfficeDrawWindow(DrawAppModel.ScreenObserver.Whnd.ToInt32());
                                if (observed != null)
                                {
                                    handleSelectedOoAccItem(observed, p, e);
                                    return;
                                }
                            }

                            // if no OpenOfficeDraw Window could be found
                            var element = UiaPicker.GetElementFromScreenPosition(p.X, p.Y);

                            if (element != null)
                            {
                                string className = element.Current.ClassName;
                                if (String.IsNullOrEmpty(className) || className.Equals(OO_DOC_WND_CLASS_NAME))
                                {
                                    if (handleTouchRequestForOO(element, p, e)) return;
                                }


                                // TODO: ask for function?

                                //generic handling
                                if (e.ReleasedButtonsToString().Equals("Gesture") && !e.AreButtonsPressed())
                                {
                                    UiaPicker.SpeakElement(element);
                                }
                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region Decision if object is OpenOffice

        /// <summary>
        /// Handles the touch request for an OpenOoffice draw object.
        /// </summary>
        /// <param name="element">The element at a specific location.</param>
        /// <param name="p">The point, that was touched.</param>
        /// <param name="e">The <see cref="tud.mci.tangram.TangramLector.GestureEventArgs"/> instance containing the event data.</param>
        /// <returns>it is handleable by openOffice observers or should be handles by UIA picker</returns>
        private bool handleTouchRequestForOO(AutomationElement element, Point p, GestureEventArgs e)
        {
            if (element != null)
            {
                int wHndl = element.Current.NativeWindowHandle;
                int pid = element.Current.ProcessId;
                int sofficePid = OoUtils.GetOoProcessID();
                int sofficeBinPid = OoUtils.GetOoBinProcessID();

                // check if the element seems to be an OpenOffice element
                if (pid != sofficePid && pid != sofficeBinPid)
                {
                    // if not, than check if the parent process is OpenOffice
                    var pp = getParentProcess(pid);
                    if (pp == null || pp.Id != sofficePid)
                    {
                        return false;
                    }
                    else // use the wHndl of the parent process
                    {
                        wHndl = (pp != null) ? pp.MainWindowHandle.ToInt32() : wHndl;
                    }
                }


                // check the window handle
                if (wHndl == 0)
                {
                    // get the main window of the process
                    Process process = Process.GetProcessById(pid);
                    if (process != null)
                    {
                        wHndl = process.MainWindowHandle.ToInt32();
                    }
                }

                // get the 
                if (OoConnector.Instance != null && OoConnector.Instance.Observer != null)
                {
                    var observed = OoConnector.Instance.Observer.ObservesWHndl(wHndl);
                    if (observed != null)
                    {
                        handleSelectedOoAccItem(observed, p, e);
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] touched element is in a known Open Office window");
                        return true;
                    }
                }
            }
            return false;
        }

        tud.mci.tangram.Accessibility.OoAccessibleDocWnd isObservedOpebnOfficeDrawWindow(int wHndl)
        {
            if (OoConnector.Instance == null) return null;
            if (OoConnector.Instance.Observer == null) return null;
            return OoConnector.Instance.Observer.ObservesWHndl(wHndl);
        }

        /// <summary>
        /// Gets the parent process of a given pid.
        /// </summary>
        /// <param name="pid">The id of the child process.</param>
        /// <returns>the parent process</returns>
        private static Process getParentProcess(int pid)
        {
            try
            {
                var query = string.Format("SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {0}", pid);
                var search = new System.Management.ManagementObjectSearcher("root\\CIMV2", query);
                var results = search.Get().GetEnumerator();
                if (!results.MoveNext()) throw new Exception("Huh?");
                var queryObj = results.Current;
                if (queryObj != null)
                {
                    uint parentId = (uint)queryObj["ParentProcessId"];
                    var parent = Process.GetProcessById((int)parentId);
                    return parent;
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "WindowManager_Eventhandling", "Can't get a parent process: " + ex);
            }

            return null;
        }

        #endregion

    }
}