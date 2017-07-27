using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Automation;
using BrailleIO;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.util;
using tud.mci.tangram.audio;
using tud.mci.tangram.Uia;
using tud.mci.tangram.controller.observer;

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
                registerAsSpecializedFunctionProxy();
                InteractionManager.InteractionModeChanged += new EventHandler<InteractionModeChangedEventArgs>(InteractionManager_InteractionModeChanged);
            }
            registerTopMostSpFProxy();
        }

        void AdapterManager_NewAdapterRegistered(object sender, BrailleIO.Interface.IBrailleIOAdapterEventArgs e)
        {
            if (e != null && e.Adapter != null && e.Adapter.Device != null)
            {
                InteractionManager.AddNewDevice(e.Adapter);
                // TODO: monitor new adapter!
            }
        }

        void AdapterManager_ActiveAdapterChanged(object sender, BrailleIO.Interface.IBrailleIOAdapterEventArgs e)
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
        {
            if (e != null && e.ReleasedGenericKeys != null)
            {
                BrailleIOScreen vs = GetVisibleScreen();

                #region General Commands

                switch (e.ReleasedGenericKeys.Count)
                {
                    case 1:
                        switch (e.ReleasedGenericKeys[0])
                        {
                            case "l": // abort speech output
                                AbortAudio();
                                e.Cancel = true;
                                return;

                            default:
                                break;
                        }
                        break;

                    case 3:
                        if (e.ReleasedGenericKeys.Contains("k1"))
                        {
                            // calibrate BrailleDis //
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k3", "hbr" }).ToList().Count == 3)
                            {
                                if (io != null)
                                {
                                    this.ScreenObserver.Stop();
                                    audioRenderer.PlaySound(LL.GetTrans("tangram.lector.wm.recalibrate"));
                                    // bool res = io.Recalibrate();
                                    bool res = io.RecalibrateAll();
                                    this.ScreenObserver.Start();
                                    audioRenderer.PlaySound(LL.GetTrans("tangram.lector.wm.recalibrate.success", (res ? "" : LL.GetTrans("tangram.lector.not") + " ")));
                                }
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] recalibrate the device");
                                e.Cancel = true;
                                return;
                            }
                        }
                        break;

                    case 4:
                        // follow GUI focus mode //
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2", "k4", "k7" }).ToList().Count == 4)
                        {
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
                                    if(OoDrawAccessibilityObserver.Instance != null){
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
                            e.Cancel = true;
                            return;
                        }
                        break;
                    default:
                        break;
                }

                #endregion

                #region minimap commands
                if (vs != null && vs.Name.Equals(BS_MINIMAP_NAME))
                {
                    if (e.PressedGenericKeys == null || e.PressedGenericKeys.Count < 1)
                    {
                        if (e.ReleasedGenericKeys.Count == 1)
                        {
                            if (e.ReleasedGenericKeys[0].Equals("k2")) { } // exit minimap 
                            else if (e.ReleasedGenericKeys[0].Equals("lr")) { } // invert image
                            else
                            {
                                vs = screenBeforeMinimap; // operation has to be executed for normal screen, minimap is automatically synchronized to this
                                switch (e.ReleasedGenericKeys[0])
                                {
                                    // panning operations will be handled later
                                    case "nsll": break;
                                    case "nsl": break;
                                    case "nsr": break;
                                    case "nsrr": break;
                                    case "nsuu": break;
                                    case "nsu": break;
                                    case "nsd": break;
                                    case "nsdd": break;
                                    case "clu": break;
                                    case "cll": break;
                                    case "clr": break;
                                    case "cld": break;

                                    default: // all other operations lead to error sound
                                        audioRenderer.PlayWaveImmediately(StandardSounds.Error);
                                        e.Cancel = true;
                                        return;
                                }
                            }
                        }
                    }
                }
                #endregion

                #region global button commands

                if (e.PressedGenericKeys == null || e.PressedGenericKeys.Count < 1)
                {
                    BrailleIOViewRange center = GetActiveCenterView(vs);
                    BrailleIOViewRange detail = null;
                    if (vs != null) detail = vs.GetViewRange(VR_DETAIL_NAME);

                    if (e.ReleasedGenericKeys.Count == 1)
                    {
                        switch (e.ReleasedGenericKeys[0])
                        {
                            #region panning operations
                            case "nsll":// links blättern
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, center.ContentBox.Width))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.left_big"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll left");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsl":// links verschieben
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, 5))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.left"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll left");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsr":// rechts verschieben
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, -5))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.right"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll right");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsrr":// rechts blättern
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, -center.ContentBox.Width))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.right_big"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll right");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsuu":// hoch blättern
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, center.ContentBox.Height))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.up_big"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll up");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsu":// hoch verschieben
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, 5))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.up"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll up");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsd":// runter verschieben
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, -5))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.down"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small scroll down");
                                    }
                                e.Cancel = true;
                                return;
                            case "nsdd":// runter blättern
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, -center.ContentBox.Height))
                                    {
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.down_big"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll down");
                                    }
                                e.Cancel = true;
                                return;
                            case "clu":// Sprung nach oben
                                jumpToTopBorder(center);
                                e.Cancel = true;
                                return;
                            case "cll":// Sprung nach links
                                jumpToLeftBorder(center);
                                e.Cancel = true;
                                return;
                            case "clr":// Sprung nach rechts
                                jumpToRightBorder(center);
                                e.Cancel = true;
                                return;
                            case "cld":// Sprung nach unten
                                jumpToBottomBorder(center);
                                e.Cancel = true;
                                return;

                            #endregion

                            #region zooming operations
                            case "rslu": // kleiner Zoom in
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, 1.5))
                                    {
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.zooming.in",
                                            GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name)));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small zoom in");
                                    }
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                e.Cancel = true;
                                return;
                            case "rsld": // kleiner Zoom out
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, 1 / 1.5))
                                    {
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.zooming.out",
                                            GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name)));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] small zoom out");
                                    }
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                e.Cancel = true;
                                return;
                            case "rsru": // großer Zoom in
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, 3))
                                    {
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.zooming.in_big",
                                            GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name)));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big zoom in");
                                    }
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                e.Cancel = true;
                                return;
                            case "rsrd": // großer Zoom out
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, (double)1 / 3))
                                    {
                                        audioRenderer.PlaySoundImmediately(
                                            LL.GetTrans("tangram.lector.wm.zooming.out_big",
                                            GetZoomPercentageBasedOnPrintZoom(vs.Name, center.Name)));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big zoom out");
                                    }
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                e.Cancel = true;
                                return;
                            #endregion

                            #region detail area scrolling

                            case "cru": // scroll detail region up
                                if (detail != null)
                                {
                                    if (scrollViewRange(detail, 5))
                                    {
                                        e.Cancel = true;
                                        break;
                                    }
                                }
                                AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                                break;

                            case "crl": // no function yet
                                AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                                break;

                            case "crr": // no function yet
                                AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                                break;

                            case "crd": // scroll detail region down
                                if (detail != null)
                                {
                                    if (scrollViewRange(detail, -5))
                                    {
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
                    }

                    else if (e.ReleasedGenericKeys.Count == 2) // TODO: maybe bring this in one
                    {
                        #region panning operations
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsdd", "nsd" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveVertical(vs.Name, center.Name, -center.ContentBox.Height))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.down_big"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll down");
                                }
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsuu", "nsu" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveVertical(vs.Name, center.Name, center.ContentBox.Height))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.up_big"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll up");
                                }
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsll", "nsl" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveHorizontal(vs.Name, center.Name, center.ContentBox.Width))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.left_big"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll left");
                                }
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsrr", "nsr" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveHorizontal(vs.Name, center.Name, -center.ContentBox.Width))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.right_big"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] big scroll right");
                                }
                        }
                        #endregion
                    }

                    else if (e.ReleasedGenericKeys.Count == 8)
                    {
                        // Ansichtswechsel
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2", "k3", "k4", "k5", "k6", "k7", "k8" }).ToList().Count == 8)
                        {
                            string message = "";
                            if (currentView.Equals(LectorView.Braille))
                            {
                                if (ChangeLectorView(LectorView.Drawing))
                                {
                                    changeActiveCenterView(vs, VR_CENTER_NAME);
                                    message = LL.GetTrans("tangram.lector.wm.views.draw_activated");
                                }
                            }
                            else if (currentView.Equals(LectorView.Drawing))
                            {
                                if (ChangeLectorView(LectorView.Braille))
                                {
                                    changeActiveCenterView(vs, VR_CENTER_2_NAME);
                                    message = LL.GetTrans("tangram.lector.wm.views.braille_activated");
                                }
                            }
                            audioRenderer.PlaySoundImmediately(message);
                            ShowTemporaryMessageInDetailRegion(message);
                            fillMainCenterContent();
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] " + message);
                            e.Cancel = true;
                            return;
                        }
                    }
                }
                #endregion

                #region Drawing mode button commands

                // button commands that are not working in Braille mode
                if (!InteractionManager.Mode.Equals(InteractionMode.Braille))
                {
                    if (e.PressedGenericKeys == null || e.PressedGenericKeys.Count < 1)
                    {
                        #region 1 key

                        if (e.ReleasedGenericKeys.Count == 1)
                        {
                            switch (e.ReleasedGenericKeys[0])
                            {
                                case "k1": // Druck-Zoom

                                    if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                                    {
                                        if (ZoomTo(vs.Name, VR_CENTER_NAME, GetPrintZoomLevel()))
                                        {
                                            audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.to_print"));
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] print zoom");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                    e.Cancel = true;
                                    return;

                                case "k2": // Minimap
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
                                    return;

                                case "k3": // Breite/Höhe einpassen
                                    if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                                    {
                                        if (ZoomTo(vs.Name, VR_CENTER_NAME, -1))
                                        {
                                            audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.fit_height"));
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] zoom fit to height");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                    e.Cancel = true;
                                    return;

                                case "k7": // 1-zu-1 Zoom
                                    if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                                    {
                                        if (ZoomTo(vs.Name, VR_CENTER_NAME, 1))
                                        {
                                            audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.zooming.1-1"));
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] zoom 1:1");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                    e.Cancel = true;
                                    return;

                                case "lr": // Invertieren
                                    if (vs != null)
                                    {
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
                                    return;

                                case "rl": // Schwellwert minus
                                    if (vs != null)
                                    {
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
                                    return;

                                case "r": // Schwellwert plus
                                    if (vs != null)
                                    {
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
                                    return;

                                default:
                                    break;
                            }
                        }
                        #endregion

                        #region 2 keys

                        else if (e.ReleasedGenericKeys.Count == 2)
                        {
                            // Standard-Schwellwert
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "rl", "r" }).ToList().Count == 2)
                            {
                                if (vs != null)
                                {
                                    setContrast(vs.Name, VR_CENTER_NAME, STANDARD_CONTRAST_THRESHOLD);
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.threshold_reset"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] reset threshold");
                                }
                            }
                        }

                        #endregion 

                        #region 3 keys

                        else if (e.ReleasedGenericKeys.Count == 3)
                        {
                            // Kopfbereich ein/aus
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k3", "k8" }).ToList().Count == 3)
                            {
                                if (vs != null && vs.GetViewRange(VR_TOP_NAME) != null)
                                {
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
                                    BrailleIOScreen nvs = GetVisibleScreen();
                                    if (nvs != null)
                                    {
                                        BrailleIOViewRange top = nvs.GetViewRange(VR_TOP_NAME);
                                        if (top != null && !top.IsVisible())
                                        {
                                            changeViewVisibility(nvs.Name, VR_TOP_NAME, true);
                                        }
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                                    }
                                }
                                e.Cancel = true;
                                return;
                            }
                        }

                        #endregion

                        #region 4 keys
                        else if (e.ReleasedGenericKeys.Count == 4)
                        {
                            // Detailbereich ein/aus
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k4", "k5", "k8" }).ToList().Count == 4)
                            {
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
                                    BrailleIOScreen nvs = GetVisibleScreen();
                                    if (nvs != null)
                                    {
                                        BrailleIOViewRange detail = nvs.GetViewRange(VR_DETAIL_NAME);
                                        if (detail != null && !detail.IsVisible())
                                        {
                                            changeViewVisibility(nvs.Name, VR_DETAIL_NAME, true);
                                        }
                                        audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                                    }
                                }
                                e.Cancel = true;
                                return;
                            }
                        }

                        #endregion 

                        #region 5 keys

                        else if (e.ReleasedGenericKeys.Count == 5)
                        {
                            // Vollbildmodus
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k3", "k4", "k6", "k8" }).ToList().Count == 5)
                            {
                                toggleFullscreen();
                                BrailleIOScreen nvs = GetVisibleScreen();
                                if (nvs != null && nvs.Name.Equals(BS_FULLSCREEN_NAME))
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_start"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] start full screen");
                                }
                                else if (nvs != null)
                                {
                                    audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.full_screen_exit"));
                                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] exit full screen");
                                }
                                e.Cancel = true;
                                return;
                            }
                        }
                        #endregion 

                        #region 6 keys

                        else if (e.ReleasedGenericKeys.Count == 6)
                        {
                            // Zoomstufe abfragen
                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k3", "k5", "k6", "k7", "k8" }).ToList().Count == 6)
                            {
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
                                return;
                            }
                        }
                        #endregion 
                    }
                }
                #endregion

                #region Braille input mode commands
                // button commands that only work in Braille writing mode
                if (InteractionManager.Mode.Equals(InteractionMode.Braille))
                {
                    if (e.PressedGenericKeys == null || e.PressedGenericKeys.Count < 1)
                    {
                        if (e.ReleasedGenericKeys.Count == 1)
                        {
                            switch (e.ReleasedGenericKeys[0])
                            {
                                case "r": // maximize detail area
                                    //toggleFullDetailScreen();
                                    audioRenderer.PlayWaveImmediately(StandardSounds.Error);
                                    e.Cancel = true;
                                    return;
                            }
                        }
                    }
                }
                #endregion
            }
        }

        private const String OO_DOC_WND_CLASS_NAME = "SALFRAME";
        protected override void im_GesturePerformed(object sender, GestureEventArgs e)
        {
            if (e != null )
            {
                if (e.Gesture != null && e.Gesture.Name.Equals("tap"))
                {
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
                            if (ScreenObserver != null && ScreenObserver.Whnd != null)
                            {
                                var observed = isObservedOpebnOfficeDrawWindow(ScreenObserver.Whnd.ToInt32());
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

                                //generic handling
                                if (e.ReleasedGenericKeys.Count < 2 &&
                                    e.PressedGenericKeys.Count < 1 &&
                                    e.ReleasedGeneralKeys.Contains(BrailleIO.Interface.BrailleIO_DeviceButton.Gesture))
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