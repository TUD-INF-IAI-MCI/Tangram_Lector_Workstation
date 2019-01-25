using System;
using tud.mci.tangram.audio;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;

namespace tud.mci.tangram.TangramLector.OO
{
    public partial class OoObserver : AbstractSpecializedFunctionProxyBase
    {

        void registerAsSpecializedFunctionProxy()
        {
            if (ScriptFunctionProxy.Instance != null)
            {
                ScriptFunctionProxy.Instance.AddProxy(this);
            }
            this.Active = true;
        }

        #region events

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            //if (Active)
            //{
            //    #region global button commands

            //    if ((e.ReleasedGenericKeys.Count == 1 && e.ReleasedGenericKeys[0] == "clc") || // center the focused element
            //        // handling if the cursor-key-pad is wrongly pressed with another key of this pad
            //        (e.ReleasedGenericKeys.Count == 2 && e.ReleasedGenericKeys.Contains("clc") &&
            //            (e.ReleasedGenericKeys.Contains("clu") || e.ReleasedGenericKeys.Contains("cll") || e.ReleasedGenericKeys.Contains("clr") || e.ReleasedGenericKeys.Contains("cld"))))
            //    {
            //        if (jumpToDomFocus()) AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.center_element"));
            //        else AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.End);
            //        Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] center focused shape");
            //        e.Cancel = true;
            //    }

            //    else if (e.ReleasedGenericKeys.Count == 2 && e.ReleasedGenericKeys.Contains("l") && e.ReleasedGenericKeys.Contains("lr"))
            //    {
            //        if (!shapeManipulatorFunctionProxy.IsShapeSelected/* .LastSelectedShape == null*/)
            //        {
            //            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.no_element"));
            //            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] mark Braille focus in GUI - no element focused");
            //        }
            //        else
            //        {
            //            InitBrailleDomFocusHighlightMode();   // retriggers focus highlighting
            //            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.mark_in_gui"));
            //            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] mark Braille focus in GUI");
            //        }
            //        e.Cancel = true;
            //    }
            //    #endregion

            //    #region InteractionMode.Normal

            //    if (InteractionManager.Instance.Mode == InteractionMode.Normal)
            //    {
            //        switch (e.ReleasedGenericKeys.Count)
            //        {

            //            #region 1 Key

            //            //TODO: stop blinking only if a manipulation mode is active
            //            //TODO: stop blinking for diagonal interaction
            //            case 1:

            //                switch (e.ReleasedGenericKeys[0])
            //                {
            //                    // DE/ACTIVATE BRAILLE TEXT ADDITION
            //                    case "k6":
            //                        string text = LL.GetTrans("tangram.lector.oo_observer.text.braille_replacement") + " ";
            //                        if (this.TextRendererHook.CanBeActivated())
            //                        {
            //                            this.TextRendererHook.Active = !this.TextRendererHook.Active;
            //                            if (this.TextRendererHook.Active) text += LL.GetTrans("tangram.lector.oo_observer.activated");
            //                            else text += LL.GetTrans("tangram.lector.oo_observer.deactivated");
            //                            AudioRenderer.Instance.PlaySoundImmediately(text);
            //                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] show Braille font in drawing: " + this.TextRendererHook.Active.ToString());
            //                        }
            //                        else
            //                        {
            //                            AudioRenderer.Instance.PlaySoundImmediately(text + LL.GetTrans("tangram.lector.oo_observer.text.braille_replacement.exception"));
            //                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] can't show Braille font in drawing");
            //                        }
            //                        break;
            //                    // TOGGLE FOCUS BLINKING
            //                    case "k8":
            //                        if (blinkFocusActive)
            //                        {
            //                            StopFocusHighlightModes();
            //                            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.blinking") + " " + LL.GetTrans("tangram.lector.oo_observer.deactivated"));
            //                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] blinking Braille focus frame off");
            //                        }
            //                        else
            //                        {
            //                            StartFocusHighlightModes();
            //                            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.blinking") + " " + LL.GetTrans("tangram.lector.oo_observer.activated"));
            //                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] blinking Braille focus frame on");
            //                        }
            //                        e.Cancel = true;
            //                        break;
            //                    case "cru":
            //                        if(shapeManipulatorFunctionProxy.Mode != ModificationMode.Unknown) PauseFocusHighlightModes();
            //                        break;
            //                    case "crr":
            //                        if (shapeManipulatorFunctionProxy.Mode != ModificationMode.Unknown) PauseFocusHighlightModes();
            //                        break;
            //                    case "crd":
            //                        if (shapeManipulatorFunctionProxy.Mode != ModificationMode.Unknown) PauseFocusHighlightModes();
            //                        break;
            //                    case "crl":
            //                        if (shapeManipulatorFunctionProxy.Mode != ModificationMode.Unknown) PauseFocusHighlightModes();
            //                        break;
            //                    case "crc":
            //                        // PauseFocusHighlightModes();
            //                        break;
            //                    default:
            //                        break;
            //                }
            //                break;

            //            #endregion

            //            #region 2 Keys

            //            case 2:

            //                #region crc with direction button

            //                if (e.ReleasedGenericKeys.Contains("crc"))
            //                {
            //                    PauseFocusHighlightModes();
            //                }

            //                #endregion

            //                #region Diagonal Interactions

            //                else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "cru", "crr" }).ToList().Count == 2)
            //                {
            //                    PauseFocusHighlightModes();
            //                }
            //                else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "cru", "crl" }).ToList().Count == 2)
            //                {
            //                    PauseFocusHighlightModes();
            //                }
            //                else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "crd", "crr" }).ToList().Count == 2)
            //                {
            //                    PauseFocusHighlightModes();
            //                }
            //                else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "crd", "crl" }).ToList().Count == 2)
            //                {
            //                    PauseFocusHighlightModes();
            //                }

            //                #endregion

            //                break;

            //            #endregion

            //            #region 3 Keys

            //            case 3:
            //                // follow DOM focus mode //
            //                if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2", "k4" }).ToList().Count == 3)
            //                {
            //                    Logger.Instance.Log(LogPriority.MIDDLE,this,"[DRAW INTERACTION] jump to DOM focus");
            //                    WindowManager.Instance.FocusMode = WindowManager.Instance.FocusMode != FollowFocusModes.FOLLOW_BRAILLE_FOCUS ? FollowFocusModes.FOLLOW_BRAILLE_FOCUS : FollowFocusModes.NONE;
            //                    if (WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_BRAILLE_FOCUS)
            //                    {
            //                        bool success = jumpToDomFocus();
            //                        if (!success)
            //                        {
            //                            audioRenderer.PlaySoundImmediately(
            //                                LL.GetTrans("tangram.lector.oo_observer.focus.folowing_braille")
            //                                + " " + LL.GetTrans("tangram.lector.oo_observer.activated")
            //                                + ". " + LL.GetTrans("tangram.lector.oo_observer.selected_no"));
            //                        }
            //                    }
            //                    e.Cancel = true;
            //                }
            //                break;

            //            #endregion

            //            #region 4 Keys

            //            case 4:

            //                // open title/description dialog
            //                if (e.ReleasedGenericKeys.Intersect(new List<String> { "k2", "k3", "k4", "k5" }).ToList().Count == 4)
            //                {
            //                    PauseFocusHighlightModes();
            //                    openTitleDescDialog();
            //                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] open title and description dialog");
            //                    e.Cancel = true;
            //                }
            //                break;

            //            #endregion

            //            default:
            //                break;
            //        }
            //    }
            //    #endregion
            //}
        }


        protected override void im_FunctionCall(object sender, FunctionCallInteractionEventArgs e)
        {
            if (Active && e != null && !String.IsNullOrEmpty(e.Function) && !e.AreButtonsPressed())
            {

                switch (e.Function)
                {
                    #region global button commands

                    case "centerFocusedElement": // center the focused element
                        if (jumpToDomFocus()) AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.center_element"));
                        else AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.End);
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] center focused shape");
                        e.Cancel = true;
                        e.Handled = true;
                        break;


                    case "highlightDOMFocusInGui":
                        if (!shapeManipulatorFunctionProxy.IsShapeSelected/* .LastSelectedShape == null*/)
                        {
                            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.no_element"));
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] mark Braille focus in GUI - no element focused");
                        }
                        else
                        {
                            InitBrailleDomFocusHighlightMode();   // retriggers focus highlighting
                            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.mark_in_gui"));
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] mark Braille focus in GUI");
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        break;

                    // DE/ACTIVATE BRAILLE TEXT ADDITION
                    case "toggleBrailleDisplayOverlay":
                        string text = LL.GetTrans("tangram.lector.oo_observer.text.braille_replacement") + " ";
                        if (this.TextRendererHook.CanBeActivated())
                        {
                            this.TextRendererHook.Active = !this.TextRendererHook.Active;
                            if (this.TextRendererHook.Active) text += LL.GetTrans("tangram.lector.oo_observer.activated");
                            else text += LL.GetTrans("tangram.lector.oo_observer.deactivated");
                            AudioRenderer.Instance.PlaySoundImmediately(text);
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] show Braille font in drawing: " + this.TextRendererHook.Active.ToString());
                        }
                        else
                        {
                            AudioRenderer.Instance.PlaySoundImmediately(text + LL.GetTrans("tangram.lector.oo_observer.text.braille_replacement.exception"));
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] can't show Braille font in drawing");
                        }
                        e.Handled = true;
                        break;

                    case "toggleDOMFocusMode":
                        // follow DOM focus mode //
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] jump to DOM focus");
                        WindowManager.Instance.FocusMode = WindowManager.Instance.FocusMode != FollowFocusModes.FOLLOW_BRAILLE_FOCUS ? FollowFocusModes.FOLLOW_BRAILLE_FOCUS : FollowFocusModes.NONE;
                        if (WindowManager.Instance.FocusMode == FollowFocusModes.FOLLOW_BRAILLE_FOCUS)
                        {
                            bool success = jumpToDomFocus();
                            if (!success)
                            {
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.oo_observer.focus.folowing_braille")
                                    + " " + LL.GetTrans("tangram.lector.oo_observer.activated")
                                    + ". " + LL.GetTrans("tangram.lector.oo_observer.selected_no"));
                            }
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        break;


                    #endregion

                    #region InteractionMode.Manipulation

                    // TOGGLE FOCUS BLINKING
                    case "toggleFocusHighlight":
                        if (InteractionManager.Instance.Mode.HasFlag(InteractionMode.Manipulation))
                        {
                            if (blinkFocusActive)
                            {
                                StopFocusHighlightModes();
                                AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.blinking") + " " + LL.GetTrans("tangram.lector.oo_observer.deactivated"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] blinking Braille focus frame off");
                            }
                            else
                            {
                                StartFocusHighlightModes();
                                AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.focus.blinking") + " " + LL.GetTrans("tangram.lector.oo_observer.activated"));
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] blinking Braille focus frame on");
                            }
                            e.Cancel = true;
                            e.Handled = true;
                        }
                        else
                        {
                            AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        }
                        break;

                    case "editTitleDesc":
                        if (InteractionManager.Instance.Mode.HasFlag(InteractionMode.Manipulation))
                        {
                            // open title/description dialog
                            PauseFocusHighlightModes();
                            openTitleDescDialog();
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] open title and description dialog");
                            e.Cancel = true;
                            e.Handled = true;
                        }
                        break;

                    case "changeUp":
                    case "changeDown":
                    case "changeLeft":
                    case "changeRight":
                    case "changeUpRight":
                    case "changeDownRight":
                    case "changeUpLeft":
                    case "changeDownLeft":
                        if (shapeManipulatorFunctionProxy.Mode != ModificationMode.Unknown)
                            PauseFocusHighlightModes();
                        break;

                    case "confirm":
                        // PauseFocusHighlightModes();
                        break;

                    #endregion

                    default:
                        break;

                }
            }
        }

        #endregion
    }
}