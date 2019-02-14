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
        { }


        protected override void im_FunctionCall(object sender, FunctionCallInteractionEventArgs e)
        {
            if (Active && e != null && !String.IsNullOrEmpty(e.Function))
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