using System;
using System.Collections.Generic;
using System.Linq;
using BrailleIO;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager
    {
        /// <summary>
        /// Change interaction possibilities for Lange Nacht der Wissenschaften
        /// </summary>
        public void InteractionForLNDW()
        {
            // fullscreen
            if (io.GetView(BS_FULLSCREEN_NAME) != null)
            {
                var view = io.GetView(BS_FULLSCREEN_NAME) as BrailleIO.Interface.IViewable;
                if (view != null && !view.IsVisible())
                {
                    toggleFullscreen();
                }
            }

            // deactivate button commands
            InteractionManager.ButtonCombinationReleased -= new EventHandler<ButtonReleasedEventArgs>(im_ButtonCombinationReleased);
            InteractionManager.ButtonCombinationReleased += new EventHandler<ButtonReleasedEventArgs>(interactionManager_ButtonCombinationReleasedLNDW);

        }

        void interactionManager_ButtonCombinationReleasedLNDW(object sender, ButtonReleasedEventArgs e)
        {
            if (e != null && e.ReleasedGenericKeys != null)
            {
                BrailleIOScreen vs = GetVisibleScreen();

                if (e.PressedGenericKeys == null || e.PressedGenericKeys.Count < 1)
                {
                    BrailleIOViewRange center = GetActiveCenterView(vs);

                    if (e.ReleasedGenericKeys.Count == 1)
                    {
                        switch (e.ReleasedGenericKeys[0])
                        {
                            case "nsll":// links blättern
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, center.ContentBox.Width))
                                        audioRenderer.PlaySoundImmediately("links blättern");
                                return;
                            case "nsl":// links verschieben
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, 5))
                                        audioRenderer.PlaySoundImmediately("links verschieben");
                                return;
                            case "nsr":// rechts verschieben
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, -5))
                                        audioRenderer.PlaySoundImmediately("rechts verschieben");
                                return;
                            case "nsrr":// rechts blättern
                                if (vs != null && center != null)
                                    if (moveHorizontal(vs.Name, center.Name, -center.ContentBox.Width))
                                        audioRenderer.PlaySoundImmediately("rechts blättern");
                                return;
                            case "nsuu":// hoch blättern
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, center.ContentBox.Height))
                                        audioRenderer.PlaySoundImmediately("hoch blättern");
                                return;
                            case "nsu":// hoch verschieben
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, 5))
                                        audioRenderer.PlaySoundImmediately("hoch verschieben");
                                return;
                            case "nsd":// runter verschieben
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, -5))
                                        audioRenderer.PlaySoundImmediately("runter verschieben");
                                return;
                            case "nsdd":// runter blättern
                                if (vs != null && center != null)
                                    if (moveVertical(vs.Name, center.Name, -center.ContentBox.Height))
                                        audioRenderer.PlaySoundImmediately("runter blättern");
                                return;
                            case "rslu": // kleiner Zoom in
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, 1.5))
                                        audioRenderer.PlaySoundImmediately("Ausschnitt vergrößert");
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                return;
                            case "rsld": // kleiner Zoom out
                                if (vs != null && center != null)
                                {
                                    if (zoomWithFactor(vs.Name, center.Name, 1 / 1.5))
                                        audioRenderer.PlaySoundImmediately("Ausschnitt verkleinert");
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                return;
                            case "l": // Breite/Höhe einpassen
                                if (vs != null && !vs.Name.Equals(BS_MINIMAP_NAME))
                                {
                                    if (ZoomTo(vs.Name, VR_CENTER_NAME, -1))
                                        audioRenderer.PlaySoundImmediately("Ausschnitt an Höhe angepasst");
                                    else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                }
                                return;
                            case "lr": // Invertieren
                                if (vs != null)
                                {
                                    invertImage(vs.Name, VR_CENTER_NAME);
                                    audioRenderer.PlaySoundImmediately("Ausgabe invertiert");
                                }
                                return;
                            case "rl": // Schwellwert minus
                                if (vs != null)
                                {
                                    updateContrast(vs.Name, VR_CENTER_NAME, -THRESHOLD_STEP);
                                    audioRenderer.PlaySoundImmediately("Schwellwert verringert");
                                }
                                return;
                            case "r": // Schwellwert plus
                                if (vs != null)
                                {
                                    updateContrast(vs.Name, VR_CENTER_NAME, THRESHOLD_STEP);
                                    audioRenderer.PlaySoundImmediately("Schwellwert erhöht");
                                }
                                return;
                            default:
                                break;
                        }
                    }
                    else if (e.ReleasedGenericKeys.Count == 2)
                    {
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsdd", "nsd" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveVertical(vs.Name, center.Name, -center.ContentBox.Height))
                                    audioRenderer.PlaySoundImmediately("runter blättern");
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsuu", "nsu" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveVertical(vs.Name, center.Name, center.ContentBox.Height))
                                    audioRenderer.PlaySoundImmediately("hoch blättern");
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsll", "nsl" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveHorizontal(vs.Name, center.Name, center.ContentBox.Width))
                                    audioRenderer.PlaySoundImmediately("links blättern");
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsrr", "nsr" }).ToList().Count == 2)
                        {
                            if (vs != null && center != null)
                                if (moveHorizontal(vs.Name, center.Name, -center.ContentBox.Width))
                                    audioRenderer.PlaySoundImmediately("rechts blättern");
                        }
                    }
                }
            }
        }

    }
}
