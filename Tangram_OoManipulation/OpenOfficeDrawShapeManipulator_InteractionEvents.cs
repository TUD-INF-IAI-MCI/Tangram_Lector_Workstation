using System;
using System.Collections.Generic;
using System.Linq;
using TangramLector.OO;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    /// <summary>
    /// Class for manipulating OpenOffice Draw document elements 
    /// </summary>
    public partial class OpenOfficeDrawShapeManipulator : AbstractSpecializedFunctionProxyBase
    {
        #region Interaction Events

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            base.im_ButtonCombinationReleased(sender, e);
            if (Active)
            {
                e.Cancel = handleButtonCombination(sender, e);
            }
        }

        private bool handleButtonCombination(Object sender, ButtonReleasedEventArgs e)
        {

            #region InteractionMode.Normal

            if (InteractionManager.Instance.Mode == InteractionMode.Normal)
            {
                switch (e.ReleasedGenericKeys.Count)
                {
                    #region 1 Key
                    case 1:
                        switch (e.ReleasedGenericKeys[0])
                        {
                            case "k4":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] set Braille focus to mouse focus");
                                chooseElementWithGuiFocus();
                                e.Cancel = true;
                                break;
                            case "k5":
                                // TODO: Taste ist noch frei
                                break;
                            case "cru":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle up button");
                                e.Cancel = handleUP();
                                break;
                            case "crr":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle right button");
                                e.Cancel = handleRIGHT();
                                break;
                            case "crd":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle down button");
                                e.Cancel = handleDOWN();
                                break;
                            case "crl":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle left button");
                                e.Cancel = handleLEFT();
                                break;
                            case "crc":
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] rotate element manipulation dialog");
                                rotateThroughModes();
                                e.Cancel = true;
                                //TODO: open element manipulation dialog
                                break;
                            case "hbr":

                                // start gesture recognition for tabs
                                break;
                            default:
                                break;
                        }
                        break;

                    #endregion

                    #region 2 Keys

                    case 2:

                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "k4", "k5" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] choose next element");
                            chooseNextElement();
                            e.Cancel = true;
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] choose previous element");
                            choosePreviousElement();
                            e.Cancel = true;
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "k5", "k6" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] speak description");
                            string desc = OoElementSpeaker.GetElementDescriptionText(LastSelectedShape);
                            if (!String.IsNullOrEmpty(desc))
                            {
                                sentTextNotification(desc);
                                sentAudioFeedback(desc);
                            }
                            e.Cancel = true;
                        }

                        #region crc with direction button

                        else if (e.ReleasedGenericKeys.Contains("crc"))
                        {
                            if (e.ReleasedGenericKeys.Contains("cru")
                                || e.ReleasedGenericKeys.Contains("crr")
                                || e.ReleasedGenericKeys.Contains("crd")
                                || e.ReleasedGenericKeys.Contains("crl")
                                )
                            {
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] rotate element manipulation dialog (weak handling)");
                                rotateThroughModes();
                                e.Cancel = true;
                            }
                        }

                        #endregion

                        #region Diagonal Interactions

                        else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "cru", "crr" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle top right");
                            e.Cancel = handleUP_RIGHT();                            
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "cru", "crl" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle top left");
                            e.Cancel = handleUP_LEFT();
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "crd", "crr" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle down right");
                            e.Cancel = handleDOWN_RIGHT();
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String>(2) { "crd", "crl" }).ToList().Count == 2)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] handle down left");
                            e.Cancel = handleDOWN_LEFT();
                        }

                        #endregion

                        break;

                    #endregion

                    #region 3 Keys

                    case 3:
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2", "k3" }).ToList().Count == 3)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] choose parent group");
                            chooseParentOfElement();
                            e.Cancel = true;
                        }
                        else if (e.ReleasedGenericKeys.Intersect(new List<String> { "k4", "k5", "k6" }).ToList().Count == 3)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] choose first child in group");
                            chooseFirstChildOfElement();
                            e.Cancel = true;
                        }
                        break;

                    #endregion

                    #region 4 Keys

                    case 4:
                        // unfocus DOM element (no element is selected on the pin device) //
                        if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2", "k4", "k5" }).ToList().Count == 4)
                        {
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[DRAW INTERACTION] unfocus DOM element");
                            LastSelectedShape = null;
                            playEdit();
                            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.clear_braille_focus"));
                            sentTextFeedback(LL.GetTrans("tangram.oomanipulation.no_element_selected"));                          
                            e.Cancel = true;
                        }
                        break;

                    #endregion

                    default:
                        break;
                }
            }

            #endregion

            return e != null ? e.Cancel : false;
        }

        #endregion

        #region Navigation Through Elements

        /// <summary>
        /// Go to the next element in the document tree.
        /// </summary>
        private void chooseNextElement()
        {
            if (LastSelectedShape == null)
            {
                //try to get the first shape of the current page
                var activeDoc = OoConnection.GetActiveDrawDocument() as OoAccessibleDocWnd;
                if (activeDoc != null)
                {
                    OoShapeObserver first = AccDomWalker.GetFirstShapeOfDocument(activeDoc);
                    if (first != null)
                    {
                        LastSelectedShape = first;
                        sayLastSelectedShape();
                    }
                }
            }
            else
            {
                OoShapeObserver next = AccDomWalker.MoveToNext(LastSelectedShape);
                if (next != null)
                {
                    LastSelectedShape = next as OoShapeObserver;
                    sayLastSelectedShape();
                }
                else
                {
                    playError();
                }
            }
        }

        private void choosePreviousElement()
        {
            if (LastSelectedShape == null)
            {
                //try to get the last shape of the current page
                var activeDoc = OoConnection.GetActiveDrawDocument() as OoAccessibleDocWnd;
                if (activeDoc != null)
                {
                    OoShapeObserver last = AccDomWalker.GetLastShapeOfDocument(activeDoc);
                    if (last != null)
                    {
                        LastSelectedShape = last;
                        sayLastSelectedShape();
                    }
                }
            }
            else
            {
                OoShapeObserver prev = AccDomWalker.MoveToPrevious(LastSelectedShape);
                if (prev != null)
                {
                    LastSelectedShape = prev as OoShapeObserver;
                    sayLastSelectedShape();
                }
                else
                {
                    playError();
                }
            }
        }

        private void chooseFirstChildOfElement()
        {
            if (LastSelectedShape == null)
            {
                //try to get the first shape of the current page
                var activeDoc = OoConnection.GetActiveDrawDocument() as OoAccessibleDocWnd;
                if (activeDoc != null)
                {
                    OoShapeObserver first = AccDomWalker.GetFirstShapeOfDocument(activeDoc);
                    if (first != null)
                    {
                        LastSelectedShape = first;
                        sayLastSelectedShape();
                    }
                    else
                    {
                        playError();
                    }
                }
            }
            else
            {
                int index = 0;
                OoShapeObserver child = AccDomWalker.MoveToChild(LastSelectedShape, ref index);
                if (child != null)
                {
                    LastSelectedShape = child as OoShapeObserver;
                    sayLastSelectedShape();
                }
                else
                {
                    //TODO: get PolygonPoints
                    //var pPoints = LastSelectedShape.GetPolygonPoints();
                    //if (pPoints != null && pPoints.Count > 0)
                    //{
                    //    var pPoint = pPoints[0];
                    //    if (pPoint != null)
                    //    {
                    //        pPoint.MoveHorizontal(10000);
                    //    }
                    //}

                    playError();
                }
            }
        }

        private void chooseParentOfElement()
        {
            if (LastSelectedShape == null)
            {
                //try to get the first shape of the current page
                var activeDoc = OoConnection.GetActiveDrawDocument() as OoAccessibleDocWnd;
                if (activeDoc != null)
                {
                    OoShapeObserver first = AccDomWalker.GetFirstShapeOfDocument(activeDoc);
                    if (first != null)
                    {
                        LastSelectedShape = first;
                        sayLastSelectedShape();
                    }
                    else
                    {
                        playError();
                    }
                }
            }
            else
            {
                OoShapeObserver parent = AccDomWalker.MoveToParent(LastSelectedShape);
                if (parent != null)
                {
                    LastSelectedShape = parent as OoShapeObserver;
                    sayLastSelectedShape();
                }
                else
                {
                    //TODO: get PolygonPoints
                    //var pPoints = LastSelectedShape.GetPolygonPoints();
                    //if (pPoints != null && pPoints.Count > 0)
                    //{
                    //    var pPoint = pPoints[0];
                    //    if (pPoint != null)
                    //    {
                    //        pPoint.MoveHorizontal(10000);
                    //    }
                    //}
                    playError();
                }
            }
        }

        /// <summary>
        /// Sets the Braille focus to that element which currently has the Mouse (GUI) focus.
        /// </summary>
        private void chooseElementWithGuiFocus()
        {
            if (OoConnection != null)
            {
                OoShapeObserver shape = OoConnection.GetCurrentDrawSelection() as OoShapeObserver;
                if (shape != null && shape.IsValid())
                {
                    LastSelectedShape = shape;
                    //AudioRenderer.Instance.PlaySoundImmediately("Braille-Fokus auf " + LastSelectedShape.Name + " gesetzt.");
                    AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.set_braille_focus", LastSelectedShape.Name));
                    return;
                }
            }
            playError();
        }

        #endregion

    }
}
