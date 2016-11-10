using tud.mci.tangram;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector;
using System;
using tud.mci.tangram.audio;

namespace TangramLector.OO
{
    /// <summary>
    /// Static helper class to walk through a OpenOffice tree of (wrapped) accessible objects
    /// </summary>
    static class AccDomWalker
    {

        #region First / Last

        /// <summary>
        /// Gets the first shape of document.
        /// </summary>
        /// <param name="doc">The observed document.</param>
        /// <param name="pages">The observed page.</param>
        /// <returns></returns>
        public static OoShapeObserver GetFirstShapeOfDocument(OoAccessibleDocWnd doc, OoDrawPageObserver pages = null)
        {
            if (doc != null)
            {
                if (pages == null) { pages = doc.GetActivePage(); }
                if (pages != null && doc.DocumentComponent != null && doc.DocumentComponent.ChildCount > 0)
                {
                    // a page doesn't have children in the accessible tree --> damn
                    // so we have to go through the shape structure
                    return pages.GetFirstChild();
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "AccDomWalker", "The document to walk through seems to have no child shapes");
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the last shape of document.
        /// </summary>
        /// <param name="doc">The doc.</param>
        /// <param name="pages">The pages.</param>
        /// <returns></returns>
        public static OoShapeObserver GetLastShapeOfDocument(OoAccessibleDocWnd doc, OoDrawPageObserver pages = null)
        {
            if (doc != null)
            {
                if (pages == null) { pages = doc.GetActivePage(); }
                if (pages != null && doc.DocumentComponent != null && doc.DocumentComponent.ChildCount > 0)
                {
                    // a page doesn't have children in the accessible tree --> damn
                    // so we have to go through the shape structure
                    return pages.GetLastChild();
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "AccDomWalker", "The document to walk through seems to have no child shapes");
                }
            }
            return null;
        }

        #endregion
        
        #region Public General Tree Movement Commands

        /// <summary>
        /// Moves to the next element.
        /// </summary>
        /// <param name="shape">The shape to get the next sibling.</param>
        /// <returns>
        /// The next shape on the screen if possible otherwise the next in the DOM if possible otherwise <c>null</c>
        /// </returns>
        public static OoShapeObserver MoveToNext(OoShapeObserver shape)
        {
            if (shape != null && shape.IsValid())
            {
                // move thought the Accessible tree because this elements should be visible on the screen and the tree navigation is faster
                // TODO remove false to use search in accessibility tree again. Currently accessing the page AccessibleShape element during search crashes openoffice!
                if (false && shape.AccComponent != null && shape.AccComponent.IsValid())
                {
                    OoAccComponent comp = shape.AccComponent;

                    while (true)
                    {
                        OoAccComponent next = moveToNextComponent(comp);

                        if (next != null && next != comp && next.IsValid())
                        {
                            if (acceptAsUsableShape(next))
                            {
                                OoShapeObserver obs = getObserverForAccessible(next, shape.Page);
                                if (obs != null) return obs;
                                else break;
                            }
                            else
                            {
                                comp = next;
                                continue;
                            }
                        }
                        break;
                    }
                }
                
                // go through the dom tree
                {
                    return moveToNextShape(shape);
                }
            }

            return null;
        }

        /// <summary>
        /// Moves to the previous element.
        /// </summary>
        /// <param name="shape">The shape to get the previous sibling.</param>
        /// <returns>
        /// The previous shape on the screen if possible otherwise the previous in the DOM if possible otherwise <c>null</c>
        /// </returns>
        public static OoShapeObserver MoveToPrevious(OoShapeObserver shape)
        {
            if (shape != null && shape.IsValid())
            {
                // move thought the Accessible tree because this elements should be visible on the screen and the tree navigation is faster
                // TODO remove false to use search in accessibility tree again. Currently accessing the page AccessibleShape element during search crashes openoffice!
                if (false && shape.AccComponent != null && shape.AccComponent.IsValid())
                {
                    OoAccComponent comp = shape.AccComponent;

                    while (true)
                    {
                        OoAccComponent prev = moveToPrevComponent(comp);

                        if (prev != null && prev != comp)
                        {
                            if (acceptAsUsableShape(prev))
                            {
                                var sObs = getObserverForAccessible(prev, shape.Page);
                                if (sObs != null) return sObs;
                                else
                                {
                                    break;
                                }
                            }
                            else
                            {
                                comp = prev;
                                continue;
                            }
                        }
                        break;
                    }
                }
                else // go through the dom tree
                {
                    return moveToPreviousShape(shape);
                }
            }

            return null;
        }

        /// <summary>
        /// Moves to the parent element.
        /// </summary>
        /// <param name="shape">The child shape to get the parent of.</param>
        /// <returns>
        /// The parent shape on the screen if possible otherwise the parent in the DOM if possible otherwise <c>null</c>
        /// </returns>
        public static OoShapeObserver MoveToParent(OoShapeObserver shape)
        {
            if (shape != null && shape.IsValid())
            {
                // move thought the Accessible tree because this elements should be visible on the screen and the tree navigation is faster
                // TODO remove false to use search in accessibility tree again. Currently accessing the page AccessibleShape element during search crashes openoffice!
                if (false && shape.AccComponent != null && shape.AccComponent.IsValid())
                {
                    OoAccComponent comp = shape.AccComponent;

                    while (true)
                    {
                        OoAccComponent parent = moveToParentComponent(comp);

                        if (parent != null && parent != comp)
                        {
                            if (acceptAsUsableShape(parent))
                            {
                                return getObserverForAccessible(parent, shape.Page);
                            }
                            else
                            {
                                parent = moveToParentComponent(parent);
                            }
                        }
                        break;
                    }
                }
                else // go through the dom tree
                {
                    return moveToParentShape(shape);
                }
            }
            return null;
        }

        /// <summary>
        /// Moves to a child element.
        /// </summary>
        /// <param name="shape">The parent shape to get the child of.</param>
        /// <param name="index">The index of the child to get (infinite loop by modulo child count).</param>
        /// <returns>
        /// The child shape on the screen if possible otherwise the child in the DOM if possible otherwise <c>null</c>
        /// </returns>
        public static OoShapeObserver MoveToChild(OoShapeObserver shape, ref int index)
        {
            if (shape != null && shape.IsValid())
            {
                // move thought the Accessible tree because this elements should be visible on the screen and the tree navigation is faster
                // TODO remove false to use search in accessibility tree again. Currently accessing the page AccessibleShape element during search crashes openoffice!
                if (false && shape.AccComponent != null && shape.AccComponent.IsValid())
                {
                    OoAccComponent comp = shape.AccComponent;

                    while (true)
                    {
                        OoAccComponent child = moveToChildComponent(comp, ref index);

                        if (child != null && child != comp)
                        {
                            if (acceptAsUsableShape(child))
                            {
                                return getObserverForAccessible(child, shape.Page);
                            }
                            else
                            {
                                child = moveToChildComponent(child, ref index);
                            }
                        }
                        break;
                    }
                }
                else // go through the dom tree
                {
                    return moveToChildShape(shape, ref index);
                }
            }
            return null;
        }

        #endregion

        #region Accessible Tree Movement

        #region Deterministic [DEPRECATED]

        /// <summary>
        /// Moves to next accessible component in the accessible tree.
        /// </summary>
        /// <param name="comp">The start component.</param>
        /// <param name="handleChildren">if set to <c>true</c> the function will go deeper in the tree structure othewise it will only search for next sibblingsor go higher.</param>
        /// <returns>
        /// the next accessible component (first child, next sibling or parents next sibling)
        /// </returns>
        private static OoAccComponent moveDeterministicToNextComponent(OoAccComponent comp, bool handleChildren)
        {
            if (comp != null)
            {
                //walk through the children and back to the parent
                //check if has children
                if (handleChildren && comp.HasChildren) // should have children --> go deeper
                {
                    //TODO: check if we accept this children - e.g. Text ???

                    // try get the first child
                    OoAccComponent child = comp.GetChild(0);
                    if (child != null)
                    {
                        System.Diagnostics.Debug.WriteLine("[MOVE NEXT] ---> return child: " + child);
                        return child;
                    }
                }

                // if has no children --> go to next sibling

                // try to get the parent
                OoAccComponent parent = comp.GetParent();
                if (parent != null)
                {
                    int pIndex = comp.IndexInParent;
                    if (pIndex >= 0)
                    {
                        int pcCount = parent.ChildCount;
                        if (pIndex < pcCount - 1) // is not the last child
                        {
                            // get next sibling
                            var nextSibling = parent.GetChild(pIndex + 1);
                            if (nextSibling != null)
                            {
                                if (nextSibling.IsVisible())
                                {
                                    System.Diagnostics.Debug.WriteLine("[MOVE NEXT] ---> return next sibling: " + nextSibling);
                                    return nextSibling;
                                }
                                else // if the sibling is invalid
                                {
                                    return moveDeterministicToNextComponent(parent, handleChildren);
                                }
                            }
                        }
                        else // is last element --> go higher
                        {
                            return moveDeterministicToNextComponent(parent, false);
                        }
                    }
                    else
                    {
                        // ERROR - this happens e.g. when the parent object is invalid
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        Logger.Instance.Log(LogPriority.OFTEN, "AccDomWalker", "[ERROR] Index in parent is negative");
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Moves to prev component.
        /// </summary>
        /// <param name="comp">The comp.</param>
        /// <param name="handleChildren">if set to <c>true</c> [handle children].</param>
        /// <returns></returns>
        private static OoAccComponent moveDeterministicToPrevComponent(OoAccComponent comp, bool handleChildren = true)
        {
            if (comp != null)
            {
                //walk through the children and back to the parent
                //check if has children
                if (handleChildren && comp.HasChildren) // should have children --> go deeper
                {
                    //TODO: check if we accept this children - e.g. Text ???

                    // try get the first child
                    OoAccComponent child = comp.GetChild(comp.ChildCount - 1);
                    if (child != null)
                    {
                        System.Diagnostics.Debug.WriteLine("[MOVE PREV] ---> return child: " + child);
                        return child;
                    }
                }

                // if has no children --> go to prev sibling
                //--------------------------
                // try to get the parent
                OoAccComponent parent = comp.GetParent();
                if (parent != null)
                {
                    int pIndex = comp.IndexInParent;
                    if (pIndex >= 0)
                    {
                        int pcCount = parent.ChildCount;
                        if (pIndex > 0) // is not the first child
                        {
                            // get prev sibling
                            var prevSibling = parent.GetChild(pIndex - 1);
                            if (prevSibling != null)
                            {
                                System.Diagnostics.Debug.WriteLine("[MOVE PREV] ---> return prev sibling: " + prevSibling);
                                return prevSibling;
                            }
                        }
                        else // is first element --> go higher
                        {
                            return moveDeterministicToPrevComponent(parent, false);
                        }
                    }
                    else
                    {
                        // ERROR - this happens e.g. when the parent object is invalid
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);
                        Logger.Instance.Log(LogPriority.OFTEN, "AccDomWalker", "[ERROR] Index in parent is negative");
                    }
                }
            }
            return null;
        }

        #endregion

        private static OoAccComponent moveToNextComponent(OoAccComponent comp)
        {
            if (comp != null && comp.IsValid())
            {
                //--------------------------
                // try to get the parent
                OoAccComponent parent = comp.GetParent();
                if (parent != null)
                {
                    int pIndex = comp.IndexInParent;
                    int pcCount = parent.ChildCount;
                    int npIndex = mod(pIndex + 1, pcCount);

                    if (pIndex != npIndex)
                    {
                        // get next sibling
                        var nextSibling = parent.GetChild(npIndex);
                        if (nextSibling != null)
                        {
                            System.Diagnostics.Debug.WriteLine("[MOVE PREV] ---> return next sibling: " + nextSibling);
                            return nextSibling;
                        }
                    }
                }
            }

            return null;
        }

        private static OoAccComponent moveToPrevComponent(OoAccComponent comp)
        {
            if (comp != null && comp.IsValid())
            {
                //--------------------------
                // try to get the parent
                OoAccComponent parent = comp.GetParent();
                if (parent != null)
                {
                    int pIndex = comp.IndexInParent;
                    int pcCount = parent.ChildCount;
                    int npIndex = mod(pIndex - 1, pcCount);

                    if (pIndex != npIndex)
                    {
                        // get prev sibling
                        var prevSibling = parent.GetChild(npIndex);
                        if (prevSibling != null)
                        {
                            System.Diagnostics.Debug.WriteLine("[MOVE PREV] ---> return prev sibling: " + prevSibling);
                            return prevSibling;
                        }
                    }
                }
            }

            return null;
        }

        private static OoAccComponent moveToParentComponent(OoAccComponent comp)
        {
            if (comp != null && comp.IsValid())
            {
                try
                {
                    OoAccComponent parent = comp.GetParent();
                    if (parent != null && acceptAsUsableShape(parent))
                    {
                        return parent;
                    }
                }
                catch { }
            }

            return null;
        }

        private static OoAccComponent moveToChildComponent(OoAccComponent comp, ref int index)
        {
            if (comp != null && comp.IsValid() && comp.HasChildren)
            {
                try
                {
                    int childNumber = comp.ChildCount;
                    index = index % childNumber;

                    return comp.GetChild(index);

                }
                catch { }
            }

            return null;
        }

        #endregion

        #region DOM Tree Movement

        /// <summary>
        /// Moves to next sibling shape in the DOM.
        /// </summary>
        /// <param name="shape">The shape to start.</param>
        /// <returns></returns>
        private static OoShapeObserver moveToNextShape(OoShapeObserver shape)
        {
            if (shape != null)
            {
                return shape.GetNextSibling();
            }
            return null;
        }

        /// <summary>
        /// Moves to next sibling shape in the DOM.
        /// </summary>
        /// <param name="shape">The shape to start.</param>
        /// <returns></returns>
        private static OoShapeObserver moveToPreviousShape(OoShapeObserver shape)
        {
            if (shape != null )
            {
                return shape.GetPreviousSibling();
            }
            return null;
        }

        /// <summary>
        /// Moves to the requested child  of the shape in the DOM.
        /// </summary>
        /// <param name="shape">The parent shape.</param>
        /// <param name="number">The childs' index.</param>
        /// <returns>the child or <c>null</c></returns>
        private static OoShapeObserver moveToChildShape(OoShapeObserver shape, ref int number)
        {
            if (shape != null)
            {
                return shape.GetChild(number);
            }
            return null;
        }

        /// <summary>
        /// Moves to the parent of the shape in the DOM.
        /// </summary>
        /// <param name="shape">The child shape.</param>
        /// <returns>the parent or <c>null</c></returns>
        private static OoShapeObserver moveToParentShape(OoShapeObserver shape)
        {
            if (shape != null)
            {
                return shape.GetParent();
            }
            return null;
        }
        
        #endregion

        #region Helper

        /// <summary>
        /// Gets the OoShapeObserver for an accessible if already registered.
        /// </summary>
        /// <param name="accComp">The acc comp to get the observer for.</param>
        /// <param name="page">The page observer to search the observer in.</param>
        /// <returns>The related OoShapeObserver if already registered</returns>
        private static OoShapeObserver getObserverForAccessible(OoAccComponent accComp, OoDrawPageObserver page)
        {

            if (accComp != null && accComp.IsValid() && page != null)
            {
                var pObs = page.PagesObserver;
                if (pObs != null)
                {
                    OoShapeObserver sObs = pObs.GetRegisteredShapeObserver(accComp);

                    if (sObs == null) // shape is not registered
                    {
                        page.Update(); // try to do it better for the next term
                        AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error);

                        ////TODO: try to get the parent
                        return null;
                    }

                    return sObs;
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "OpenOfficeDrawShapeManipulator", "[ERROR] can't get the pages observer to get a ShapeObserver for the next item");
                }
            }
            return null;
        }

        /// <summary>
        /// Try to figure out if the accessible element is usable as interactive Shape
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        private static bool acceptAsUsableShape(OoAccComponent element)
        {
            if (element != null && element != null)
            {
                // check if the child is the paragraph
                if (element.ServiceNames.Contains(tud.mci.tangram.util.OO.Services.TEXT_ACCESSIBLE_PARAGRAPH))
                {
                    return false;
                }
                else
                {
                    if (element.Name.StartsWith("PageShape: "))
                    {
                        //TODO: decide if is a proper result
                        return false;
                    }
                }
            }

            return element != null;
        }
        
        /// <summary>
        /// Modulo function for negative values
        /// </summary>
        /// <param name="x">The value to mod.</param>
        /// <param name="m">The modulo.</param>
        /// <returns>an always positive modulo</returns>
        static int mod(int x, int m)
        {
            int r = x % m;
            return r < 0 ? r + m : r;
        }

        #endregion
    }
}
