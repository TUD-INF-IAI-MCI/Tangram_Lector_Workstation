using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.container;
using tud.mci.tangram.util;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Observes a XShape or an XShapes and his children
    /// </summary>
    public partial class OoShapeObserver : PropertiesEventForwarderBase, IUpdateable, IDisposable, IDisposingObserver
    {
        #region Child Handling

        /// <summary>
        /// Get all available observers for the children in the DOM tree.
        /// ATTENTION: this function brings high load to the connection to OO. It should bee avoided to use this function to often.
        /// Cache the result of this function for child handling etc.
        /// </summary>
        /// <returns>list of child observers</returns>
        public List<OoShapeObserver> GetChilderen()
        {
            List<OoShapeObserver> children = new List<OoShapeObserver>();

            if (Shape != null && !Disposed)
            {
                int childcount = ChildCount;
                if (childcount > 0)
                {
                    for (int i = 0; i < childcount; i++)
                    {
                        var child = GetChild(i);
                        if (child != null)
                        {
                            children.Add(child);
                        }
                    }
                }
            }

            return children;
        }


        /// <summary>
        /// collects or updates all the child shapes if the Shape is an XShapes (a group).
        /// </summary>
        private void handleChildren()
        {
            if (Shape is XShapes && ((XShapes)Shape).hasElements()) //Group
            {
                for (int i = 0; i < ((XShapes)Shape).getCount(); i++) { handleChild(i); }
                //FIXME: use this if OpenOffice stops falling into deadlock
                //Parallel.For(0, ((XShapes)Shape).getCount(), (i) => { handleChild(i); });
            }
        }

        private void handleChild(int i)
        {
            try
            {
                var child = ((XShapes)Shape).getByIndex(i);
                if (child.hasValue() && child.Value is XShape)
                {
                    if (Page != null && Page.PagesObserver != null)
                    {
                        // this updates the lists of observers
                        var _so = Page.PagesObserver.GetRegisteredShapeObserver(child.Value as XShape, Page);
                        _so.Parent = this;
                    }
                }
            }
            catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't get access to child via child handling in shapeObserver", ex); }
        }

        /// <summary>
        /// Updates the child list and check for new ones.
        /// </summary>
        public void UpdateChildren()
        {
            handleChildren();
            //TODO: remove disposed ones
        }

        ///// <summary>
        ///// Determine if the XShape child is already known as a child.
        ///// </summary>
        ///// <param name="shape">The shape.</param>
        ///// <returns> the releated ShapeObserver to the known child otherwise <c>null</c></returns>
        //OoShapeObserver childListContainsXShape(XShape shape)
        //{
        //    foreach (var child in Children)
        //    {
        //        if (child.Equals(shape)) return child;
        //    }
        //    return null;
        //}

        #endregion

        #region DOM Tree Handling

        /// <summary>
        /// gets an OoShapeObserver for the given shape.
        /// </summary>
        /// <param name="s">The s.</param>
        /// <returns>An already registered or a new observer for the shape</returns>
        private OoShapeObserver getShapeObserverFromXShape(XShape s)
        {
            if (s != null)
            {
                OoShapeObserver sObs = this.Page.PagesObserver.GetRegisteredShapeObserver(s, this.Page);
                if (sObs == null && this.Page != null && this.Page.PagesObserver != null)
                {
                    sObs = new OoShapeObserver(s,
                        this.Page.PagesObserver.GetRegisteredPageObserver(
                            OoDrawUtils.GetPageForShape(s)));
                    this.Page.PagesObserver.RegisterUniqueShape(sObs);
                }
                return sObs;
            }
            return null;
        }

        /// <summary>
        /// Goes to the next DOM sibling if possible 
        /// </summary>
        /// <returns>The observer for the next sibling (infinite child loop) or the same if there is only on child or <c>null</c> if no sibling could be found.</returns>
        public OoShapeObserver GetNextSibling()
        {
            if (this.Page != null && this.Page.PagesObserver != null)
            {
                XShape s = getNextSiblingByXShape();
                OoShapeObserver sobs = getShapeObserverFromXShape(s);

                if (sobs.Disposed)
                {
                    sobs = new OoShapeObserver(s, Page);
                    Page.PagesObserver.RegisterUniqueShape(sobs);
                }

                return sobs;

            }
            return null;
        }

        XShape getNextSiblingByXShape()
        {
            //TODO: infinite loop -- so if this is the last, than return the first
            if (Shape != null && Shape is XChild)
            {
                XShapes parent = getParentByXShape();
                if (parent != null && parent.hasElements())
                {
                    int childCount = ((XShapes)parent).getCount();
                    if (childCount > 0)
                    {
                        //string current_name = Name;

                        bool lastWasCurrent = false;
                        //find this child
                        for (int i = 0; i < childCount; i++)
                        {
                            var c = parent.getByIndex(i).Value;

                            if (lastWasCurrent && c != null && c is XShape)
                            {
                                return c as XShape;
                            }
                            if (c == Shape)
                            {
                                lastWasCurrent = true;
                            }
                        }
                        // if no child was in front, return the las child at all
                        return getChildByXShape(0, parent as XShapes);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Goes to the previous DOM sibling if possible
        /// </summary>
        /// <returns>The observer for the prevoiuse sibling (infinite child loop) or the same if there is only on child or <c>null</c> if no sibling could be found.</returns>
        public OoShapeObserver GetPreviousSibling()
        {
            if (this.Page != null && this.Page.PagesObserver != null)
            {
                XShape s = getPreviousSiblingByXShape();
                OoShapeObserver sobs = getShapeObserverFromXShape(s);

                if (sobs.Disposed)
                {
                    sobs = new OoShapeObserver(s, Page);
                    Page.PagesObserver.RegisterUniqueShape(sobs);
                }

                return sobs;
            }
            return null;
        }

        /// <summary>
        /// Gets the previous sibling by X shape.
        /// </summary>
        XShape getPreviousSiblingByXShape()
        {
            //TODO: infinite loop -- so if this is the first, than return the last
            if (Shape != null && Shape is XChild)
            {
                XShapes parent = getParentByXShape();
                if (parent != null && parent.hasElements())
                {
                    int childCount = parent.getCount();
                    if (childCount > 0)
                    {
                        //string current_name = Name;

                        XShape lastChild = null;
                        //find this child
                        for (int i = 0; i < childCount; i++)
                        {
                            var c = ((XShapes)parent).getByIndex(i).Value;

                            if (c == Shape)
                            {
                                if (lastChild != null)
                                {
                                    return lastChild;
                                }
                            }
                            else if (c is XShape)
                            {
                                lastChild = c as XShape;
                            }
                        }

                        // if no child was in front, return the las child at all
                        return getChildByXShape(-1, parent as XShapes);

                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Goes to the parent in DOM if possible.
        /// </summary>
        /// <returns>The observer for the direct parent or <c>null</c> if no parent could be found.</returns>
        public OoShapeObserver GetParent()
        {
            if (this.Page != null && this.Page.PagesObserver != null)
            {
                XShapes s = getParentByXShape();
                return getShapeObserverFromXShape(s as XShape);
            }
            return null;
        }

        /// <summary>
        /// Gets the parent by X shape.
        /// </summary>
        XShapes getParentByXShape(XShape shape = null)
        {
            XShapes parent;
            bool successs = tryGetParentByXShape(out parent, shape);
            return parent;
        }

        /// <summary>
        /// Tries to get the parent of an XShape.
        /// </summary>
        /// <param name="parent">The parent. 
        /// Can be <c>NULL</c> if no parent is available - could bee the case when the 
        /// shape was deleted but not disposed for keeping it in undo/history</param>
        /// <param name="shape">The shape.</param>
        /// <returns><c>true</c> if the parent could been get, otherwise <c>false</c></returns>
        bool tryGetParentByXShape(out XShapes parent, XShape shape = null)
        {
            bool success = false;
            XShapes par = null;

            TimeLimitExecutor.WaitForExecuteWithTimeLimit(200, () =>
            {
                if (shape == null) shape = Shape;
                if (shape != null && shape is XChild)
                {
                    var Parent = ((XChild)shape).getParent();
                    if (Parent != null)
                    {
                        par = Parent as XShapes;
                        success = true;
                    }
                }
            }, "GetParent");

            parent = par;
            return success;
        }

        /// <summary>
        /// Goes to the child at given index in DOM if possible.
        /// </summary>
        /// <param name="number">The number. Will be handled by modulo child count. 
        /// So this should not get invalid. It is also possible to receive the 
        /// last child by asking for '-1'</param>
        /// <returns>The observer for the child at given index (infinite child loop by modulu of child count) or <c>null</c> if no child could be found.</returns>
        public OoShapeObserver GetChild(int number)
        {
            if (this.Page != null && this.Page.PagesObserver != null)
            {
                XShape s = getChildByXShape(number);
                return getShapeObserverFromXShape(s);
            }
            return null;
        }

        static int mod(int x, int m)
        {
            int r = x % m;
            return r < 0 ? r + m : r;
        }

        /// <summary>
        /// Gets the child by X shape.
        /// </summary>
        /// <param name="index">The index.</param>
        XShape getChildByXShape(int number, XShapes shape = null)
        {
            if (shape == null) shape = Shape as XShapes;
            if (shape != null && shape is XShapes && ((XShapes)shape).hasElements())
            {
                int cCount = ((XShapes)shape).getCount();
                number = mod(number, cCount);

                return ((XShapes)shape).getByIndex(number).Value as XShape;
            }
            return null;
        }

        #endregion

    }
}
