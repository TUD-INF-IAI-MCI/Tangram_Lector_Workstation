using System;
using System.Collections.Generic;
using System.Threading;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;

namespace tud.mci.tangram.Accessibility
{
    /// <summary>
    /// Encapsulates some important properties and functions of accessible objects.
    /// </summary>
    public class OoAccComponent
    {
        #region Members

        internal XAccessibleComponent AccComp { get; private set; }
        internal XAccessibleContext AccCont { get; private set; }

        private readonly object _accLock = new Object();
        private readonly object _accScreenBoundsLock = new Object();
              

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccComponent"/> struct.
        /// </summary>
        /// <param name="accComp">The base XAccessibleComponent that shoulb be wraped.</param>
        public OoAccComponent(Object acc) : this(acc as XAccessibleComponent) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccComponent"/> struct.
        /// </summary>
        /// <param name="accComp">The base XAccessibleComponent that shoulb be wraped.</param>
        private OoAccComponent(XAccessibleComponent accComp)
        {
            this.AccComp = accComp;
            AccCont = (AccComp != null && AccComp is XAccessibleContext) ? AccComp as XAccessibleContext :
                AccComp is XAccessible ? ((XAccessible)AccComp).getAccessibleContext() : null;
        }

        #endregion

        #region Property Access

        System.Drawing.Rectangle _lastBounds = new System.Drawing.Rectangle();
        /// <summary>
        /// The bounding box of this element and the position inside the frame.
        /// </summary>
        public System.Drawing.Rectangle Bounds
        {
            get
            {
                lock (_accLock)
                {
                    if (AccComp != null)
                    {
                        try
                        {
                            var b = AccComp.getBounds();
                            _lastBounds = new System.Drawing.Rectangle(b.X, b.Y, b.Width, b.Height);
                        }
                        catch (unoidl.com.sun.star.lang.DisposedException) { AccCont = null; AccComp = null; IsValid(); }
                        catch (ThreadAbortException) { Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility (Bounds)", "can't get bounds of accessible object"); }
                        catch (Exception e) { Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility (Bounds)", "can't get bounds of accessible object", e); }
                    }
                    return _lastBounds;
                }
            }
        }

        private System.Drawing.Rectangle _lastScreenBounds = new System.Drawing.Rectangle();
        /// <summary>
        /// The bounding box of this element and its position on the screen.
        /// </summary>
        /// <value>The screen bounds.</value>
        /// <remarks>Parts of this function are time limited to 200 ms.</remarks>
        public System.Drawing.Rectangle ScreenBounds
        {
            get
            {
                lock (_accScreenBoundsLock)
                {
                    if (AccComp != null
                        //&& (OoAccessibility.GetAccessibleRole(AccComp as XAccessible) != AccessibleRole.INVALID)
                                )
                    {
                        System.Drawing.Rectangle sB = _lastScreenBounds;
                        try
                        {
                            var screenbounder = TimeLimitExecutor.WaitForExecuteWithTimeLimit(200,
                                () =>
                                {
                                    try
                                    {
                                        var b = Bounds;
                                        var screenLocation = AccComp != null ? AccComp.getLocationOnScreen() : new unoidl.com.sun.star.awt.Point();
                                        sB = new System.Drawing.Rectangle(screenLocation.X, screenLocation.Y, b.Width, b.Height);
                                        _lastScreenBounds = sB;
                                    }
                                    catch (Exception e)
                                    {
                                        System.Diagnostics.Debug.WriteLine("[ERROR] can't get screenbounds of window: " + e);
                                    }
                                    finally
                                    {
                                        sB = _lastScreenBounds;
                                    }
                                }
                                , "Get Screen bounds"
                                );
                        }
                        catch (System.Threading.ThreadAbortException)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "can't get bounds of accessible object - time out");
                        }
                        catch (Exception e)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "can't get bounds of accessible object", e);
                        }
                        return sB;
                    }
                    return new System.Drawing.Rectangle();
                }
            }
        }

        private AccessibleRole _lastRole = AccessibleRole.INVALID;
        /// <summary>
        /// The Accessible role (object type identifier) of this element.
        /// </summary>
        public AccessibleRole Role
        {
            get
            {
                lock (_accLock)
                {
                    if (_lastRole == AccessibleRole.INVALID || _lastRole == AccessibleRole.UNKNOWN)
                    {
                        if (AccComp != null)
                        {
                            _lastRole = OoAccessibility.GetAccessibleRole(AccComp as XAccessible);
                        }
                        else { _lastRole = AccessibleRole.INVALID; }
                    }
                    return _lastRole;
                }
            }
        }

        /// <summary>
        /// The accessible name of this element (combination of "Name" and "Title").
        /// </summary>
        public String Name
        {
            get
            {
                lock (_accLock)
                {
                    if (AccComp != null)
                    {
                        if (AccCont != null)
                        {
                            try
                            {
                                return AccCont.getAccessibleName();
                            }
                            catch (System.Exception){}
                        }
                    }
                    return String.Empty;
                }
            }
        }

        /// <summary>
        /// A user- or system defined description of this element ("Description").
        /// </summary>
        public String Description
        {
            get
            {
                lock (_accLock)
                {
                    if (AccComp != null)
                    {
                        XAccessibleContext context = AccComp as XAccessibleContext;
                        if (context != null)
                        {
                            return context.getAccessibleDescription();
                        }
                    }
                    return String.Empty;
                }
            }
        }

        /// <summary>
        /// List of current accessible states of this element (e.g. FOCUSED, SELECED, etc.).
        /// </summary>
        public List<AccessibleStateType> States
        {
            get
            {
                lock (_accLock)
                {
                    if (AccComp != null)
                    {
                        XAccessibleContext context = AccComp as XAccessibleContext;
                        if (context != null)
                        {
                            return OoAccessibility.GetAccessibleStates(context.getAccessibleStateSet());
                        }
                    }
                    return new List<AccessibleStateType>() { AccessibleStateType.INVALID };
                }
            }
        }

        /// <summary>
        /// get all Service names of the XAccessibleContext object 
        /// </summary>
        /// <value>The service names.</value>
        public List<String> ServiceNames
        {
            get
            {
                lock (_accLock)
                {
                    if (AccCont != null)
                    {
                        return new List<String>(util.Debug.GetAllServicesOfObject(AccCont, false));
                    }
                    return new List<string>();
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has children.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance has children; otherwise, <c>false</c>.
        /// </value>
        public bool HasChildren
        {
            get
            {
                if (AccCont != null)
                {
                    return ChildCount > 0;
                }
                return false;
            }

        }

        /// <summary>
        /// Gets the child count.
        /// </summary>
        /// <value>The child count.</value>
        public int ChildCount
        {
            get
            {
                lock (_accLock)
                {
                    if (AccCont != null)
                    {
                        return AccCont.getAccessibleChildCount();
                    }
                    return 0;
                }
            }
        }

        /// <summary>
        /// Gets the index of this element in parents child list.
        /// </summary>
        /// <value>The index in parent.</value>
        public int IndexInParent
        {
            get
            {
                lock (_accLock)
                {
                    if (AccCont != null)
                    {
                        return AccCont.getAccessibleIndexInParent();
                    }
                    return -1;
                }
            }
        }


        /// <summary>
        /// Determines whether this instance is valid.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </returns>
        /// <remarks>This function is time limited to 100 ms.</remarks>
        public bool IsValid()
        {
            if (AccComp != null && AccCont != null)
            {
                try
                {
                    int c = -1;
                    bool success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(100, () => { c = AccCont.getAccessibleChildCount(); });
                    if (c >= 0)
                        return true;
                    return false;
                }
                catch
                {
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// Determines whether this instance is visible (has positive bounds).
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance is visible; otherwise, <c>false</c>.
        /// </returns>
        /// <remarks>Parts of this function are time limited to 100 ms.</remarks>
        public bool IsVisible()
        {
            if (IsValid())
            {
                System.Drawing.Rectangle b = new System.Drawing.Rectangle();
                var bT = TimeLimitExecutor.ExecuteWithTimeLimit(100, () => { b = Bounds; });

                while (bT != null && bT.IsAlive && bT.ThreadState == ThreadState.Running)
                {
                    Thread.Sleep(1);
                }

                if (b != null && b.Height > 0 && b.Width > 0)
                {
                    //TODO: test the states for visible?!
                    return true;
                }

            }

            return false;
        }

        #endregion

        #region Function Access

        /// <summary>
        /// Determines whether the specified XAccessibleComponent contains the point.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <returns>
        /// 	<c>true</c> if the specified XAccessibleComponent contains point; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsPoint(int x, int y) { return ContainsPoint(new System.Drawing.Point(x, y)); }
        /// <summary>
        /// Determines whether the specified XAccessibleComponent contains the point.
        /// </summary>
        /// <param name="p">The p.</param>
        /// <returns>
        /// 	<c>true</c> if the specified XAccessibleComponent contains point; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsPoint(System.Drawing.Point p)
        {
            if (this.AccComp != null)
            {
                return this.AccComp.containsPoint(new unoidl.com.sun.star.awt.Point(p.X, p.Y));
            }
            return false;
        }

        /// <summary>
        /// Returns the Accessible child that is rendered under the given point.
        /// The test point's coordinates are defined relative to the coordinate system
        /// of the object. That means that when the object is an opaque rectangle then
        /// both the points (0,0) and (with-1,height-1) would yield a true value.
        /// </summary>
        /// <param name="x">Coordinates of the test point for which to find the Accessible child. The
        /// origin of the coordinate system is the upper left corner of the object's
        /// bounding box as returned by the getBounds. The scale of the coordinate
        /// system is identical to that of the screen coordinate system.</param>
        /// <param name="y">The y.</param>
        /// <returns>
        /// If there is one child which is rendered so that its bounding box contains
        /// the test point then a reference to that object is returned. If there is
        /// more than one child which satisfies that condition then a reference to
        /// that one is returned that is painted on top of the others. If no there
        /// is no child which is rendered at the test point an empty reference is
        /// returned.
        /// </returns>
        public XAccessible GetAccessibleFromPoint(int x, int y) { return GetAccessibleFromPoint(new System.Drawing.Point(x, y)); }
        /// <summary>
        /// Returns the Accessible child that is rendered under the given point. 
        /// The test point's coordinates are defined relative to the coordinate system 
        /// of the object. That means that when the object is an opaque rectangle then 
        /// both the points (0,0) and (with-1,height-1) would yield a true value.
        /// </summary>
        /// <param name="p">
        /// Coordinates of the test point for which to find the Accessible child. The 
        /// origin of the coordinate system is the upper left corner of the object's 
        /// bounding box as returned by the getBounds. The scale of the coordinate 
        /// system is identical to that of the screen coordinate system. 
        /// </param>
        /// <returns>
        /// If there is one child which is rendered so that its bounding box contains 
        /// the test point then a reference to that object is returned. If there is 
        /// more than one child which satisfies that condition then a reference to 
        /// that one is returned that is painted on top of the others. If no there 
        /// is no child which is rendered at the test point an empty reference is 
        /// returned.
        /// </returns>
        public XAccessible GetAccessibleFromPoint(System.Drawing.Point p)
        {
            XAccessible result = null;
            if (this.AccComp != null)
            {
                result = this.AccComp.getAccessibleAtPoint(new unoidl.com.sun.star.awt.Point(p.X, p.Y));
            }
            return result;
        }

        /// <summary>
        /// Returns the Accessible child that is rendered under the given point.
        /// </summary>
        /// <param name="p">Coordinates of the test point for which to find the Accessible child. The
        /// origin of the coordinate system the screen coordinate system.</param>
        /// <returns>
        /// If there is one child which is rendered so that its bounding box contains
        /// the test point then a reference to that object is returned. If there is
        /// more than one child which satisfies that condition then a reference to
        /// that one is returned that is painted on top of the others. If no there
        /// is no child which is rendered at the test point an empty reference is
        /// returned.
        /// </returns>
        public OoAccComponent GetAccessibleFromScreenPos(System.Drawing.Point p)
        {
            var sb = ScreenBounds;
            System.Drawing.Point p2 = new System.Drawing.Point(p.X - sb.Left, p.Y - sb.Top);
            var acc = GetAccessibleFromPoint(p2);
            return new OoAccComponent(acc);
        }
        /// <summary>
        /// Returns the Accessible child that is rendered under the given point.
        /// </summary>
        /// <param name="x">X-coordinate of the test point for which to find the Accessible child. The
        /// origin of the coordinate system the screen coordinate system.</param>
        /// <param name="y">The y-coordinate.</param>
        /// <returns>
        /// If there is one child which is rendered so that its bounding box contains
        /// the test point then a reference to that object is returned. If there is
        /// more than one child which satisfies that condition then a reference to
        /// that one is returned that is painted on top of the others. If no there
        /// is no child which is rendered at the test point an empty reference is
        /// returned.
        /// </returns>
        public OoAccComponent GetAccessibleFromScreenPos(int x, int y)
        {
            return GetAccessibleFromScreenPos(new System.Drawing.Point(x, y));
        }

        /// <summary>
        /// Gets the parent element.
        /// </summary>
        /// <returns>The Parent of the object or <c>null</c></returns>
        public OoAccComponent GetParent()
        {
            lock (_accLock)
            {
                if (AccCont != null)
                {
                    var par = AccCont.getAccessibleParent();

                    if (par != null) return new OoAccComponent(par);

                }
                return null;
            }
        }

        /// <summary>
        /// Gets the childs.
        /// </summary>
        /// <returns>List of accessible child elements</returns>
        public List<XAccessible> GetChilds()
        {
            lock (_accLock)
            {
                List<XAccessible> childs = new List<XAccessible>();
                if (AccCont != null)
                {
                    int c = ChildCount;
                    for (int i = 0; i < c; i++)
                    {
                        var child = AccCont.getAccessibleChild(i);
                        //if (child != null) // don't destroy the parent index!
                        childs.Add(child);
                    }
                }
                return childs;
            }
        }

        /// <summary>
        /// Gets the child of this element at the given index.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns>The child as wrapt object or <c>null</c></returns>
        public OoAccComponent GetChild(int index)
        {
            lock (_accLock)
            {
                if (AccCont != null)
                {
                    int cCount = ChildCount;
                    if (cCount > index)
                    {
                        try
                        {
                            var child = AccCont.getAccessibleChild(index);
                            if (child != null)
                            {
                                return new OoAccComponent(child);
                            }
                        }
                        catch { }
                    }
                }
                return null;
            }
        }



        #endregion

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            if (IsValid())
            {
                var sb = ScreenBounds;
                var b = Bounds;
                return "AccCompPos Bounds: " + "X: " + b.X + " Y: " + b.Y + " Width: " + b.Width + " Height: " + b.Height + " Position on Screen: X: " + sb.X + " Y: " + sb.Y;
            }
            return "[FATAL ERROR] Element is not valid";
        }
    }
}