using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.container;
using System.Threading.Tasks;
using tud.mci.tangram.Accessibility;
using unoidl.com.sun.star.accessibility;
using tud.mci.tangram.util;
using System.Threading;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.lang;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Observes one XDrawPage and collect the XShapes on this page
    /// </summary>
    public class OoDrawPageObserver : PropertiesEventForwarderBase, IUpdateable, IDisposable, IDisposingObserver
    {
        #region Members
        #region public
        /// <summary>
        /// The DOM object for the page object [XDrawPage]
        /// </summary>
        public Object DrawPage_anonymouse {get{return DrawPage;}}
        /// <summary>
        /// The DOM object for the page object
        /// </summary>
        internal readonly XDrawPage DrawPage;
        /// <summary>
        /// the corresponding observer for all pages in this DRAW document (the parent)
        /// </summary>
        public readonly OoDrawPagesObserver PagesObserver;
        /// <summary>
        /// the XAccessible counterpart to the DOM object [XAccessible]
        /// </summary>
        public Object AcccessibleCounterpart_anonymouse { get { return AcccessibleCounterpart; } }
        /// <summary>
        /// the XAccessible counterpart to the DOM object
        /// </summary>
        internal XAccessible AcccessibleCounterpart { get; set; }

        /// <summary>
        /// Gets the orientation of the page - portrait or landscape.
        /// </summary>
        /// <value>
        /// The orientation of the page.
        /// </value>
        public tud.mci.tangram.util.OO.PaperOrientation Orientation { get; private set; }
        /// <summary>
        /// Gets the width of the page in 100th/mm.
        /// </summary>
        /// <value>
        /// The width of the page.
        /// </value>
        public int Width { get; private set; }
        /// <summary>
        /// Gets the height of the page in 100th/mm.
        /// </summary>
        /// <value>
        /// The height of the page.
        /// </value>
        public int Height { get; private set; }
        /// <summary>
        /// Gets the padding from the drawing area to the page end at the bottom.
        /// </summary>
        /// <value>
        /// The padding at the bottom.
        /// </value>
        public int BorderBottom { get; private set; }
        /// <summary>
        /// Gets the padding from the drawing area to the page end at the bottom.
        /// </summary>
        /// <value>
        /// The padding at the bottom.
        /// </value>
        public int BorderLeft { get; private set; }
        /// <summary>
        /// Gets the padding from the drawing area to the page end at the left.
        /// </summary>
        /// <value>
        /// The padding at the left.
        /// </value>
        public int BorderRight { get; private set; }
        /// <summary>
        /// Gets the padding from the drawing area to the page end at the top.
        /// </summary>
        /// <value>
        /// The padding at the top.
        /// </value>
        public int BorderTop { get; private set; }

        /// <summary>
        /// List of all found and available XShapes in this draw page
        /// </summary>
        public readonly List<OoShapeObserver> shapeList = new List<OoShapeObserver>();
        #endregion

        #region private

        private System.Timers.Timer refreshPagePropertiesTimer;  // e.g. every 2.5s

        #endregion
        #endregion

        #region Constructor / Destructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoDrawPageObserver"/> class.
        /// </summary>
        /// <param name="dp">The DOM page object to observe</param>
        /// <param name="parent">The parent obeserver for all pages in this DRAW document.</param>
        public OoDrawPageObserver(XDrawPage dp, OoDrawPagesObserver parent)
            : base()
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "Create new PageObserver");
            DrawPage = dp;
            PagesObserver = parent;

            setAccessibleCounterpart();
            handleChildren();

            registerListeners();
        }

        ~OoDrawPageObserver()
        {
            refreshPagePropertiesTimer.Stop();
            try
            {
                if (DrawPage != null)
                {
                    XPropertySet pageProperties = (XPropertySet)DrawPage;
                    if (pageProperties != null)
                    {

                        pageProperties.removePropertyChangeListener("", eventForwarder);
                        pageProperties.removeVetoableChangeListener("", eventForwarder);

                    }
                    XMultiPropertySet pageMultiProperties = (XMultiPropertySet)DrawPage;
                    if (pageMultiProperties != null)
                    {
                        pageMultiProperties.removePropertiesChangeListener(eventForwarder);
                    }
                }
            }
            catch (Exception) { }
        }

        #endregion 

        #region Properties

        private static volatile bool _propertyUpdating;
        private DateTime _lastPagePropertyUpdate = DateTime.Now;
        private static readonly TimeSpan _propertyUpdateTimespan = new TimeSpan(0, 0, 0, 1);
        private static readonly Random rnd = new Random(DateTime.Now.Millisecond);
        /// <summary>
        /// registers various property listeners
        /// </summary>
        private void registerListeners()
        {
            try
            {
                if (refreshPagePropertiesTimer != null) { refreshPagePropertiesTimer.Stop(); refreshPagePropertiesTimer.Dispose(); }
                refreshPagePropertiesTimer = new System.Timers.Timer(5000);
                XPropertySet pageProperties = DrawPage as XPropertySet;
                if (pageProperties != null)
                {
                    //Debug.PrintAllProperties(pageProperties);

                    onRefreshPage(this, null);

                    pageProperties.addPropertyChangeListener("Orientation", eventForwarder);
                    pageProperties.addVetoableChangeListener("Orientation", eventForwarder);

                    pageProperties.addPropertyChangeListener("", eventForwarder);

                    refreshPagePropertiesTimer.Elapsed += new System.Timers.ElapsedEventHandler(onRefreshPage);

                    // start the refresh timer with some delay so not all pages will update their properties at the same time!
                    Task t = new Task(new Action(() =>
                    {
                        Thread.Sleep(rnd.Next(1000));
                        refreshPagePropertiesTimer.Start();
                    }));
                    t.Start();

                    //eventForwarder.Modified += eventForwarder_Modified;
                    //eventForwarder.NotifyEvent += eventForwarder_NotifyEvent;
                    //eventForwarder.PropertiesChange += eventForwarder_PropertiesChange;
                    //eventForwarder.PropertyChange += eventForwarder_PropertyChange;
                    //eventForwarder.VetoableChange += eventForwarder_VetoableChange;

                }
                XMultiPropertySet pageMultiProperties = (XMultiPropertySet)DrawPage;
                if (pageMultiProperties != null)
                {
                    pageMultiProperties.addPropertiesChangeListener(new string[] { }, eventForwarder);
                }
            }
            catch (DisposedException) { ((IDisposable)this).Dispose(); }
            catch (Exception) { }
        }
        #region event forwarder listeners

        //void eventForwarder_VetoableChange(object sender, ForwardedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //void eventForwarder_PropertyChange(object sender, ForwardedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //void eventForwarder_PropertiesChange(object sender, ForwardedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //void eventForwarder_NotifyEvent(object sender, ForwardedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //void eventForwarder_Modified(object sender, ForwardedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        #endregion

        private void onRefreshPage(object source, System.Timers.ElapsedEventArgs e)
        {
            updatePageProperties();
        }

        /// <summary>
        /// Updates the page properties.
        /// </summary>
        /// <param name="forceUpdate">if set to <c>true</c> to force an update even if the page is currently not the active one.</param>
        /// <returns>
        ///   <c>true</c> if the update was successfully fulfilled.
        /// </returns>
        /// <remarks>This function is time limited to 200 ms.</remarks>
        public bool updatePageProperties(bool forceUpdate = false)
        {
            if (!forceUpdate && !IsActive()) 
                return true;

            DateTime now = DateTime.Now;
            if (now - _lastPagePropertyUpdate < _propertyUpdateTimespan) return false;

            //System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " +++++++++ Update page properties REQUEST ["+this.GetHashCode()+"] +++++++++");

            bool successs = false;
            try
            {
                if (!_propertyUpdating && DrawPage != null)
                {
                    _propertyUpdating = true;
                    try
                    {
                        successs = TimeLimitExecutor.WaitForExecuteWithTimeLimit(200, new Action(() =>
                        {
                            Width = OoUtils.GetIntProperty(DrawPage, "Width");
                            Height = OoUtils.GetIntProperty(DrawPage, "Height");
                            BorderBottom = OoUtils.GetIntProperty(DrawPage, "BorderBottom");
                            BorderLeft = OoUtils.GetIntProperty(DrawPage, "BorderLeft");
                            BorderRight = OoUtils.GetIntProperty(DrawPage, "BorderRight");
                            BorderTop = OoUtils.GetIntProperty(DrawPage, "BorderTop");
                            Orientation = (tud.mci.tangram.util.OO.PaperOrientation)OoUtils.GetProperty(DrawPage, "Orientation");
                        }), "UpdatePageProperties");
                    }
                    catch (NullReferenceException) { ((IDisposable)this).Dispose(); }
                    catch (DisposedException)
                    {
                        ((IDisposable)this).Dispose();
                    }
                    _lastPagePropertyUpdate = now;
                    return successs;
                }
            }
            catch { }
            finally { _propertyUpdating = false; }
            return successs;
        }

        /// <summary>
        /// The cached page number. This field is updated by the <see cref="GetPageNum"/> function.
        /// Use this only to not stress the API to much.
        /// </summary>
        internal int CachedPageNum = 0;
        /// <summary>
        /// Gets the page number.
        /// </summary>
        /// <returns>The current number of this page (Numbers are starting with 1)</returns>
        /// <remarks>This function is time limited to 100 ms.</remarks>
        public int GetPageNum()
        {
            TimeLimitExecutor.WaitForExecuteWithTimeLimit(
                100,
                new Action(() => { CachedPageNum = OoUtils.GetIntProperty(DrawPage, "Number"); })
                , "PageObserver-GetPageNum");

            return CachedPageNum;
        }

        /// <summary>
        /// Gets the preview image bitmap data for the page.
        /// </summary>
        /// <param name="bitmapData">The bitmap data.</param>
        /// <remarks>This function is time limited to 100 ms.</remarks>
        public void GetPreview(out byte[] bitmapData)
        {
            byte[] bitmapDataResult = new byte[0];
            XPropertySet pageProperties = DrawPage as XPropertySet;
            if (pageProperties != null)
            {
                TimeLimitExecutor.WaitForExecuteWithTimeLimit(100,
                    () =>
                    {
                        try { bitmapDataResult = (byte[])pageProperties.getPropertyValue("Preview").Value; }
                        catch (ThreadAbortException) { }
                    }
                    , "GetPrieviewBitmap");
            }
            bitmapData = bitmapDataResult;
        }

        int _requestAmount = 0;
        readonly int _maxRequestAmountUntilRefresh = 10;
        /// <summary>
        /// Determines whether this page is the active one or not.
        /// </summary>
        /// <param name="forceRefresh">if set to <c>true</c> it forces a refresh on the real page properties, but not on the cached data - can force performance problems.</param>
        /// <returns>
        ///   <c>true</c> if this page is active; otherwise, <c>false</c>.
        /// </returns>
        /// <remarks>This function uses cached data if it is called frequently.</remarks>
        public bool IsActive(bool forceRefresh = false)
        {
            bool active = false;

            // get active Page
            if (PagesObserver != null && PagesObserver.DocWnd != null)
            {
                int activePageId = -1;
                if (forceRefresh 
                    || _requestAmount++ > _maxRequestAmountUntilRefresh 
                    || (this.PagesObserver != null && this.PagesObserver.DocWnd != null && this.PagesObserver.DocWnd.CachedCurrentPid < 1))
                {
                    if (PagesObserver.DocWnd.Controller != null && PagesObserver.DocWnd.Controller is XDrawView)
                    {
                        // refresh caches
                        var page = ((XDrawView)PagesObserver.DocWnd.Controller).getCurrentPage();

                        if (page.Equals(this.DrawPage))
                        {
                            this.PagesObserver.DocWnd.CachedCurrentPid = activePageId = GetPageNum();
                            active = true;
                        }
                        _requestAmount = 0;
                    }
                }
                else
                {
                    // check caches
                    if (this.PagesObserver != null && this.PagesObserver.DocWnd != null)
                    {
                        activePageId = PagesObserver.DocWnd.CachedCurrentPid;
                    }

                    active = activePageId == (CachedPageNum > 0 ? CachedPageNum : GetPageNum());
                }
            }
            return active;
        }

        #endregion

        /// <summary>
        /// find the corresponding accessible counterpart object for this page in the accessible tree
        /// </summary>
        private void setAccessibleCounterpart()
        {
            String name = OoUtils.GetStringProperty(DrawPage, "LinkDisplayName");
            name = "PageShape: " + name; //TODO: find a more general way

            XAccessible accCounter = OoAccessibility.GetAccessibleChildWithName(name, PagesObserver.Document.AccComp as unoidl.com.sun.star.accessibility.XAccessibleContext);
            if (accCounter != null) AcccessibleCounterpart = accCounter;
        }

        #region Child Handling

        private volatile bool _updating = false;
        /// <summary>
        /// collects all XShapes of this page and create a ShapeObserer for each one.
        /// </summary>
        private void handleChildren()
        {
            if (_updating) return;
            _updating = true;
            try
            {
                if (DrawPage != null)
                {
                    XIndexAccess ia = DrawPage as XIndexAccess;
                    if (ia != null)
                    {
                        int childs = ia.getCount();

                        for (int i = 0; i < childs; i++)
                        {
                            // Attention -> in asynchronous handling the index pointer i can be increased before the 
                            // thread was handled. So clone them to make sure that it is not changing while handling!
                            int j = i;
                            handleChild(j, ia);
                        }

                        //FIXME: use this if OpenOffice stops falling into dead lock while using this faster function
                        //Parallel.For(0, ia.getCount(), (i) => { handleChild(i, ref ia); });

                    }
                }
            }
            catch { }
            finally { _updating = false; }
        }

        readonly Object _childHandleLock = new Object();
        private void handleChild(int i, XIndexAccess ia)
        {
            //System.Diagnostics.Debug.WriteLine("[UPDATE] --- handle child [" + i + "]");
            try
            {
                if (PagesObserver != null)
                {
                    lock (_childHandleLock)
                    {
                        var anyShape = ia.getByIndex(i);
                        if (anyShape.hasValue() && anyShape.Value is XShape)
                        {
                            if (PagesObserver.ShapeAlreadyRegistered(anyShape.Value as XShape, this))
                            {
                                //System.Diagnostics.Debug.WriteLine("[UPDATE] Shape " + anyShape.Value + " already exists ");

                                OoShapeObserver so = PagesObserver.GetRegisteredShapeObserver(anyShape.Value as XShape, this);
                                if (so != null) so.UpdateChildren();
                                else
                                {
                                    Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Shape should exist but could not been found!!!");

                                    so = OoShapeObserverFactory.BuildShapeObserver(anyShape.Value, this);  //new OoShapeObserver(anyShape.Value as XShape, this);
                                    //shapeList.Add(so);
                                }
                            }
                            else
                            {
                                //System.Diagnostics.Debug.WriteLine("[UPDATE] New Shape " + anyShape.Value + " will be registered ");
                                OoShapeObserver so = null;
                                try
                                {
                                    if (OoUtils.ElementSupportsService(anyShape.Value, OO.Services.DRAW_SHAPE_TEXT))
                                    {
                                        so = OoShapeObserverFactory.BuildShapeObserver(anyShape.Value, this);  //new OoShapeObserver(anyShape.Value as XShape, this);
                                    }
                                    else
                                    {
                                        so = OoShapeObserverFactory.BuildShapeObserver(anyShape.Value, this);  //new OoShapeObserver(anyShape.Value as XShape, this);
                                        //System.Diagnostics.Debug.WriteLine("[UPDATE] Shape: " + so.Name + " will be registered");
                                    }
                                }
                                catch (unoidl.com.sun.star.uno.RuntimeException ex)
                                {
                                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR]  internal while register ShapeObserver", ex);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] can not register ShapeObserver", ex);
                                }
                                //finally
                                //{
                                //    if (so != null) shapeList.Add(so);
                                //}
                            }
                        }
                    }
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] PagesObserver is null");
                }
            }
            catch (System.Threading.ThreadAbortException ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "[OO Deadlock] can't get access to children via child handling in DrawPageObserver", ex); }
            catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't get access to children via child handling in DrawPageObserver", ex); }
        }

        /// <summary>
        /// Gets the first child of the current page.
        /// </summary>
        /// <returns></returns>
        public OoShapeObserver GetFirstChild()
        {
            if (DrawPage != null && PagesObserver != null)
            {
                if (DrawPage is XDrawPage)
                {
                    try
                    {
                        int childCount = ((XDrawPage)DrawPage).getCount();
                        if (childCount > 0)
                        {
                            var anyFirstChild = ((XDrawPage)DrawPage).getByIndex(0);
                            if (anyFirstChild.hasValue())
                            {
                                var firstChild = anyFirstChild.Value;

                                if (firstChild is XShape)
                                {
                                    var shapeObs = PagesObserver.GetRegisteredShapeObserver(firstChild as XShape, this);
                                    if (shapeObs != null)
                                    {
                                        if (shapeObs.AcccessibleCounterpart == null)
                                        {
                                            shapeObs.Update();
                                        }
                                    }
                                    else
                                    {
                                        //TODO: register this shape
                                        OoShapeObserver newShapeObserver = OoShapeObserverFactory.BuildShapeObserver(firstChild, this);  //new OoShapeObserver(firstChild as XShape, this);
                                        PagesObserver.RegisterUniqueShape(newShapeObserver);
                                        return newShapeObserver;
                                    }
                                    return shapeObs;
                                }
                                else
                                {
                                    //util.Debug.GetAllInterfacesOfObject(firstChild);
                                    Logger.Instance.Log(LogPriority.DEBUG, this, "[UNEXPECTED] The first child of a page is not a shape!: ");
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] Can't get first child: " + ex);
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the last child on the current.
        /// </summary>
        /// <returns></returns>
        public OoShapeObserver GetLastChild()
        {
            if (DrawPage != null && PagesObserver != null)
            {
                if (DrawPage is XDrawPage)
                {
                    try
                    {
                        int childCount = ((XDrawPage)DrawPage).getCount();
                        if (childCount > 0)
                        {
                            var anyLastChild = ((XDrawPage)DrawPage).getByIndex(childCount - 1);
                            if (anyLastChild.hasValue())
                            {
                                var lastChild = anyLastChild.Value;
                                //util.Debug.GetAllInterfacesOfObject(lastChild);

                                if (lastChild is XShape)
                                {
                                    return PagesObserver.GetRegisteredShapeObserver(lastChild as XShape, this);
                                }
                                else
                                {
                                    Logger.Instance.Log(LogPriority.DEBUG, this, "[UNEXPECTED] The first child of a page is not a shape!");
                                }

                            }

                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] Can't get last child: " + ex);
                    }

                }
            }

            return null;
        }

        #endregion

        #region IUpdateable
        public void Update()
        {
            updatePageProperties();
            handleChildren();
        }
        #endregion

        #region IDisposable

        void IDisposable.Dispose()
        {
            this.refreshPagePropertiesTimer.Stop();
            this.AcccessibleCounterpart = null;

            try { foreach (var oS in shapeList) { oS.Dispose(); } }
            catch { }

            this.shapeList.Clear();
            Disposed = true;
            fire_DisposingEvent();
            OO.CheckConnection();
        }

        #endregion

        #region IDisposingObserver

        /// <summary>
        /// Gets a value indicating whether this <see cref="AbstractDisposingBase"/> is disposed.
        /// </summary>
        /// <value><c>true</c> if disposed; otherwise, <c>false</c>.</value>
        public bool Disposed { get; private set; }

        /// <summary>
        /// Occurs when this object is disposing.
        /// </summary>
        public event EventHandler ObserverDisposing;

        /// <summary>
        /// Fires the disposing event.
        /// </summary>
        protected virtual void fire_DisposingEvent()
        {
            if (ObserverDisposing != null)
            {
                try
                {
                    ObserverDisposing.Invoke(this, null);
                }
                catch { }
            }
        }

        #endregion

    }
}