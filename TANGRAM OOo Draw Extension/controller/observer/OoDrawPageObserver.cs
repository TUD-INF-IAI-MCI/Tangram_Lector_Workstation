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
        /// The DOM object for the page object
        /// </summary>
        public readonly XDrawPage DrawPage;
        /// <summary>
        /// the corresponding observer for all pages in this DRAW document (the parent)
        /// </summary>
        public readonly OoDrawPagesObserver PagesObserver;
        /// <summary>
        /// the XAccessible counterpart to the DOM object
        /// </summary>
        public XAccessible AcccessibleCounterpart { get; set; }

        public unoidl.com.sun.star.view.PaperOrientation Orientation { get; private set; }
        public int Width { get; private set; }
        public int Height { get; private set; }
        public int BorderBottom { get; private set; }
        public int BorderLeft { get; private set; }
        public int BorderRight { get; private set; }
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
        /// <summary>
        /// registers various property listeners
        /// </summary>
        private void registerListeners()
        {
            try
            {
                refreshPagePropertiesTimer = new System.Timers.Timer(5000);
                XPropertySet pageProperties = DrawPage as XPropertySet;
                if (pageProperties != null)
                {
                    //Debug.PrintAllProperties(pageProperties);

                    onRefreshPage(this, null);

                    pageProperties.addPropertyChangeListener("Orientation", eventForwarder);
                    pageProperties.addVetoableChangeListener("Orientation", eventForwarder);


                    refreshPagePropertiesTimer.Elapsed += new System.Timers.ElapsedEventHandler(onRefreshPage);
                    refreshPagePropertiesTimer.Start();
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
            catch (Exception ex)
            {

            }
        }

        private volatile bool _pageUpdating = false;

        private void onRefreshPage(object source, System.Timers.ElapsedEventArgs e)
        {
            if (_pageUpdating) return;

            _pageUpdating = true;
            TimeLimitExecutor.WaitForExecuteWithTimeLimit(300, () => {
                UpdatePageProperties(); 
            }, "RefreshPageProperteis");
            _pageUpdating = false;
        }

        public bool UpdatePageProperties()
        {
            bool successs = false;

            if (DrawPage != null)
            {

                try
                {
                    //util.Debug.GetAllProperties(DrawPage);
                    //util.Debug.GetAllProperties(this.PagesObserver.Controller);

                    //pageProperties.
                    unoidl.com.sun.star.view.PaperOrientation orientation = (unoidl.com.sun.star.view.PaperOrientation)OoUtils.GetProperty(DrawPage, "Orientation");
                    Width = OoUtils.GetIntProperty(DrawPage, "Width");
                    Height = OoUtils.GetIntProperty(DrawPage, "Height");
                    BorderBottom = OoUtils.GetIntProperty(DrawPage, "BorderBottom");
                    BorderLeft = OoUtils.GetIntProperty(DrawPage, "BorderLeft");
                    BorderRight = OoUtils.GetIntProperty(DrawPage, "BorderRight");
                    BorderTop = OoUtils.GetIntProperty(DrawPage, "BorderTop");

                    //if (this.PagesObserver != null && this.PagesObserver.Controller != null)
                    //{
                    //    ZoomValue = OoUtils.GetIntProperty(this.PagesObserver.Controller, "ZoomValue");
                    //    unoidl.com.sun.star.awt.Point offset = OoUtils.GetProperty(this.PagesObserver.Controller, "ViewOffset") as unoidl.com.sun.star.awt.Point;
                    //    if(offset != null)
                    //        ViewOffset = new System.Drawing.Point(offset.X, offset.Y);
                    //}

                    successs = true;
                }
                catch (NullReferenceException) { ((IDisposable)this).Dispose(); }
                catch (DisposedException)
                {
                    ((IDisposable)this).Dispose();
                }
            }
            return successs;
        }


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
                            // Attention -> in assynchronouse handling the index pointer i can be increased before the 
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

                                    so = new OoShapeObserver(anyShape.Value as XShape, this);
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
                                        so = new OoShapeObserver(anyShape.Value as XShape, this);
                                    }
                                    else
                                    {
                                        so = new OoShapeObserver(anyShape.Value as XShape, this);
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
                                        OoShapeObserver newShapeObserver = new OoShapeObserver(firstChild as XShape, this);
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
            handleChildren();
        }
        #endregion

        #region XEventListener

        protected override void disposing(EventObject Source)
        {
            if (Source.Source == DrawPage)
            {
                refreshPagePropertiesTimer.Stop();
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
            ((IDisposable)this).Dispose();
        }

        #endregion

        #region XPropertyChangeListener

        protected override void propertyChange(PropertyChangeEvent evt)
        {
            if (evt.Source == DrawPage)
            {
                System.Diagnostics.Debug.WriteLine(evt.PropertyName + " changed to " + evt.NewValue.Value.ToString() + ", further: " + evt.Further);
            }
        }

        #endregion

        #region XVetoableChangeListener

        protected override void vetoableChange(PropertyChangeEvent aEvent)
        {
            if (aEvent.Source == DrawPage)
            {
                System.Diagnostics.Debug.WriteLine(aEvent.PropertyName + " vetoable changed to " + aEvent.NewValue.Value.ToString() + ", further: " + aEvent.Further);
            }
        }

        #endregion

        #region XPropertiesChangeListener

        protected override void propertiesChange(PropertyChangeEvent[] aEvent)
        {
            foreach (PropertyChangeEvent evt in aEvent)
            {
                if (evt.Source == DrawPage)
                {
                    System.Diagnostics.Debug.WriteLine(evt.PropertyName + " are changed to " + evt.NewValue.Value.ToString() + ", further: " + evt.Further);
                }
            }
        }

        #endregion

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