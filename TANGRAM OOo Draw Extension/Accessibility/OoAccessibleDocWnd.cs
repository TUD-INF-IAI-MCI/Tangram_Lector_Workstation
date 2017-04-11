using System;
using System.Collections.Generic;
using tud.mci.tangram.controller;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.view;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace tud.mci.tangram.Accessibility
{
    /// <summary>
    /// Encapsulates important properties of a document and its surrounding window elements.
    /// </summary>
    public partial class OoAccessibleDocWnd : IUpdateable, IDisposable, IDisposingObserver
    {
        #region Member

        /// <summary>
        /// The lock object for a synchronized object access of this window.
        /// </summary>
        public readonly Object SynchLock = new Object();

        /// <summary>
        /// The main window of the document [XAccessible]. 
        /// </summary>
        public Object MainWindow_anonymouse { get { return MainWindow; } }
        /// <summary>
        /// The main window of the document. 
        /// </summary>
        internal readonly XAccessible MainWindow;
        /// <summary>
        /// The document component containing the document objects [XAccessible].
        /// </summary>
        public Object Document_anonymouse { get { return Document; } }
        /// <summary>
        /// The document component containing the document objects.
        /// </summary>
        internal XAccessible Document { get; private set; }
        /// <summary>
        /// the parent window/frame containing the document and is child of the main window [XAccessible].
        /// </summary>
        public Object DocumentWindow_anonymouse { get { return DocumentWindow; } }
        /// <summary>
        /// the parent window/frame containing the document and is child of the main window.
        /// </summary>
        internal XAccessible DocumentWindow { get; private set; }
        /// <summary>
        /// The window handle of the main window
        /// </summary>
        public IntPtr Whnd { get; private set; }
        /// <summary>
        /// The process ID of the current OpenOffice instance
        /// </summary>
        public readonly int ProcessId = OoUtils.GetOoProcessID();

        /// <summary>
        /// The OoAccComponent for the documents. allows of collision testing etc.
        /// </summary>
        public OoAccComponent DocumentComponent
        {
            get
            {
                try
                {
                    var doc = new OoAccComponent(Document as XAccessibleComponent);
                    if (doc.Role == AccessibleRole.INVALID)
                    {
                        Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] document component invalid");

                        if (MainWindow != null)
                        {
                            setDocumentAndDocumentWindow(MainWindow);
                            doc = new OoAccComponent(Document as XAccessibleComponent);
                            if (doc.Role == AccessibleRole.INVALID)
                            {
                                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] document component invalid even after reset");
                            }
                        }
                        else
                        {
                            // document seems to be invalid ore already disposed
                            Dispose();
                        }
                    }
                    else if (doc.Role == AccessibleRole.DISPOSED)
                    {
                        Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] document component seems to be already disposed");
                        Dispose();
                    }
                    return doc;
                }
                catch (unoidl.com.sun.star.lang.DisposedException)
                {
                    Dispose();
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the cached title. This is not the current title. 
        /// it is updated by the Title.get function.
        /// Do not use this for getting the current title but for not stressing the API 
        /// - e.g. in case of logging or debugging.
        /// </summary>
        /// <value>
        /// The cached title.
        /// </value>
        public string CachedTitle { get; private set; }
        /// <summary>
        /// The title of the document
        /// </summary>
        public String Title
        {
            get
            {
                string t = getTitle();
                if (String.IsNullOrEmpty(t))
                {
                    t = CachedTitle;
                }
                else
                {
                    CachedTitle = t;
                }
                return t;
            }
        }

        XDrawPagesSupplier _dps = null;
        /// <summary>
        /// The corresponding DOM object for the DRAW document to create pages [XDrawPagesSupplier]
        /// </summary>
        public Object DrawPageSupplier
        {
            get
            {
                if (_dps == null) getXDrawPageSupplier();
                return _dps;
            }
            set
            {
                _dps = value as XDrawPagesSupplier;
                _lastController = null;
                _lastController = Controller as XController;

                // set the meta data access
                MetaData = new DrawDocMetaDataSet(_dps as unoidl.com.sun.star.document.XDocumentPropertiesSupplier);
            }
        }
        /// <summary>
        /// The corresponding OodrawPagesObserver for this XDrawPagesSupplier
        /// </summary>
        public OoDrawPagesObserver DrawPagesObs;

        #region EventListener Forwarder
        readonly XWindowListener2_Forwarder xWindowListener = new XWindowListener2_Forwarder();
        readonly XSelectionListener_Forwarder xSelectionListener = new XSelectionListener_Forwarder();
        readonly XAccessibleEventListener_Forwarder xAccessibleListener = new XAccessibleEventListener_Forwarder();
        #endregion

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccessibleDocWnd"/> struct.
        /// </summary>
        /// <param name="mainWnd">The main window (top Window).</param>
        /// <param name="doc">The DOCUMENT (role) object.</param>
        public OoAccessibleDocWnd(XAccessible mainWnd, XAccessible doc)
        {

            CachedTitle = "Unknown Document " + (this.GetHashCode() % 100);
            registerToEventForwarder();
            DrawPageSupplier = null;
            Whnd = IntPtr.Zero;
            MainWindow = mainWnd;
            Document = doc;
            setDocumentAndDocumentWindow(doc);

            if (MainWindow != null)
            {
                Whnd = OoUtils.GetWindowHandle(MainWindow as XSystemDependentWindowPeer);
            }

            // Main Window Events
            addWindowListeners(MainWindow);

            Task t = new Task(
                () =>
                {
                    int trys = 0;
                    while (!getXDrawPageSupplier() && ++trys < 10)
                    {
                        Thread.Sleep(1000);
                    }

                    //addSelectionListener(DrawPageSupplier);
                }
                );
            t.Start();
        }

        #endregion

        #region IUpdateable

        /// <summary>
        /// NOT IMPLEMENTED YET
        /// </summary>
        public void Update() { }

        #endregion

        #region private initialization functions

        void registerToEventForwarder()
        {
            if (xWindowListener != null)
            {
                xWindowListener.Disposing += new EventHandler<EventObjectForwarder>(disposing);
                xWindowListener.WindowDisabled += new EventHandler<EventObjectForwarder>(xWindowListener_WindowDisabled);
                xWindowListener.WindowEnabled += new EventHandler<EventObjectForwarder>(xWindowListener_WindowEnabled);
                xWindowListener.WindowHidden += new EventHandler<EventObjectForwarder>(xWindowListener_WindowHidden);
                xWindowListener.WindowMoved += new EventHandler<WindowEventForwarder>(xWindowListener_WindowMoved);
                xWindowListener.WindowResized += new EventHandler<WindowEventForwarder>(xWindowListener_WindowResized);
                xWindowListener.WindowShown += new EventHandler<EventObjectForwarder>(xWindowListener_WindowShown);
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] xWindowListener_Forwarder is null");
            }

            if (xAccessibleListener != null)
            {
                xAccessibleListener.Disposing += new EventHandler<EventObjectForwarder>(disposing);
                xAccessibleListener.NotifyEvent += new EventHandler<AccessibleEventObjectForwarder>(xAccessibleListener_NotifyEvent);
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] xAccessibleEvent_Forwarder is null");
            }

            if (xSelectionListener != null)
            {
                xSelectionListener.Disposing += new EventHandler<EventObjectForwarder>(disposing);
                xSelectionListener.SelectionChanged += new EventHandler<EventObjectForwarder>(xSelectionListener_SelectionChanged);
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] xWSelectionListener_Forwarder is null");
            }

        }

        private void setDocumentAndDocumentWindow(XAccessible doc)
        {
            if (doc == null)
            {
                doc = OoAccessibility.IsDrawWindow(MainWindow) as XAccessible;
            }
            if (doc != null && doc is XAccessibleContext)
            {
                var par = ((XAccessibleContext)doc).getAccessibleParent();
                while (par != null && !(par is XWindow2))
                {
                    ((XAccessibleContext)par).getAccessibleParent();
                }
                DocumentWindow = par as XAccessible;
            }
            else
            {
                DocumentWindow = null;
            }

            //Document Window Events
            addWindowListeners(DocumentWindow);

            //Document Events
            addDocumentListeners(Document);
        }

        private String getTitle()
        {
            String title = String.Empty;

            TimeLimitExecutor.WaitForExecuteWithTimeLimit(
                100,
                new Action(() =>
                {

                    if (DrawPagesObs != null)
                    {
                        title = DrawPagesObs.Title;
                    }

                    if (String.IsNullOrWhiteSpace(title))
                    {

                        String docT = String.Empty;
                        if (Document != null && Document.getAccessibleContext() != null)
                        {
                            docT = OoAccessibility.GetAccessibleName(Document);
                        }
                        if (!String.IsNullOrEmpty(docT))
                        {
                            docT = docT.Substring(0, Math.Max(0, docT.LastIndexOf('-'))).Trim();
                            title = docT;
                        }

                        String mainWndT = String.Empty;

                        if (MainWindow != null && MainWindow.getAccessibleContext() != null)
                        {
                            mainWndT = OoAccessibility.GetAccessibleName(MainWindow);
                        }

                        if (!String.IsNullOrEmpty(mainWndT)) //TODO: check why this is empty on CO2-Labels and unbenannt
                        {
                            mainWndT = mainWndT.Substring(0,
                                    mainWndT.LastIndexOf('.') > 0 ?
                                        mainWndT.LastIndexOf('.') : Math.Max(0, docT.LastIndexOf('-'))
                                    ).Trim();
                            title = mainWndT;
                        }

                        if (mainWndT.Equals(docT))
                            title = docT;
                        else
                            Logger.Instance.Log(LogPriority.IMPORTANT, this, "Cannot get correct title of the Oo document: Window title: '" + mainWndT + "' Document title: '" + docT + "'");
                    }

                }), "GetDocumentTitle");

            return title;
        }

        private void addWindowListeners(object p)
        {
            if (p != null && p is XWindow2)
            {
                try { ((XWindow2)p).addWindowListener(xWindowListener); }
                catch
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, "can't add window listener");
                    try { ((XWindow2)p).removeWindowListener(xWindowListener); }
                    catch
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(50);
                            ((XWindow2)p).removeWindowListener(xWindowListener);
                        }
                        catch (Exception e) { Logger.Instance.Log(LogPriority.ALWAYS, this, "can't remove window listener", e); }
                    }
                    try { ((XWindow2)p).addWindowListener(xWindowListener); }
                    catch (Exception e)
                    {
                        Logger.Instance.Log(LogPriority.ALWAYS, this, "finally can't add window listener", e);
                    }
                }
            }
        }

        private void addDocumentListeners(object dw)
        {
            if (xAccessibleListener != null)
            {
                if (dw != null)
                {
                    if (dw is XAccessibleEventBroadcaster)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, "Add Accessible event listener to: " + OoAccessibility.PrintAccessibleInfos(dw as XAccessible));
                        try
                        {
#if LIBRE
                            ((XAccessibleEventBroadcaster)dw).removeAccessibleEventListener(xAccessibleListener);
#else
                            ((XAccessibleEventBroadcaster)dw).removeEventListener(xAccessibleListener);
#endif
                        }
                        catch (unoidl.com.sun.star.lang.DisposedException)
                        {
                            this.Dispose();
                            return;
                        }
                        catch
                        {
                            try
                            {
#if LIBRE
                                ((XAccessibleEventBroadcaster)dw).removeAccessibleEventListener(xAccessibleListener);
#else
                                ((XAccessibleEventBroadcaster)dw).removeEventListener(xAccessibleListener);
#endif
                            }
                            catch (unoidl.com.sun.star.lang.DisposedException) { Dispose(); return; }
                            catch (Exception e) { Logger.Instance.Log(LogPriority.ALWAYS, this, "can't remove document listener", e); }
                        }

                        try
                        {
#if LIBRE
                            ((XAccessibleEventBroadcaster)dw).addAccessibleEventListener(xAccessibleListener);
#else
                            ((XAccessibleEventBroadcaster)dw).addEventListener(xAccessibleListener);
#endif
                        }
                        catch (unoidl.com.sun.star.lang.DisposedException) { Dispose(); return; }
                        catch (System.Exception e)
                        {
                            Logger.Instance.Log(LogPriority.ALWAYS, this, "can't add document listener", e);
                        }
                    }
                }
            }
        }

        private void addSelectionListener(Object dps)
        {
            /*
             * The document listener only throws selection events after 
             * editing the document the first time. Until this, in loaded 
             * drawing, no selections were reported through modify events.
             * The original selection event from the XSelectionSupplier 
             * has to been used for this. 
             */

            if (dps != null && dps is XModel2)
            {
                try
                {
                    var controller = ((XModel2)dps).getCurrentController();
                    if (controller != null && controller is XSelectionSupplier)
                    {
                        ((XSelectionSupplier)controller).addSelectionChangeListener(xSelectionListener);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] Can't add selection listener to document window: " + ex);
                }
            }
        }

        volatile bool _dpsSerach = false;
        /// <summary>
        /// Try to find the corresponding XDrawPageSupplier to the given Window.
        /// </summary>
        private bool getXDrawPageSupplier()
        {
            try
            {
                if (_dpsSerach)
                {
                    int tries = 0;
                    while (_dpsSerach && tries++ < 100 ) { Thread.Sleep(10); }
                    return tries < 100;
                }
            }
            catch (Exception)
            {
                return false;
            }

            _dpsSerach = true;
            try
            {
                List<XDrawPagesSupplier> dps = OoDrawUtils.GetDrawPageSuppliers(OO.GetDesktop());
                if (dps.Count > 0)
                {
                    foreach (var item in dps)
                    {
                        //string runtimeID = OoUtils.GetStringProperty(item, "RuntimeUID");
                        //string buildID = OoUtils.GetStringProperty(item, "BuildId");

                        XAccessible acc = OoDrawUtils.GetXAccessibleFromDrawPagesSupplier(item);
                        if (acc != null)
                        {
                            var root = OoAccessibility.GetRootPaneFromElement(acc);
                            if (root != null)
                            {
                                if (root.Equals(MainWindow))
                                {
                                    Logger.Instance.Log(LogPriority.DEBUG, this, "XDrawPageSuppliere found!!");
                                    DrawPageSupplier = item;

                                    //prepareForSelection(item); // Bad hack for getting accessible selection events

                                    DrawPagesObs = new OoDrawPagesObserver(DrawPageSupplier as XDrawPagesSupplier, DocumentComponent, this);
                                    return true;
                                }
                                else
                                {
                                    Logger.Instance.Log(LogPriority.OFTEN, this, "[ERROR] - Can't find root element from DrawPagesSupplier (root is not MainWindow)");
                                }
                            }
                            else
                            {
                                Logger.Instance.Log(LogPriority.OFTEN, this, "[ERROR] - Can't find root element from DrawPagesSupplier (root is null)");
                            }
                        }
                        else
                        {
                            Logger.Instance.Log(LogPriority.OFTEN, this, "[ERROR] - Can't find XAccessible for DrawPagesSupplier (root is null)");
                        }
                    }
                }
                Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] cannot find XDrawPageSuppliere for window");
                return false;
            }
            finally
            {
                _dpsSerach = false;
            }
        }

        /// <summary>
        /// Hack to activate the accessibility events for selection.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <remarks>this function is time limited.</remarks>
        private void prepareForSelection(XDrawPagesSupplier item)
        {
            // adding and deleting one shape to the document so the selections 
            // will be activated for notify events of the XAccessibleEventBroadcaster
            // FIXME: remove this hack - if the event works properly

            if (item != null)
            {

                TimeLimitExecutor.ExecuteWithTimeLimit(
                        () =>
                        {
                            var shape = OoDrawUtils.CreateEllipseShape(item, 0, 0, 3000, 3000);
                            var page = OoDrawUtils.GetCurrentPage(item);
                            if (page != null && shape != null)
                            {
                                OoDrawUtils.AddShapeToDrawPage(shape, page);
                                OoDrawUtils.RemoveShapeFromDrawPage(shape, page);
                            }
                        }
                    );
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] Cannot add shapes for activating selection events");
            }
        }

        #endregion

        #region EventListeners

        void disposing(object sender, EventObjectForwarder e) { disposing(e.E); }

        #region XWindowListener2

        void xWindowListener_WindowShown(object sender, EventObjectForwarder e) { windowShown(e.E); }
        void xWindowListener_WindowResized(object sender, WindowEventForwarder e) { windowResized(e.E); }
        void xWindowListener_WindowMoved(object sender, WindowEventForwarder e) { windowMoved(e.E); }
        void xWindowListener_WindowHidden(object sender, EventObjectForwarder e) { windowHidden(e.E); }
        void xWindowListener_WindowEnabled(object sender, EventObjectForwarder e) { windowEnabled(e.E); }
        void xWindowListener_WindowDisabled(object sender, EventObjectForwarder e) { windowDisabled(e.E); }

        /// <summary>
        /// Gets the last known state of the window.
        /// </summary>
        /// <value>
        /// The last known state of the  window.
        /// </value>
        public WindowEventType LastKnownWindowState { get; private set; }

        void windowDisabled(unoidl.com.sun.star.lang.EventObject e) { LastKnownWindowState = WindowEventType.DEACTIVATED; fireWindowEvent(WindowEventType.DEACTIVATED, e); }
        void windowEnabled(unoidl.com.sun.star.lang.EventObject e) { LastKnownWindowState = WindowEventType.ACTIVATED; fireWindowEvent(WindowEventType.ACTIVATED, e); }
        void windowHidden(unoidl.com.sun.star.lang.EventObject e) { LastKnownWindowState = WindowEventType.MINIMIZED; fireWindowEvent(WindowEventType.MINIMIZED, e); }
        void windowMoved(WindowEvent e) { fireWindowEvent(WindowEventType.CHANGED, e); }
        void windowResized(WindowEvent e) { fireWindowEvent(WindowEventType.CHANGED, e); }
        void windowShown(unoidl.com.sun.star.lang.EventObject e) { LastKnownWindowState = WindowEventType.OPENED; fireWindowEvent(WindowEventType.OPENED, e); }

        void disposing(unoidl.com.sun.star.lang.EventObject Source) { LastKnownWindowState = WindowEventType.CLOSED; fireWindowEvent(WindowEventType.CLOSED, Source); }

        #endregion

        #region XAccessibleEventListener

        void xAccessibleListener_NotifyEvent(object sender, AccessibleEventObjectForwarder e) { notifyEvent(e.E); }

        void notifyEvent(AccessibleEventObject aEvent)
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "[ACCESSIBLE] accessible event in document " + (aEvent != null ? OoAccessibility.GetAccessibleEventIdFromShort(aEvent.EventId).ToString() : ""));

            //var eventType = OoAccessibility.GetAccessibleEventIdFromShort(aEvent.EventId);

            if (aEvent != null && aEvent.EventId == (short)AccessibleEventId.INVALIDATE_ALL_CHILDREN)
            {
                resetActivePagesObsCache();
            }


            /* you can not rely on the selection events. 
             * They will only be thrown after editing the document. So selections in 
             * all loaded document will not been reported ever! 
             * Use the XSelectionProveider of the XControllen instead!
             */
            fireAccessibleEvent(aEvent);
        }

        #endregion

        #region XSelectionChangeListener

        void xSelectionListener_SelectionChanged(object sender, EventObjectForwarder e)
        {
            selectionChanged(e.E);
        }

        /// <summary>
        /// Selections the changed.
        /// Prepare this event that it looks like the accessible modify selection event.
        /// The selection supplier returns the real XShapes!!! and not the XAccessible objects :-/
        /// </summary>
        /// <param name="aEvent">A event.</param>
        void selectionChanged(EventObject aEvent)
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "[SELECTION] selection changed event in document " + Title);
            //System.Diagnostics.Debug.WriteLine("## AccessibleDocWnd --> selction changed (Selection Listener)");
            if (aEvent != null && aEvent.Source != null)
            {
                fireSelectionEvent(new OoSelectionChandedEventArgs(aEvent.Source));
            }
        }

        #endregion

        #endregion

        #region Events

        /// <summary>
        /// Event forwarder for events from the two observed windows (Main window and Document window)
        /// </summary>
        public event EventHandler<OoAccessibleDocWindowEventArgs> WindowEvent;
        /// <summary>
        /// Forwarder for accessibility events of the observed Document. 
        /// </summary>
        public event EventHandler<OoAccessibleDocAccessibleEventArgs> AccessibleEvent;
        /// <summary>
        /// Occurs when the selection changed.
        /// </summary>
        public event EventHandler<OoSelectionChandedEventArgs> SelectionEvent;

        void fireWindowEvent(WindowEventType type, EventObject e)
        {
            if (WindowEvent != null)
            {
                try
                {
                    WindowEvent.Invoke(this, new OoAccessibleDocWindowEventArgs(type, e));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire window event", ex); }
            }
        }

        void fireAccessibleEvent(AccessibleEventObject aEvents)
        {
            if (AccessibleEvent != null)
            {
                try
                {
                    AccessibleEvent.Invoke(this, new OoAccessibleDocAccessibleEventArgs(aEvents));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire accessibility event", ex); }
            }
        }

        void fireSelectionEvent(OoSelectionChandedEventArgs aEvents)
        {
            if (SelectionEvent != null)
            {
                try
                {
                    SelectionEvent.Invoke(this, aEvents);
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire selection event", ex); }
            }
        }

        #endregion

        #region Override

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return "OoAccessibleDocWnd [" + GetHashCode() + "] "
                + "Title:'" + CachedTitle
                + "' wHndl:" + Whnd
                + " MainWnd hash:" + (MainWindow != null ? MainWindow.GetHashCode().ToString() : "null")
                + " DocWnd hash:" + (DocumentWindow != null ? DocumentWindow.GetHashCode().ToString() : "null")
                + " Doc hash:" + (Document != null ? Document.GetHashCode().ToString() : "null");
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return Whnd != IntPtr.Zero ? (int)Whnd : base.GetHashCode();
        }

        #endregion

        #region Shape Observer Get

        /// <summary>
        /// Gets the registered shape observer to the given component if it is already registered.
        /// </summary>
        /// <param name="xShape">The shape.</param>
        /// <param name="page">The page to look for.</param>
        /// <returns>
        /// an accessible component if it is already registered
        /// </returns>
        internal OoShapeObserver GetRegisteredShapeObserver(XShape xShape, OoDrawPageObserver page)
        {
            OoShapeObserver shape = DrawPagesObs.GetRegisteredShapeObserver(xShape, page);
            return shape;
        }

        /// <summary>
        /// Gets the registered shape observer to the accessible component if it is already registered.
        /// </summary>
        /// <param name="xAccessibleComponent">The x accessible component.</param>
        /// <returns>an accessible component if it is already registered</returns>
        internal OoShapeObserver GetRegisteredShapeObserver(XAccessibleComponent xAccessibleComponent)
        {
            OoShapeObserver shape = DrawPagesObs.GetRegisteredShapeObserver(xAccessibleComponent);
            return shape;
        }

        /// <summary>
        /// Gets the registered shape observer to the accessible component if it is already registered.
        /// </summary>
        /// <param name="xAccessibleComponent">The x accessible component.</param>
        /// <returns>an accessible component if it is already registered</returns>
        public OoShapeObserver GetRegisteredShapeObserver(OoAccComponent accessibleComponent)
        {
            return GetRegisteredShapeObserver(accessibleComponent.AccComp);
        }

        /// <summary>
        /// Gets the registered shape observer to the accessible component if it is already registered.
        /// </summary>
        /// <param name="shape">The shape [OoAccComponent], [XAccessibleComponent] or [XShape].</param>
        /// <param name="page">The page the shape should be.</param>
        /// <returns>
        /// an accessible component if it is already registered
        /// </returns>
        public OoShapeObserver GetRegisteredShapeObserver(Object shape, OoDrawPageObserver page = null)
        {
            if (shape != null)
            {
                if (shape is OoAccComponent) return GetRegisteredShapeObserver(shape as OoAccComponent);
                else if (shape is XAccessibleComponent) return GetRegisteredShapeObserver(shape as XAccessibleComponent);
                else if (shape is XShape && page != null) return GetRegisteredShapeObserver(shape as XShape, page); 
            }
            return null;
        }

        #endregion

        #region Page Get

        /// <summary>
        /// Return the current active page.
        /// Tries to call the active page observer only once
        /// </summary>
        /// <returns></returns>
        public OoDrawPageObserver GetActivePage()
        {
            int i = 0;
            while (_runing && ++i < 10)
            {
                Thread.Sleep(50);
            }

            if (!_runing)
            {
                OoDrawPageObserver pObs = getActivePageObserver();
                if (pObs != null)
                {
                    return pObs;
                }
            }
            return null;
        }

        int _cachedPageCount = 0;
        /// <summary>
        /// Gets the page count.
        /// </summary>
        /// <returns>The number of available pages or 0</returns>
        public int GetPageCount()
        {
            int pageCount = _cachedPageCount;
            var success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(100, () =>
            {
                if (DrawPageSupplier != null && DrawPageSupplier is XDrawPagesSupplier)
                {
                    var pages = ((XDrawPagesSupplier)DrawPageSupplier).getDrawPages();
                    if (pages != null) pageCount = pages.getCount();
                }
            }, "GetPageCount");

            _cachedPageCount = pageCount;
            return pageCount;
        }

        void resetActivePagesObsCache() { _lastPageObs = null; }
        DateTime _lastPageObsObserving;
        TimeSpan _maxCachePageTime = new TimeSpan(0, 1, 0);
        volatile bool _runing = false;
        OoDrawPageObserver _lastPageObs = null;
        /// <summary>
        /// Return the observer of the current active page.
        /// </summary>
        /// <returns></returns>
        private OoDrawPageObserver getActivePageObserver()
        {
            if (_lastPageObs != null && DateTime.Now - _lastPageObsObserving < _maxCachePageTime) return _lastPageObs;

            // if this is hanging or called multiple times
            int tries = 0;
            while (_runing && tries++ < 10) { Thread.Sleep(10); }
            if (_runing) { return _lastPageObs; }

            _runing = true;
            OoDrawPageObserver pageObs = null;
            try
            {
                if (DrawPagesObs != null && DrawPagesObs.HasDrawPages()
                    && DrawPageSupplier != null && DrawPageSupplier is XModel2
                    )
                {
                    var dps = DrawPagesObs.GetDrawPages();
                    if (dps.Count > 1) // if only one page is known - take it
                    {
                        // get the controller of the Draw application
                        XController contr = Controller as XController;

                        int pid = GetCurrentActivePageId(contr);

                        // find the OoDrawPageObserver for the correct page number
                        if (pid > 0)
                        {
                            foreach (var page in dps)
                            {
                                if (page != null && page.DrawPage != null)
                                {
                                    int pNum = util.OoUtils.GetIntProperty(page.DrawPage, "Number");
                                    if (pNum == pid)
                                    {
                                        pageObs = page;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        pageObs = dps.Count > 0 ? dps[0] : null;
                    }
                }

                _lastPageObs = pageObs;
                _lastPageObsObserving = DateTime.Now;
                return pageObs;
            }
            finally
            {
                _runing = false;
            }
        }

        #region get Page ID

        /// <summary>
        /// The cached current active page id. This property will be updated by 
        /// the <see cref="GetCurrentActivePageId"/> function. 
        /// Use this for getting access to the property without stressing the API.
        /// </summary>
        internal int CachedCurrentPid = -1;
        private bool _gettingPageID = false;
        /// <summary>
        /// Try to get the page number of the current active page
        /// </summary>
        /// <param name="contr"></param>
        /// <param name="pid"></param>
        /// <returns>The ID (page number) of the current active page.</returns>
        /// <remarks>Parts of this function are time limited to 100 ms.</remarks>
        internal int GetCurrentActivePageId(XController contr)
        {
            if (_gettingPageID)
                return CachedCurrentPid;
            int pid = -1;
            try
            {
                //lock (SynchLock)
                //{
                bool success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(100,
                         () =>
                         {
                             try
                             {
                                 if (contr != null && contr is XDrawView)
                                 {
                                     // get the current page
                                     var page = ((XDrawView)contr).getCurrentPage();
                                     // get the number
                                     pid = util.OoUtils.GetIntProperty(page, "Number");
                                 }
                             }
                             catch (DisposedException)
                             {
                                 Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] Controller seems to be already disposed");
                                 this.Dispose();
                             }
                             catch (System.Exception ex)
                             {
                                 Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] cant use the controller of document window: "/*, ex*/);
                             }
                         }
                         , "useControllerOfModel"
                         );
                if (!success) return CachedCurrentPid;
                //}
                CachedCurrentPid = pid;
                return pid;
            }
            finally { _gettingPageID = false; }
        }

        #endregion

        #endregion

        #region Controller

        XController _lastController = null;
        volatile bool _gettingController = false;
        /// <summary>
        /// Tries the get the controller of the document model [XController].
        /// </summary>
        /// <param name="model">The model.</param>
        /// <returns>The controller for this whole DRAW document.</returns>
        /// <remarks>Parts of this function are time limited to 200 ms.</remarks>
        public Object Controller
        {
            get
            {
                if (_lastController != null) return _lastController;
                if (_gettingController) return _lastController;

                XController contr = null;
                //lock (_controlerLock)
                {
                    _gettingController = true;
                    XModel2 model = DrawPageSupplier as XModel2;
                    if (model != null)
                    {
                        TimeLimitExecutor.WaitForExecuteWithTimeLimit(200,
                                    () =>
                                    {
                                        try
                                        {
                                            if (model != null) contr = model.getCurrentController();
                                        }
                                        catch (DisposedException ex)
                                        {
                                            Dispose();
                                            Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] controller for document window already disposed: " + ex);
                                        }
                                        catch (System.Exception ex)
                                        {
                                            Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] can't get controller for document window: " + ex);
                                        }
                                    }
                                    , "getControllerOfModel"
                                    );
                    }
                    _lastController = contr;

                    if (_lastController is XSelectionSupplier)
                    {
                        try
                        {
                            XSelectionSupplier supl = _lastController as XSelectionSupplier;
                            if (supl != null)
                            {
                                try { supl.removeSelectionChangeListener(xSelectionListener); }
                                catch { }
                                try
                                {
                                    supl.addSelectionChangeListener(xSelectionListener);
                                }
                                catch
                                {
                                    System.Threading.Thread.Sleep(5);
                                    try
                                    {
                                        supl.addSelectionChangeListener(xSelectionListener);
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Could not add selection listener to documents' controler");
                        }
                    }

                    _gettingController = false;
                }
                return contr;
            }
        }

        #endregion

        #region Window State Checks

        #region user32.dll Invoke

        [DllImport("user32.dll")]
        static extern int IsIconic(int hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowInfo(IntPtr hWnd, out WINDOWINFO pwi);

        [DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        static extern bool IsWindow(IntPtr hWnd);

        #endregion

        /// <summary>
        /// Determines whether this window is visible.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this window is visible; otherwise, <c>false</c>.
        /// </returns>
        public bool IsVisible()
        {
            return (IsWindow(this.Whnd) && IsIconic((int)this.Whnd) == 0);
        }

        /// <summary>
        /// Determines whether this window is active.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this window is active; otherwise, <c>false</c>.
        /// </returns>
        public bool IsActive()
        {
            if (IsVisible())
            {
                WINDOWINFO pwi;
                // dwWindowStatus = 0x0001 == ACTIVE
                if (GetWindowInfo(this.Whnd, out pwi))
                {
                    return pwi.dwWindowStatus == 0x0001;
                }
            }
            return false;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INFO]" + Whnd + " window closed and Observer disposing");
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

    #region Structs

    [StructLayout(LayoutKind.Sequential)]
    internal struct WINDOWINFO
    {
        public uint cbSize;
        public RECT rcWindow;
        public RECT rcClient;
        public uint dwStyle;
        public uint dwExStyle;
        public uint dwWindowStatus;
        public uint cxWindowBorders;
        public uint cyWindowBorders;
        public ushort atomWindowType;
        public ushort wCreatorVersion;

        public WINDOWINFO(Boolean? filler)
            : this()   // Allows automatic initialization of "cbSize" with "new WINDOWINFO(null/true/false)".
        {
            cbSize = (UInt32)(Marshal.SizeOf(typeof(WINDOWINFO)));
        }

    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT
    {
        public int Left, Top, Right, Bottom;

        public RECT(int left, int top, int right, int bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public RECT(System.Drawing.Rectangle r) : this(r.Left, r.Top, r.Right, r.Bottom) { }

        public int X
        {
            get { return Left; }
            set { Right -= (Left - value); Left = value; }
        }

        public int Y
        {
            get { return Top; }
            set { Bottom -= (Top - value); Top = value; }
        }

        public int Height
        {
            get { return Bottom - Top; }
            set { Bottom = value + Top; }
        }

        public int Width
        {
            get { return Right - Left; }
            set { Right = value + Left; }
        }

        public System.Drawing.Point Location
        {
            get { return new System.Drawing.Point(Left, Top); }
            set { X = value.X; Y = value.Y; }
        }

        public System.Drawing.Size Size
        {
            get { return new System.Drawing.Size(Width, Height); }
            set { Width = value.Width; Height = value.Height; }
        }

        public static implicit operator System.Drawing.Rectangle(RECT r)
        {
            return new System.Drawing.Rectangle(r.Left, r.Top, r.Width, r.Height);
        }

        public static implicit operator RECT(System.Drawing.Rectangle r)
        {
            return new RECT(r);
        }

        public static bool operator ==(RECT r1, RECT r2)
        {
            return r1.Equals(r2);
        }

        public static bool operator !=(RECT r1, RECT r2)
        {
            return !r1.Equals(r2);
        }

        public bool Equals(RECT r)
        {
            return r.Left == Left && r.Top == Top && r.Right == Right && r.Bottom == Bottom;
        }

        public override bool Equals(object obj)
        {
            if (obj is RECT)
                return Equals((RECT)obj);
            else if (obj is System.Drawing.Rectangle)
                return Equals(new RECT((System.Drawing.Rectangle)obj));
            return false;
        }

        public override int GetHashCode()
        {
            return ((System.Drawing.Rectangle)this).GetHashCode();
        }

        public override string ToString()
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{{Left={0},Top={1},Right={2},Bottom={3}}}", Left, Top, Right, Bottom);
        }
    }

    #endregion

    #region Event Arg Definition

    public class OoAccessibleDocWindowEventArgs : EventArgs
    {
        /// <summary>
        /// The type of the window event
        /// </summary>
        public readonly WindowEventType Type;
        /// <summary>
        /// The original native event object (Can be EventObject or WindowEvent)
        /// </summary>
        public readonly EventObject E;

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccessibleDocWindowEventArgs"/> class.
        /// </summary>
        /// <param name="type">The original event type.</param>
        /// <param name="e">The native event object.</param>
        public OoAccessibleDocWindowEventArgs(WindowEventType type, EventObject e)
        {
            Type = type;
            E = e;
        }

    }

    public class OoAccessibleDocAccessibleEventArgs : EventArgs
    {
        public readonly AccessibleEventObject E;

        public OoAccessibleDocAccessibleEventArgs(AccessibleEventObject aEvent)
        {
            E = aEvent;
        }
    }

    #endregion

    #region Enums

    /// <summary>
    /// Define the type of window event
    /// </summary>
    [Flags]
    public enum WindowEventType
    {
        /// <summary>
        /// no defined reason
        /// </summary>
        NONE = 0,
        /// <summary>
        /// unknown reason
        /// </summary>
        UNKNOWN = 1,
        /// <summary>
        /// window opened
        /// </summary>
        OPENED = 2,
        /// <summary>
        /// window activated
        /// </summary>
        ACTIVATED = 4,
        /// <summary>
        /// window deactivated
        /// </summary>
        DEACTIVATED = 8,
        /// <summary>
        /// window minimized
        /// </summary>
        MINIMIZED = 16,
        /// <summary>
        /// window closed
        /// </summary>
        CLOSED = 32,
        /// <summary>
        /// Some properties have changed
        /// </summary>
        CHANGED = 64,
    }

    #endregion
}