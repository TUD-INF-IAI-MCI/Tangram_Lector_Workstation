using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.view;
using tud.mci.tangram.models;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Singleton class that observes openOffice and looking for Draw/Impress documents, 
    /// register changes and provides a central place for suppling them to outer project parts.
    /// </summary>
    public class OoDrawAccessibilityObserver : IResetable, IDisposable
    {
        #region Members
        /// <summary>
        /// Key is the Accessible Document. 
        /// List of all known draw page suppliers Windows (DRAW and IMPRESS documents)
        /// </summary>
        readonly ConcurrentDictionary<XAccessible, OoAccessibleDocWnd> drawDocs = new ConcurrentDictionary<XAccessible, OoAccessibleDocWnd>();
        /// <summary>
        /// Key is the Accessible Window directly surrounding the Document. 
        /// List of all known draw page suppliers Windows (DRAW and IMPRESS documents)
        /// </summary>
        readonly ConcurrentDictionary<XAccessible, OoAccessibleDocWnd> drawDocWnds = new ConcurrentDictionary<XAccessible, OoAccessibleDocWnd>();
        /// <summary>
        /// Key is the Accessible Application Window of the Document. 
        /// List of all known draw page suppliers Windows (DRAW and IMPRESS documents)
        /// </summary>
        readonly ConcurrentDictionary<XAccessible, OoAccessibleDocWnd> drawWnds = new ConcurrentDictionary<XAccessible, OoAccessibleDocWnd>();
        private static volatile bool initalized = false;
        OoTopWindowObserver tpwndo = OoTopWindowObserver.Instance;

        #endregion

        private static OoDrawAccessibilityObserver instance = new OoDrawAccessibilityObserver();

        private OoDrawAccessibilityObserver()
        {
            OoSelectionObserver.Instance.SelectionChanged += new EventHandler<OoSelectionChandedEventArgs>(Instance_SelectionChanged);
            //drawPgSuppl = OoDrawUtils.GetDrawPageSuppliers(OO.GetDesktop());

            OnShapeBoundRectChange = new OoShapeObserver.BoundRectChangeEventHandler(ShapeBoundRectChangeHandler);
            OnViewOrZoomChange = new OoDrawPagesObserver.ViewOrZoomChangeEventHandler(ShapeBoundRectChangeHandler);
            new Task(initalize).Start();
        }

        private void initalize()
        {
            #region OoTopWindowObserver initialization
            try
            {
                tpwndo.WindowActivated += new EventHandler<OoEventArgs>(tpwndo_WindowActivated);
                //tpwndo.Disposing += new EventHandler<OoEventArgs>(tpwndo_Disposing);
                tpwndo.WindowClosed += new EventHandler<OoEventArgs>(tpwndo_WindowClosed);
                //tpwndo.WindowClosing += new EventHandler<OoEventArgs>(tpwndo_WindowClosing);
                //tpwndo.WindowDeactivated += new EventHandler<OoEventArgs>(tpwndo_WindowDeactivated);
                tpwndo.WindowMinimized += new EventHandler<OoEventArgs>(tpwndo_WindowMinimized);
                tpwndo.WindowNormalized += new EventHandler<OoEventArgs>(tpwndo_WindowNormalized);
                tpwndo.WindowOpened += new EventHandler<OoEventArgs>(tpwndo_WindowOpened);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.ALWAYS, this, "Accessibility Observer TopWindowObserever initialization failed", ex);
            }

            #endregion

            #region TopWindow initialization

            var tpwnds = OoAccessibility.GetAllTopWindows();
            foreach (var item in tpwnds)
            {
                tpwndo_WindowOpened(this, new OoEventArgs(item));
            }

            initalized = true;

            var actTpwnd = OoAccessibility.GetActiveTopWindow();
            if (actTpwnd != null)
                tpwndo_WindowActivated(this, new OoEventArgs(actTpwnd));

            #endregion
        }

        /// <summary>
        /// Returns the Singleton instance.
        /// </summary>
        /// <value>The instance.</value>
        public static OoDrawAccessibilityObserver Instance
        {
            get
            {
                int i = 0;
                while (!initalized && ++i < 20) { Thread.Sleep(10); }

                return instance;
            }
        }

        /// <summary>
        /// Returns a list off all known draw/impress docs.
        /// </summary>
        /// <returns>list off all known draw/impress docs</returns>
        public List<OoAccessibleDocWnd> GetDrawDocs() { return drawWnds.Values.ToList(); }

        #region OoTopWindowObserver Events

        void tpwndo_WindowOpened(object sender, OoEventArgs e)
        {
            //check if window is a draw doc
            if (e != null && e.Source != null && e.Source is XAccessible)
            {
                var doc = OoAccessibility.IsDrawWindow(e.Source as XAccessible);
                if (doc != null && !(doc is bool))
                {
                    registerNewDrawWindow(e.Source as XAccessible, doc);
                }
            }
        }

        void tpwndo_WindowNormalized(object sender, OoEventArgs e)
        {
            if (e != null && e.Source != null && e.Source is XAccessible)
            {
                OoAccessibleDocWnd doc = getCorrespondingAccessibleDocForXaccessible(e.Source as XAccessible);

                if (doc != null && doc.MainWindow != null)
                {
                    fireDrawWindowPropertyChangeEvent(doc);
                }
            }
        }

        void tpwndo_WindowMinimized(object sender, OoEventArgs e)
        {
            if (e != null && e.Source != null && e.Source is XAccessible)
            {
                OoAccessibleDocWnd doc = getCorrespondingAccessibleDocForXaccessible(e.Source as XAccessible);

                if (doc != null && doc.MainWindow != null)
                {
                    fireDrawWindowMinimizedEvent(doc);
                }
            }
        }

        void tpwndo_WindowDeactivated(object sender, OoEventArgs e) { }

        void tpwndo_WindowClosing(object sender, OoEventArgs e) { }

        void tpwndo_WindowClosed(object sender, OoEventArgs e)
        {
            if (drawWnds.ContainsKey(e.Source as XAccessible))
            {
                //System.Diagnostics.Debug.WriteLine("\t\tknown drawdoc closed");
                OoAccessibleDocWnd doc;
                drawWnds.TryRemove(e.Source as XAccessible, out doc);
                OoAccessibleDocWnd doc2;
                if (doc.DocumentWindow != null && drawDocWnds.ContainsKey(doc.DocumentWindow))
                {
                    drawDocWnds.TryRemove(doc.DocumentWindow, out doc2);
                }
                if (doc.Document != null && drawDocs.ContainsKey(doc.Document))
                {
                    drawDocs.TryRemove(doc.Document, out doc2);
                }

                // release the document
                doc.Dispose();
                fireDrawWindowClosedEvent(doc);
            }
            tpwndo = OoTopWindowObserver.Instance;
        }

        void tpwndo_Disposing(object sender, OoEventArgs e) { }

        void tpwndo_WindowActivated(object sender, OoEventArgs e)
        {
            if (e != null && e.Source != null && e.Source is XAccessible)
            {
                OoAccessibleDocWnd doc = getCorrespondingAccessibleDocForXaccessible(e.Source as XAccessible);

                if (doc != null && doc.MainWindow != null)
                {
                    fireDrawWindowActivatedEvent(doc);
                }
                else
                {
                    var docObj = OoAccessibility.IsDrawWindow(e.Source as XAccessible);
                    if (docObj != null && !(docObj is bool))
                    {
                        registerNewDrawWindow(e.Source as XAccessible, docObj);
                    }
                }
            }
        }

        #region utils

        #region user32

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern bool IsChild(IntPtr hWndParent, IntPtr hwnd);

        #endregion


        /// <summary>
        /// try to gets the corresponding accessible doc for XAccessible.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        private OoAccessibleDocWnd getCorrespondingAccessibleDocForXaccessible(XAccessible key)
        {
            OoAccessibleDocWnd doc = null;
            if (key != null)
            {
                if (!drawDocs.TryGetValue(key, out doc) || doc == null)
                {
                    if (!drawWnds.TryGetValue(key, out doc) || doc == null)
                    {
                        drawDocWnds.TryGetValue(key, out doc);
                    }
                }
            }
            return doc;
        }

        /// <summary>
        /// Registers a new draw window.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="dw">The dw.</param>
        private void registerNewDrawWindow(XAccessible source, Object dw)
        {
            if (source != null && dw != null && !dw.Equals(false))
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "Draw window check: '" + OoAccessibility.GetAccessibleName(source));

                if (getCorrespondingAccessibleDocForXaccessible(source) == null)
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, "Register new Draw window: '" + OoAccessibility.GetAccessibleName(source));

                    OoAccessibleDocWnd doc = new OoAccessibleDocWnd(source, dw as XAccessible);

                    drawWnds[source] = doc; // add main window to list
                    if (doc.Document != null) { drawDocs[source] = doc; }
                    if (doc.DocumentWindow != null) { drawDocWnds[doc.DocumentWindow] = doc; }

                    //drawPgSuppl.Clear();
                    //drawPgSuppl.AddRange(OoDrawUtils.GetDrawPageSuppliers(OO.GetDesktop()));
                    //TODO: call the observers

                    addListeners(doc);
                    fireDrawWindowOpendEvent(doc);
                }
            }
        }

        #endregion

        #endregion

        #region Listeners

        private void addListeners(OoAccessibleDocWnd doc)
        {
            if (doc != null)
            {
                doc.AccessibleEvent += new EventHandler<OoAccessibleDocAccessibleEventArgs>(doc_AccessibleEvent);
                doc.WindowEvent += new EventHandler<OoAccessibleDocWindowEventArgs>(doc_WindowEvent);
                doc.ObserverDisposing += new EventHandler(doc_Disposing);
                doc.SelectionEvent += new EventHandler<OoSelectionChandedEventArgs>(doc_SelectionEvent);
            }
        }


        void doc_Disposing(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ATTENTION]: DrawDocument Observer is disposing");

            OoAccessibleDocWnd w = sender as OoAccessibleDocWnd;
            OoAccessibleDocWnd trash;
            if (w != null)
            {
                try
                {
                    //TODO: delete it from the lists
                    foreach (var item in drawDocs.Keys)
                    {
                        if (drawDocs[item] == w)
                        {
                            drawDocs.TryRemove(item, out trash);
                        }
                    }

                    foreach (var item in drawWnds.Keys)
                    {
                        if (drawWnds[item] == w)
                        {
                            drawWnds.TryRemove(item, out trash);
                        }
                    }

                    foreach (var item in drawDocWnds.Keys)
                    {
                        if (drawDocWnds[item] == w)
                        {
                            drawDocWnds.TryRemove(item, out trash);
                        }
                    }

                    w.AccessibleEvent -= new EventHandler<OoAccessibleDocAccessibleEventArgs>(doc_AccessibleEvent);
                    w.WindowEvent -= new EventHandler<OoAccessibleDocWindowEventArgs>(doc_WindowEvent);
                    w.ObserverDisposing -= new EventHandler(doc_Disposing);

                }
                catch { }
            }

        }

        void doc_WindowEvent(object sender, OoAccessibleDocWindowEventArgs e)
        {
            if (e != null && sender != null && sender is OoAccessibleDocWnd)
            {
                OoAccessibleDocWnd drawDoc = sender as OoAccessibleDocWnd;

                switch (e.Type)
                {
                    case WindowEventType.OPENED:
                        fireDrawWindowActivatedEvent(drawDoc);
                        break;
                    case WindowEventType.CLOSED:
                        break;
                    default:
                        fireDrawWindowPropertyChangeEvent(drawDoc);
                        break;
                }
            }
        }

        void doc_AccessibleEvent(object sender, OoAccessibleDocAccessibleEventArgs e)
        {
            if (sender != null && sender is OoAccessibleDocWnd && e != null && e.E != null)
            {
                tud.mci.tangram.Accessibility.AccessibleEventId id = OoAccessibility.GetAccessibleEventIdFromShort(e.E.EventId);
                handleAccessibleEvent(sender as OoAccessibleDocWnd, id, e.E);
            }
        }

        void doc_SelectionEvent(object sender, OoSelectionChandedEventArgs e)
        {
            if (sender is OoAccessibleDocWnd && e != null)
            {
                selectionChanged(sender as OoAccessibleDocWnd, new AccessibleEventObject());
            }
        }

        #endregion

        #region Accessible Events

        private void handleAccessibleEvent(OoAccessibleDocWnd doc, tud.mci.tangram.Accessibility.AccessibleEventId id, AccessibleEventObject aEvent)
        {
            System.Diagnostics.Debug.WriteLine("Accessible event from DrawDocWnd :'" + doc.Title + "' ID: " + id);

            switch (id)
            {
                //case tud.mci.tangram.Accessibility.AccessibleEventId.NONE:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.ACTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.ACTIVE_DESCENDANT_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.ACTIVE_DESCENDANT_CHANGED_NOFOCUS:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.BOUNDRECT_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.CARET_CHANGED:
                //    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.CHILD:
                    //handleAccessibleChildEvent(doc, aEvent.Source, aEvent.NewValue, aEvent.OldValue);
                    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.COLUMN_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.CONTENT_FLOWS_FROM_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.CONTENT_FLOWS_TO_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.CONTROLLED_BY_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.CONTROLLER_FOR_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.DESCRIPTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.HYPERTEXT_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.INVALIDATE_ALL_CHILDREN:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.LABEL_FOR_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.LABELED_BY_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.LISTBOX_ENTRY_COLLAPSED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.LISTBOX_ENTRY_EXPANDED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.MEMBER_OF_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.NAME_CHANGED:
                //    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.PAGE_CHANGED:
                    fireDrawWindowActivatedEvent(doc);
                    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.SECTION_CHANGED:
                //    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED:
                    selectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_ADD:
                    selectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_REMOVE:
                    selectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_WITHIN:
                    selectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.STATE_CHANGED:
                    stateChanged(doc, aEvent);
                    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.SUB_WINDOW_OF_RELATION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_CAPTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_COLUMN_DESCRIPTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_COLUMN_HEADER_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_MODEL_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_ROW_DESCRIPTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_ROW_HEADER_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TABLE_SUMMARY_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TEXT_ATTRIBUTE_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TEXT_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.TEXT_SELECTION_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.VALUE_CHANGED:
                //    break;
                //case tud.mci.tangram.Accessibility.AccessibleEventId.VISIBLE_DATA_CHANGED:
                //    break;
                default:
                    break;
            }
        }

        private void stateChanged(OoAccessibleDocWnd doc, AccessibleEventObject aEvent)
        {
            if (doc != null && aEvent != null)
            {
                if (aEvent.Source != null)
                {
                    // fire window activated when document gets mouse focus
                    var role = OoAccessibility.GetAccessibleRole(aEvent.Source as XAccessible);
                    if (role == Accessibility.AccessibleRole.DOCUMENT)
                    {
                        if (OoAccessibility.HasAccessibleState(aEvent.Source as XAccessible, Accessibility.AccessibleStateType.FOCUSED))
                        {
                            fireDrawWindowActivatedEvent(doc);
                        }
                    }
                }
            }
        }


        /// <summary>
        /// add a new selection event th the selection stack
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="aEvent"></param>
        private void selectionChanged(OoAccessibleDocWnd doc, AccessibleEventObject aEvent)
        {
            if (aEvent != null)
            {
                selectionStack.Push(new KeyValuePair<OoAccessibleDocWnd, AccessibleEventObject>(doc, aEvent));
                initSelectionHandlerThread();
            }
        }


        private void handleSelectionChanged(OoAccessibleDocWnd doc, AccessibleEventObject aEvent)
        {

            //FIXME: Do this only if the selection is requested. 
            // Kill the selection changed handler and let the event handler request for the selected elements

            // check the global selection supplier
            if (doc != null)
            {
                try
                {
                    var controller = doc.Controller;
                    if (controller != null && controller is XSelectionSupplier)
                    {
                        XShapes selectedShapes = OoSelectionObserver.GetSelection(controller as XSelectionSupplier) as XShapes;

                        OoDrawPagesObserver pagesObserver = doc.DrawPagesObs;
                        if (selectedShapes != null && pagesObserver != null)
                        {
                            List<OoShapeObserver> selectedShapesList = new List<OoShapeObserver>();
                            int count = selectedShapes.getCount();
                            for (int i = 0; i < count; i++)
                            {
                                XShape shape = selectedShapes.getByIndex(i).Value as XShape;
                                if (shape != null)
                                {
                                    OoShapeObserver shapeObserver = pagesObserver.GetRegisteredShapeObserver(shape, null);
                                    if (shapeObserver != null
                                        //&& shapeObserver.IsValid()
                                        )
                                    {
                                        if (shapeObserver.IsValid())
                                        {
                                            selectedShapesList.Add(shapeObserver);
                                        }
                                        else
                                        {
                                            shapeObserver.Dispose();
                                            XDrawPage page = OoDrawUtils.GetPageForShape(shape);
                                            OoDrawPageObserver dpObs = pagesObserver.GetRegisteredPageObserver(page);
                                            OoShapeObserver so = OoShapeObserverFactory.BuildShapeObserver(shape, dpObs); //new OoShapeObserver(shape, dpObs);
                                            pagesObserver.RegisterUniqueShape(so);
                                        }

                                    }
                                }
                            }
                            fireDrawSelectionChangedEvent(doc, selectedShapesList, aEvent == null);
                        }
                        else
                        {
                            // no selection
                            fireDrawSelectionChangedEvent(doc, new List<OoShapeObserver>(), aEvent == null);
                            return;
                        }
                    }
                }
                catch (unoidl.com.sun.star.lang.DisposedException ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Source + " " + ex.Message);
                }
            }
        }

        private readonly ConcurrentStack<KeyValuePair<OoAccessibleDocWnd, AccessibleEventObject>> selectionStack = new ConcurrentStack<KeyValuePair<OoAccessibleDocWnd, AccessibleEventObject>>();
        private Thread selectionHandlerThread;

        /// <summary>
        /// builds and starts the selection collection thread
        /// </summary>
        /// <returns></returns>
        private Thread initSelectionHandlerThread()
        {
            try
            {
                if (selectionHandlerThread != null)
                {
                    if (selectionHandlerThread.IsAlive && selectionHandlerThread.ThreadState == ThreadState.Running)
                    {
                        return selectionHandlerThread;
                    }
                    else if (selectionHandlerThread.ThreadState != ThreadState.Aborted && selectionHandlerThread.ThreadState != ThreadState.Stopped)
                    {
                        if (selectionHandlerThread.ThreadState == ThreadState.Unstarted)
                        {
                            Thread.Sleep(5);
                            if (selectionHandlerThread.ThreadState == ThreadState.Unstarted)
                            { selectionHandlerThread.Start(); }
                        }
                        return selectionHandlerThread;
                    }
                }

                selectionHandlerThread = new Thread(waitForNewSelectionChanges);
                selectionHandlerThread.Name = "OoAccesibilityObserverSelectionCollector";
                selectionHandlerThread.Start();

            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Thread exception" + ex);
            }
            return selectionHandlerThread;
        }

        /// <summary>
        /// Waits for new selection changes.
        /// </summary>
        private void waitForNewSelectionChanges()
        {
            int oldStackSize = 0;
            int counts = 0;
            while (oldStackSize < selectionStack.Count && ++counts < 30)
            {
                oldStackSize = selectionStack.Count;
                Thread.Sleep(50);
            }

            KeyValuePair<OoAccessibleDocWnd, AccessibleEventObject> e;
            bool success = selectionStack.TryPop(out e);
            if (success)
            {
                TimeLimitExecutor.ExecuteWithTimeLimit(2000, () =>
                {
                    handleSelectionChanged(e.Key, e.Value);
                }, "HandleSelectionChanged");
                selectionStack.Clear();
            }
            else
            {
                //TODO: how to do this again
            }
        }

        #endregion

        // on selection you get access to the real XShape object but not to the accessible component
        void Instance_SelectionChanged(object sender, OoSelectionChandedEventArgs e)
        {


            //var selection = e.Selection;
            //System.Diagnostics.Debug.WriteLine("Accessibility Observer: Selection changed (global Selection Listener)");
            //util.Debug.GetAllInterfacesOfObject(selection);
        }

        #region Event Throwing

        /// <summary>
        /// Occurs when a new OpenOffice draw window was opened.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> DrawWindowOpend;
        /// <summary>
        /// Occurs when some properties of a draw window has been changed.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> DrawWindowPropertyChange;
        /// <summary>
        /// Occurs when a draw window was closed.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> DrawWindowClosed;
        /// <summary>
        /// Occurs when a draw window is minimized.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> DrawWindowMinimized;
        /// <summary>
        /// Occurs when a draw window was activated.
        /// </summary>
        public event EventHandler<OoWindowEventArgs> DrawWindowActivated;

        /// <summary>
        /// Occurs when selections changed inside a draw document.
        /// </summary>
        public event EventHandler<OoAccessibilitySelectionEventArgs> DrawSelectionChanged;

        #region Fire Event

        /// <summary>
        /// Gets the last known selection collection.
        /// </summary>
        /// <value>
        /// The last selection.
        /// </value>
        public OoAccessibilitySelectionEventArgs LastSelection { get; private set; }
        private OoShapeObserver.BoundRectChangeEventHandler OnShapeBoundRectChange;
        private OoDrawPagesObserver.ViewOrZoomChangeEventHandler OnViewOrZoomChange;

        private void fireDrawWindowOpendEvent(OoAccessibleDocWnd window)
        {
            if (DrawWindowOpend != null)
            {
                try
                {
                    DrawWindowOpend.DynamicInvoke(this, new OoWindowEventArgs(window, WindowEventType.OPENED));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window opened event", ex); }
            }
            // also, get selection after window opened, but wait a bit to first get the model and controller registered
            // and register for view offset / zoom changes 
            Task t = new Task(
            () =>
            {
                Thread.Sleep(4000);
                handleSelectionChanged(window, null);
                if (window.DrawPagesObs != null)
                {
                    window.DrawPagesObs.ViewOrZoomChangeEventHandlers -= OnViewOrZoomChange;
                    window.DrawPagesObs.ViewOrZoomChangeEventHandlers += OnViewOrZoomChange;
                }
            }
            );
            t.Start();

        }
        private void fireDrawWindowPropertyChangeEvent(OoAccessibleDocWnd window)
        {
            if (DrawWindowPropertyChange != null)
            {
                try
                {
                    DrawWindowPropertyChange.DynamicInvoke(this, new OoWindowEventArgs(window, WindowEventType.CHANGED));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire window property change event", ex); }
            }
        }
        private void fireDrawWindowClosedEvent(OoAccessibleDocWnd window)
        {
            if (DrawWindowClosed != null)
            {
                try
                {
                    DrawWindowClosed.DynamicInvoke(this, new OoWindowEventArgs(window, WindowEventType.CLOSED | WindowEventType.DEACTIVATED));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire window closed event", ex); }
            }
            if (window.DrawPagesObs != null)
            {
                window.DrawPagesObs.ViewOrZoomChangeEventHandlers -= OnViewOrZoomChange;
            }
        }
        private void fireDrawWindowMinimizedEvent(OoAccessibleDocWnd window)
        {
            if (DrawWindowMinimized != null)
            {
                try
                {
                    DrawWindowMinimized.DynamicInvoke(this, new OoWindowEventArgs(window, WindowEventType.MINIMIZED | WindowEventType.DEACTIVATED));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire window minimized event", ex); }
            }
        }
        private void fireDrawWindowActivatedEvent(OoAccessibleDocWnd window)
        {
            if (DrawWindowActivated != null)
            {
                try
                {
                    DrawWindowActivated.DynamicInvoke(this, new OoWindowEventArgs(window, WindowEventType.ACTIVATED));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire window activated event", ex); }
            }
            // also, get selection on document switch
            handleSelectionChanged(window, null);
        }

        private void fireDrawSelectionChangedEvent(OoAccessibleDocWnd doc, List<OoShapeObserver> selectedShapeObservers, bool silent = false)
        {
            if (DrawSelectionChanged != null)
            {
                try
                {
                    if (LastSelection != null)
                    {
                        foreach (OoShapeObserver shapeObs in LastSelection.SelectedItems)
                        {
                            shapeObs.BoundRectChangeEventHandlers -= OnShapeBoundRectChange;
                        }
                    }
                    LastSelection = new OoAccessibilitySelectionEventArgs(doc, selectedShapeObservers, silent);
                    DrawSelectionChanged.DynamicInvoke(this, LastSelection);
                    foreach (OoShapeObserver shapeObs in LastSelection.SelectedItems)
                    {
                        shapeObs.BoundRectChangeEventHandlers += OnShapeBoundRectChange;
                    }
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire selection changed event", ex); }
            }
        }

        // called on shape bound change event, refreshes current selection bounds and invokes as silent selection change
        private void ShapeBoundRectChangeHandler()
        {
            if (LastSelection != null)
            {
                try
                {
                    LastSelection.refreshBounds();
                    LastSelection.Silent = true;
                    DrawSelectionChanged.DynamicInvoke(this, LastSelection);
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire selection changed event", ex); }
            }
        }

        #endregion

        #endregion

        #region TESTS


        #region CHILD Events

        //private void handleAccessibleChildEvent(OoAccessibleDocWnd doc, object p, uno.Any newVal, uno.Any oldVal)
        //{
        //    if (oldVal.hasValue())
        //    {
        //        /************************************************************************/
        //        /*  can be a delete or a change                                        */
        //        /************************************************************************/
        //        var val = oldVal.Value;
        //        if (val != null)
        //        {
        //            // System.Diagnostics.Debug.WriteLine("old CHILD: ");
        //            //OoAccessibility.PrintAccessibleInfos(val as XAccessible);
        //            //util.Debug.GetAllInterfacesOfObject(val);

        //            //System.Diagnostics.Debug.WriteLine("HashCode: " + val.GetHashCode());
        //            //util.Debug.GetAllServicesOfObject(val);

        //        }
        //    }
        //    handleAccessibleChildEvent(doc, p, newVal);
        //}

        //private void handleAccessibleChildEvent(OoAccessibleDocWnd doc, object p, uno.Any newVal)
        //{
        //    if (newVal.hasValue())
        //    {
        //        /************************************************************************/
        //        /*  can be an add or a change                                           */
        //        /************************************************************************/
        //        var val = newVal.Value;
        //        if (val != null)
        //        {
        //            // System.Diagnostics.Debug.WriteLine("new CHILD: ");
        //            //OoAccessibility.PrintAccessibleInfos(val as XAccessible);
        //            //util.Debug.GetAllInterfacesOfObject(val);

        //            //System.Diagnostics.Debug.WriteLine("HashCode: " + val.GetHashCode());

        //        }
        //    }
        //    else { }

        //    handleAccessibleChildEvent(doc, p);
        //}


        //private void handleAccessibleChildEvent(OoAccessibleDocWnd doc, object p)
        //{
        //    if (p != null)
        //    {
        //        XAccessible accEle = p as XAccessible;

        //        if (accEle != null)
        //        {
        //            if (OoAccessibility.GetAccessibleRole(accEle) == tud.mci.tangram.Accessibility.AccessibleRole.DOCUMENT)
        //            {

        //            }
        //        }
        //    }
        //}

        #endregion


        //#region AllAccesibleElements

        //List<OoAccComponent> accComponents = new List<OoAccComponent>();

        //Dictionary<String, Object> accCompDict = new Dictionary<String, Object>();

        //#endregion

        #endregion

        /// <summary>
        /// Resets this instance and his related Objects.
        /// </summary>
        public void Reset()
        {
            drawDocs.Clear();
            drawDocWnds.Clear();
            drawWnds.Clear();
            //drawPgSuppl.Clear();
            //drawPgsObsvr.Clear();
            //while (!drawPgsObsvrBag.IsEmpty)
            //{
            //    OoDrawPagesObserver t;
            //    drawPgsObsvrBag.TryTake(out t);
            //}
            initalized = false;
            //drawPgSuppl.InsertRange(0, OoDrawUtils.GetDrawPageSuppliers(OO.GetDesktop()));
        }

        void IDisposable.Dispose()
        {
            //TODO: do some disposables
            try{
                foreach (var item in this.drawDocs)
                {
                    item.Value.Dispose();
                }
            }catch{}
            instance = null;
        }
    }

    #region Event Args
    /// <summary>
    /// Event arguments for window events
    /// </summary>
    /// <seealso cref="System.EventArgs" />
    public class OoWindowEventArgs : EventArgs
    {
        /// <summary>
        /// The corresponding and event throwing window.
        /// </summary>
        public readonly OoAccessibleDocWnd Window;

        /// <summary>
        /// The type/reason for the event
        /// </summary>
        public readonly WindowEventType Type;

        /// <summary>
        /// Initializes a new instance of the <see cref="OoWindowEventArgs" /> class.
        /// </summary>
        /// <param name="window">The corresponding and event throwing window.</param>
        /// <param name="_type">The type of event.</param>
        public OoWindowEventArgs(OoAccessibleDocWnd window, WindowEventType _type = WindowEventType.UNKNOWN)
        {
            Window = window;
            Type = _type;
        }
    }

    /// <summary>
    /// Event arguments for selection events
    /// </summary>
    /// <seealso cref="System.EventArgs" />
    public class OoAccessibilitySelectionEventArgs : EventArgs
    {
        /// <summary>
        /// List of selected items.
        /// </summary>
        /// <value>The selected items.</value>
        public List<OoShapeObserver> SelectedItems { private set; get; }
        /// <summary>
        /// The combined bounding box of all selected items inside the document. Pixel coordinates, relative to draw view.
        /// </summary>
        /// <value>The selection bounds.</value>
        public System.Drawing.Rectangle SelectionBounds { private set; get; }
        /// <summary>
        /// The combined bounding box of all selected items on the screen. Pixel coordinates, absolute on screen.
        /// </summary>
        /// <value>The selection screen bounds.</value>
        public System.Drawing.Rectangle SelectionScreenBounds { private set; get; }
        /// <summary>
        /// The corresponding window/document the event is thrown from.
        /// </summary>
        public readonly OoAccessibleDocWnd Source;
        /// <summary>
        /// Flag that marks the event as not to be announced by the audio renderer, e.g. for just updating the boundaries if the selection stays the same but selected elements are resized/moved. 
        /// False by default.
        /// </summary>
        public bool Silent { set; get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccessibilitySelectionEventArgs"/> class.
        /// </summary>
        /// <param name="source">The corresponding window/document the event is thrown from.</param>
        /// <param name="selectedItems">The selected items.</param>
        /// <param name="type">The accessibility event type resulting in this event.</param>
        public OoAccessibilitySelectionEventArgs(OoAccessibleDocWnd source, List<OoShapeObserver> selectedShapeObservers, bool silent)
        {
            Source = source;
            SelectedItems = selectedShapeObservers;
            updateBounds(selectedShapeObservers);
            this.Silent = silent;
        }

        private void updateBounds(List<OoShapeObserver> items)
        {
            System.Drawing.Rectangle bounds = new System.Drawing.Rectangle();
            System.Drawing.Rectangle screenBounds = new System.Drawing.Rectangle();

            foreach (OoShapeObserver item in items)
            {
                //System.Drawing.Rectangle pageBounds = Source.DocumentComponent.ScreenBounds;
                System.Drawing.Rectangle shapeScreenBoundsRelative = item.GetRelativeScreenBoundsByDom();
                System.Drawing.Rectangle tempSbRel = new System.Drawing.Rectangle(shapeScreenBoundsRelative.X, shapeScreenBoundsRelative.Y, shapeScreenBoundsRelative.Width, shapeScreenBoundsRelative.Height);
                System.Drawing.Rectangle shapeScreenBoundsAbsolute = item.GetAbsoluteScreenBoundsByDom(shapeScreenBoundsRelative);
                System.Drawing.Rectangle tempSbAbs = new System.Drawing.Rectangle(shapeScreenBoundsAbsolute.X, shapeScreenBoundsAbsolute.Y, shapeScreenBoundsAbsolute.Width, shapeScreenBoundsAbsolute.Height);

                if (shapeScreenBoundsRelative.Height > 0 || shapeScreenBoundsRelative.Width > 0)
                {
                    if (bounds.Width == 0 && bounds.Height == 0) //initial set
                    {
                        bounds = tempSbRel;
                        screenBounds = tempSbAbs;
                    }
                    else
                    {
                        int right = Math.Max(bounds.Right, shapeScreenBoundsRelative.Right);
                        int bottom = Math.Max(bounds.Bottom, shapeScreenBoundsRelative.Bottom);

                        bounds.X = Math.Min(bounds.X, shapeScreenBoundsRelative.X);
                        bounds.Y = Math.Min(bounds.Y, shapeScreenBoundsRelative.Y);

                        screenBounds.X = Math.Min(screenBounds.X, tempSbAbs.X);
                        screenBounds.Y = Math.Min(screenBounds.Y, tempSbAbs.Y);

                        bounds.Width = right - bounds.X;
                        bounds.Height = bottom - bounds.Y;

                        screenBounds.Size = bounds.Size;
                    }
                }
            }
            this.SelectionBounds = bounds;
            this.SelectionScreenBounds = screenBounds;
        }

        public void refreshBounds()
        {
            if (SelectedItems != null) updateBounds(SelectedItems);
        }
    }

    #endregion
}