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

        /// <summary>
        /// The delay time until selection changes will be collected before forwarding them (in milliseconds).
        /// </summary>
        public static int SelectionDelayTime = 100;

        #endregion

        #region Singleton Construction

        private static OoDrawAccessibilityObserver instance = new OoDrawAccessibilityObserver();

        private OoDrawAccessibilityObserver()
        {
            //OoSelectionObserver.Instance.SelectionChanged += new EventHandler<OoSelectionChandedEventArgs>(Instance_SelectionChanged);
            //drawPgSuppl = OoDrawUtils.GetDrawPageSuppliers(OO.GetDesktop());

            //OnShapeBoundRectChange = new OoShapeObserver.BoundRectChangeEventHandler(ShapeBoundRectChangeHandler);
            //OnViewOrZoomChange = new OoDrawPagesObserver.ViewOrZoomChangeEventHandler(ShapeBoundRectChangeHandler);
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

        #endregion

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

        #region OoAccessibleDocWnd Events

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
                SelectionChanged(sender as OoAccessibleDocWnd, new AccessibleEventObject());
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
                    handleAccessibleChildEvent(doc, aEvent.Source, aEvent.NewValue, aEvent.OldValue);
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
                    SelectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_ADD:
                    SelectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_REMOVE:
                    SelectionChanged(doc, aEvent);
                    break;
                case tud.mci.tangram.Accessibility.AccessibleEventId.SELECTION_CHANGED_WITHIN:
                    SelectionChanged(doc, aEvent);
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

        private void handleAccessibleChildEvent(OoAccessibleDocWnd doc, object p, uno.Any newValue, uno.Any oldValue)
        {
            Task t = new Task(() =>
            {
                util.Debug.GetAllInterfacesOfObject(p);
                if (doc != null && !newValue.hasValue() && oldValue.hasValue())
                {
                    // shape was deleted!

                    var oldShapeAcc = oldValue.Value;

                    util.Debug.GetAllInterfacesOfObject(oldShapeAcc);


                    if (p != null && p is XAccessibleComponent)
                    {
                        var shapeObs = doc.GetRegisteredShapeObserver(p as XAccessibleComponent);
                        if (shapeObs != null && !shapeObs.IsValid(true))
                        {
                            shapeObs.Dispose();
                        }
                        else
                        {
                            doc.Update();
                        }
                    }
                }
            });
            t.Start();
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

        #region Selection Change Handling

        /// <summary>
        /// add a new selection event th the selection stack
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="aEvent"></param>
        internal void SelectionChanged(OoAccessibleDocWnd doc, AccessibleEventObject aEvent)
        {
            if (aEvent != null)
            {
                resetSelectionCache(doc);
                resetSelectionDelayTimer(doc);
            }
        }


        Timer selecttionDelayTimer = null;
        readonly Object _selectiondelayLock = new Object();

        void resetSelectionDelayTimer(OoAccessibleDocWnd doc)
        {
            lock (_selectiondelayLock)
            {
                if (selecttionDelayTimer != null)
                {
                    selecttionDelayTimer.Change(SelectionDelayTime, Timeout.Infinite);
                }
                else
                {
                    selecttionDelayTimer = new Timer(handleSelectionChanged, doc, SelectionDelayTime, Timeout.Infinite);
                }
            }
        }

        private void handleSelectionChanged(Object doc)
        {
            //System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " SELECTION EVENT FORWARD BECAUSE TIMER ELAPSED.");
            handleSelectionChanged(doc as OoAccessibleDocWnd, new AccessibleEventObject());
        }
        private void handleSelectionChanged(OoAccessibleDocWnd ooAccessibleDocWnd, AccessibleEventObject accessibleEventObject)
        {
            lock (_selectiondelayLock)
            {
                fireDrawSelectionChangedEvent(ooAccessibleDocWnd, true);
                if (selecttionDelayTimer != null)
                {
                    selecttionDelayTimer.Dispose();
                    selecttionDelayTimer = null;
                }
            }
        }

        #endregion

        #endregion

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
        public event EventHandler<OoAccessibilitySelectionChangedEventArgs> DrawSelectionChanged;

        #region Fire Event

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
                //System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " SELECTION EVENT FORWARD BECAUSE WINDOW OPEND.");
                handleSelectionChanged(window, null);
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
            //System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " SELECTION EVENT FORWARD BECAUSE WINDOW ACTIVATED.");
            // also, get selection on document switch
            handleSelectionChanged(window, null);
        }

        private void fireDrawSelectionChangedEvent(OoAccessibleDocWnd doc, bool silent = false)
        {
            if (DrawSelectionChanged != null)
            {
                try
                {
                    //System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " SELECTION EVENT FORWARD");
                    DrawSelectionChanged.Invoke(this, new OoAccessibilitySelectionChangedEventArgs(doc, silent));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire selection changed event", ex); }
            }
        }

        #endregion

        #endregion

        #region public Functions

        #region Selection

        readonly ConcurrentDictionary<OoAccessibleDocWnd, OoAccessibilitySelection> _cachedSelection = new ConcurrentDictionary<OoAccessibleDocWnd, OoAccessibilitySelection>();

        void resetSelectionCache(OoAccessibleDocWnd doc = null)
        {
            _cachedSelection.Clear();
            //if (doc == null) 
            //else if (_cachedSelection.ContainsKey(doc))
            //{
            //    OoAccessibilitySelection trash;
            //    int i = 0;
            //    while (!_cachedSelection.TryRemove(doc, out trash) && i++ < 3) { Thread.Sleep(5); }
            //}
        }

        readonly object _selectionLock = new Object();
        /// <summary>
        /// Tries to the get current selection.
        /// </summary>
        /// <param name="doc">The document window.</param>
        /// <param name="selectedShapesList">The list of currently selected shapes.</param>
        /// <returns><c>true</c> if the selection could be achieved; otherwise <c>false</c> (e.g. because it was aborted by time limit)</returns>
        /// <remarks>This request is time limited to 300 ms.</remarks>
        public bool TryGetSelection(OoAccessibleDocWnd doc, out OoAccessibilitySelection selection)
        {
            selection = null;
            List<OoShapeObserver> selectedShapesList2 = new List<OoShapeObserver>();
            bool innerSuccess = false;
            bool success = false;

            if (doc != null)
            {
                try
                {
                    if (_cachedSelection.ContainsKey(doc))
                    {
                        int i = 0;
                        while (!_cachedSelection.TryGetValue(doc, out selection) && i++ < 3) { Thread.Sleep(5); }
                        if (i < 3)
                        {
                            if (selection != null
                                && DateTime.Now - selection.SelectionCreationTime < new TimeSpan(0, 1, 0))
                            {
                                //System.Diagnostics.Debug.WriteLine("******* GET Cached Selection for WND: " + doc + " result in " + selection.Count + " selected Items.");
                                return true;
                            }
                        }
                    }

                    success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(300, () =>
                    {
                        try
                        {
                            lock (doc.SynchLock)
                            {
                                innerSuccess = tryGetSelection(doc, out selectedShapesList2);
                            }
                        }
                        catch { innerSuccess = false; }
                    }, "HandleSelectionChanged");

                    var selection2 = new OoAccessibilitySelection(doc, selectedShapesList2);
                    if (success & innerSuccess)
                    {
                        //System.Diagnostics.Debug.WriteLine("++++++ ADD to cached Selection for WND: " + doc + " result in " + selection2.Count + " selected Items.");
                        _cachedSelection.AddOrUpdate(doc, selection2, (key, oldValue) => selection2);
                        selection = selection2;
                    }
                    //else resetSelectionCache();                    
                }
                catch (ThreadInterruptedException) { }
                catch (ThreadAbortException) { }
            }

            //System.Diagnostics.Debug.WriteLine(".... Selection Return: " + (selection != null ? selection.Count.ToString() : " invalid selection"));

            return success & innerSuccess;
        }

        private bool tryGetSelection(OoAccessibleDocWnd doc, out List<OoShapeObserver> selectedShapesList)
        {
            //System.Diagnostics.Debug.WriteLine("  ---> try Get Selection (inner critical Call)");
            selectedShapesList = new List<OoShapeObserver>();
            bool success = false;

            // check the global selection supplier
            if (doc != null)
            {
                try
                {
                    var controller = doc.Controller;
                    if (controller != null && controller is XSelectionSupplier)
                    {
                        Object selection = OoSelectionObserver.GetSelection(controller as XSelectionSupplier);
                        XShapes selectedShapes = selection as XShapes;

                        OoDrawPagesObserver pagesObserver = doc.DrawPagesObs;
                        if (selectedShapes != null && pagesObserver != null)
                        {

                            int count = selectedShapes.getCount();
                            for (int i = 0; i < count; i++)
                            {
                                XShape shape = selectedShapes.getByIndex(i).Value as XShape;
                                if (shape != null)
                                {
                                    OoShapeObserver shapeObserver = pagesObserver.GetRegisteredShapeObserver(shape, null);
                                    if (shapeObserver != null)
                                    {
                                        selectedShapesList.Add(shapeObserver);
                                    }
                                }
                            }
                            success = true;
                        }
                        else
                        {
                            // no selection
                            if (selection is bool && ((bool)selection) == false) success = false;
                            else if (pagesObserver != null) success = true;
                        }
                    }
                }
                catch (unoidl.com.sun.star.lang.DisposedException ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Source + " " + ex.Message);
                }
            }

            //System.Diagnostics.Debug.WriteLine("  ---> ~~~~~~~~~ (" + success + ") GET Selection for WND: " + doc + " result in " + selectedShapesList.Count + " selected Items.");

            return success;
        }

        #endregion

        /// <summary>
        /// Resets this instance and his related Objects.
        /// </summary>
        public void Reset()
        {
            drawDocs.Clear();
            drawDocWnds.Clear();
            drawWnds.Clear();
            resetSelectionCache();
            initalized = false;
        }

        #endregion

        #region IDisposable

        void IDisposable.Dispose()
        {
            try
            {
                foreach (var item in this.drawDocs)
                {
                    item.Value.Dispose();
                }
            }
            catch { }
            instance = null;
        }

        #endregion

    }

    #region Event Args and Special Classes

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
    /// Event arguments for selection changed events
    /// </summary>
    /// <seealso cref="System.EventArgs" />
    public class OoAccessibilitySelectionChangedEventArgs : EventArgs
    {
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
        public OoAccessibilitySelectionChangedEventArgs(OoAccessibleDocWnd source, bool silent = false)
        {
            Source = source;
            this.Silent = silent;
        }
    }

    /// <summary>
    /// Selection of objects
    /// </summary>
		/// <remarks> </remarks>
    public class OoAccessibilitySelection
    {
        #region private Member

        private readonly OoShapeObserver.BoundRectChangeEventHandler OnShapeBoundRectChange;
        private readonly OoDrawPagesObserver.ViewOrZoomChangeEventHandler OnViewOrZoomChange;

        #endregion

        #region public Members

        /// <summary>
        /// The timestamps when this selection collection was build.
        /// </summary>
        public readonly DateTime SelectionCreationTime = DateTime.Now;
        /// <summary>
        /// Gets the amount of selected Items.
        /// </summary>
        /// <value>
        /// The count of selected items.
        /// </value>
        public int Count { get { return SelectedItems != null ? SelectedItems.Count : 0; } }
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

        #endregion

        #region Constructor / Destructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoAccessibilitySelectionEventArgs"/> class.
        /// </summary>
        /// <param name="source">The corresponding window/document the event is thrown from.</param>
        /// <param name="selectedItems">The selected items.</param>
        /// <param name="type">The accessibility event type resulting in this event.</param>
        public OoAccessibilitySelection(OoAccessibleDocWnd source, List<OoShapeObserver> selectedShapeObservers)
        {
            OnShapeBoundRectChange = new OoShapeObserver.BoundRectChangeEventHandler(shapeBoundRectChangeHandler);
            OnViewOrZoomChange = new OoDrawPagesObserver.ViewOrZoomChangeEventHandler(shapeBoundRectChangeHandler);
            Source = source;
            SelectedItems = selectedShapeObservers;
            RefreshBounds();
            registerToEvents();
        }

        ~OoAccessibilitySelection()
        {
            try
            {
                if (Count > 0)
                {
                    foreach (var item in SelectedItems)
                    {
                        try
                        {
                            item.BoundRectChangeEventHandlers -= OnShapeBoundRectChange;
                            item.BoundRectChangeEventHandlers += OnShapeBoundRectChange;
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        #endregion

        #region BoundsHandling

        DateTime _lastUpdate = DateTime.MinValue;
        TimeSpan _updateTimeOut = new TimeSpan(0, 0, 1);

        /// <summary>
        /// Refreshes the bounds.
        /// </summary>
        public void RefreshBounds()
        {
            var now = DateTime.Now;
            if (now - _lastUpdate > _updateTimeOut)
            {
                this.refreshBounds(false);
                _lastUpdate = now;
            }
        }
        private void refreshBounds(bool throwEvent = false)
        {
            if (Count > 0)
            {
                Task t = new Task(new Action(() => updateBounds(throwEvent)));
                t.Start();
            }
        }

        private void updateBounds(bool throwEvent = false)
        {
            var success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(500, () =>
            {
                System.Drawing.Rectangle bounds = new System.Drawing.Rectangle();
                System.Drawing.Rectangle screenBounds = new System.Drawing.Rectangle();

                if (SelectedItems != null && SelectedItems.Count > 0)
                {
                    foreach (OoShapeObserver item in SelectedItems)
                    {
                        System.Drawing.Rectangle tempSbRel = item.GetRelativeScreenBoundsByDom();
                        System.Drawing.Rectangle tempSbAbs = item.GetAbsoluteScreenBoundsByDom(tempSbRel);

                        if (tempSbRel.Height > 0 || tempSbRel.Width > 0)
                        {
                            if (bounds.Width == 0 && bounds.Height == 0) //initial set
                            {
                                bounds = tempSbRel;
                                screenBounds = tempSbAbs;
                            }
                            else
                            {
                                int right = Math.Max(bounds.Right, tempSbRel.Right);
                                int bottom = Math.Max(bounds.Bottom, tempSbRel.Bottom);

                                bounds.X = Math.Min(bounds.X, tempSbRel.X);
                                bounds.Y = Math.Min(bounds.Y, tempSbRel.Y);

                                screenBounds.X = Math.Min(screenBounds.X, tempSbAbs.X);
                                screenBounds.Y = Math.Min(screenBounds.Y, tempSbAbs.Y);

                                bounds.Width = right - bounds.X;
                                bounds.Height = bottom - bounds.Y;

                                screenBounds.Size = bounds.Size;
                            }
                        }
                    }
                }
                this.SelectionBounds = bounds;
                this.SelectionScreenBounds = screenBounds;
            }, "updateBoundsOfShapes");

            if (success == throwEvent)
            {
                OoDrawAccessibilityObserver.Instance.SelectionChanged(Source, null);
            }
        }

        private void registerToEvents()
        {
            if (Count > 0)
            {
                foreach (var item in SelectedItems)
                {
                    item.BoundRectChangeEventHandlers -= OnShapeBoundRectChange;
                    item.BoundRectChangeEventHandlers += OnShapeBoundRectChange;
                }
            }
        }

        // called on shape bound change event, refreshes current selection bounds 
        // and invokes as silent selection change
        private void shapeBoundRectChangeHandler()
        {
            try
            {
                refreshBounds(true);
            }
            catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "cant fire selection changed event", ex); }
        }

        #endregion

        #region Override

        public override bool Equals(object obj)
        {
            if (obj != null && obj is OoAccessibilitySelection)
            {
                OoAccessibilitySelection objS = obj as OoAccessibilitySelection;
                if (objS.Source.Equals(this.Source))
                {
                    if (objS.Count == this.Count)
                    {
                        if (objS.Count > 0)
                            foreach (var item in objS.SelectedItems)
                            {
                                if (!this.SelectedItems.Contains(item))
                                    return false;
                            }
                        return true;
                    }
                }
            }

            return false;
        }

        public override int GetHashCode()
        {
            int hash = base.GetHashCode();

            hash += Source.GetHashCode();
            if (Count > 0)
            {
                foreach (var item in SelectedItems)
                {
                    hash += item.GetHashCode();
                }
            }
            return hash.GetHashCode();
        }

        #endregion
    }

    #endregion
}