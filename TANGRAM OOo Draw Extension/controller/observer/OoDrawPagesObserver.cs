using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.util;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Observes all draw pages of this oo document instance
    /// </summary>
    public class OoDrawPagesObserver : PropertiesEventForwarderBase, IUpdateable, IDisposingObserver
    {
        #region Members
        /// <summary>
        /// The Accessible Document Window (implementing com.sun.start.drawing.AccessibleDrawDocumentView). A reference, set by constructor for easy access to its screen coordinates.
        /// </summary>
        public readonly OoAccessibleDocWnd DocWnd;
        internal XDrawPagesSupplier PagesSupplier { get; private set; }
        private readonly ConcurrentDictionary<XDrawPage, OoDrawPageObserver> DrawPageobservers = new ConcurrentDictionary<XDrawPage, OoDrawPageObserver>();

        private XController2 _ctrl = null;
        private readonly Object _ctrlLock = new Object();
        /// <summary>
        /// Gets the controller from the startup for this DRAW document instance.
        /// </summary>
        /// <value>The controller.</value>
        internal XController2 Controller
        {
            get
            {
                //lock (_ctrlLock)
                {
                    // get Zoom and ViewOffset first time
                    if (_ctrl == null && PagesSupplier != null && PagesSupplier is XModel)
                    {
                        XController2 c = (XController2)((XModel)PagesSupplier).getCurrentController();
                        if (c != null)
                        {
                            _ctrl = c;
                        }
                    }
                    return _ctrl;
                }
            }
        }

        #region uniqie id for shapes
        /// <summary>
        /// Dictionary of shape names to their observer
        /// </summary>
        private readonly ConcurrentDictionary<String, OoShapeObserver> shapes = new ConcurrentDictionary<String, OoShapeObserver>();
        /// <summary>
        /// Dictionary of shape accessible context to their observer
        /// </summary>
        private readonly ConcurrentDictionary<XAccessibleContext, OoShapeObserver> accshapes = new ConcurrentDictionary<XAccessibleContext, OoShapeObserver>();
        /// <summary>
        /// Dictionary of xShapes to their observer
        /// </summary>
        private readonly ConcurrentDictionary<XShape, OoShapeObserver> domshapes = new ConcurrentDictionary<XShape, OoShapeObserver>();

        #endregion

        /// <summary>
        /// Title of the document. Ends with .odg or even another file extension.
        /// </summary>
        public String Title = String.Empty;

        private OoAccComponent _doc;
        public OoAccComponent Document
        {
            get
            {
                if (_doc.Role == tud.mci.tangram.Accessibility.AccessibleRole.UNKNOWN
                    || _doc.Role == tud.mci.tangram.Accessibility.AccessibleRole.INVALID) Document = new OoAccComponent(OoDrawUtils.GetXAccessibleFromDrawPagesSupplier(PagesSupplier));
                return _doc;
            }
            set
            {
                unregisterDocumentListener();
                _doc = value;
                addDocumentListeners();
            }
        }

        //private readonly XSelectionListener_Forwarder xSelectionListener = new XSelectionListener_Forwarder();

        #endregion

        #region Constructor / Destructor

        /// <summary>
        /// static constructor for initializing system display DPI settings once
        /// </summary>
        static OoDrawPagesObserver()
        {
            double ppmX, ppmY;
            OoUtils.GetDeviceResolutionPixelPerMeter(out ppmX, out ppmY);
            PixelPerMeterX = ppmX;
            PixelPerMeterY = ppmY;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OoDrawPagesObserver"/> class.
        /// </summary>
        /// <param name="dp">The Draw document.</param>
        /// <param name="doc">The document related accessibility component.</param>
        /// <param name="docWnd">The related document accessible window component.</param>
        public OoDrawPagesObserver(XDrawPagesSupplier dp, OoAccComponent doc, OoAccessibleDocWnd docWnd = null)
        {
            this.PagesSupplier = dp;
            Document = doc;
            DocWnd = docWnd;

            // get Zoom and ViewOffset first time
            if (Controller != null)
            {
                if (Controller is XPropertySet)
                {
                    refreshDrawViewProperties((XPropertySet)(Controller));
                }
                // try to get dpi settings from openoffice
                XWindow componentWindow = Controller.ComponentWindow;
                if (componentWindow != null && componentWindow is XDevice)
                {
                    DeviceInfo deviceInfo = (DeviceInfo)((XDevice)componentWindow).getInfo();
                    if (deviceInfo != null)
                    {
                        PixelPerMeterX = deviceInfo.PixelPerMeterX;
                        PixelPerMeterY = deviceInfo.PixelPerMeterY;
                    }
                }
            }
            // register for Zoom and ViewOffset updates
            addVisibleAreaPropertyChangeListener();

            if (this.PagesSupplier != null)
            {
                List<XDrawPage> dpL = OoDrawUtils.DrawDocGetXDrawPageList(dp);

                if (PagesSupplier is unoidl.com.sun.star.frame.XTitle)
                {
                    Title = ((unoidl.com.sun.star.frame.XTitle)PagesSupplier).getTitle();
                }

                Logger.Instance.Log(LogPriority.DEBUG, this, "create DrawPagesObserver for supplier " + dp.GetHashCode() + " width title '" + Title + "' - having " + dpL.Count + " pages");

                //FIXME: Do this if the api enable parallel access
                //Parallel.ForEach(dpL, (drawPage) =>
                //{
                //    OoDrawPageObserver dpobs = new OoDrawPageObserver(drawPage, this);
                //    DrawPageobservers[drawPage] = dpobs;
                //    DrawPages.Add(dpobs);
                //});

                foreach (var drawPage in dpL)
                {
                    OoDrawPageObserver dpobs = new OoDrawPageObserver(drawPage, this);
                    RegisterDrawPage(dpobs);
                }

                XModifyBroadcaster mdfBc = PagesSupplier as XModifyBroadcaster;
                if (mdfBc != null)
                {
                    mdfBc.addModifyListener(eventForwarder);
                }
            }
        }

        ~OoDrawPagesObserver()
        {
            removeVisibleAreaPropertyChangeListener();
        }

        #endregion

        #region Draw pages

        /// <summary>
        /// Registers a new draw page observer.
        /// </summary>
        /// <param name="dpobs">The dpobs.</param>
        internal void RegisterDrawPage(OoDrawPageObserver dpobs)
        {
            if (dpobs != null && dpobs.DrawPage != null)
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] Register new DrawPage");
                DrawPageobservers[dpobs.DrawPage] = dpobs;
                // DrawPages.Add(dpobs);
                dpobs.ObserverDisposing += new EventHandler(dpobs_ObserverDisposing);
            }
        }

        void dpobs_ObserverDisposing(object sender, EventArgs e)
        {
            OoDrawPageObserver dpobs = sender as OoDrawPageObserver;
            if (dpobs != null)
            {
                //DrawPages.

                foreach (var item in DrawPageobservers)
                {
                    if (item.Value == dpobs)
                    {
                        OoDrawPageObserver trash;
                        DrawPageobservers.TryRemove(item.Key, out trash);
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] DrawPage removed");
                    }

                    if (!HasDrawPages())
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] Last page disposed");
                    }
                }
            }
        }

        /// <summary>
        /// Gets all known draw pages of this document.
        /// </summary>
        /// <returns></returns>
        public List<OoDrawPageObserver> GetDrawPages()
        {
            return DrawPageobservers.Values.ToList();
        }

        /// <summary>
        /// Determines whether this document knows about draw some pages.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if knows seome draw pages; otherwise, <c>false</c>.
        /// </returns>
        public bool HasDrawPages() { return DrawPageobservers != null && DrawPageobservers.Count > 0; }

        /// <summary>
        /// Gets the registered page observer to the DOM page object.
        /// </summary>
        /// <param name="dp">The DOM draw page.</param>
        /// <returns>The already registered page observer or register a new one and return it.</returns>
        internal OoDrawPageObserver GetRegisteredPageObserver(XDrawPage dp)
        {
            OoDrawPageObserver dpobs = null;
            if (dp != null)
            {
                if (DrawPageobservers.ContainsKey(dp))
                {
                    dpobs = DrawPageobservers[dp];
                }
                else
                {
                    dpobs = new OoDrawPageObserver(dp, this);
                    RegisterDrawPage(dpobs);
                }
            }
            return dpobs;
        }

        #endregion

        #region IUpdateable

        /// <summary>
        /// Updates this instance. Updates all PageObservers and there children
        /// </summary>
        public void Update() { Update(PagesSupplier); }
        private void Update(XDrawPagesSupplier dps)
        {
            List<XDrawPage> dpL = OoDrawUtils.DrawDocGetXDrawPageList(dps);

            Parallel.ForEach(dpL, (drawPage) =>
            {
                if (DrawPageobservers.ContainsKey(drawPage))
                {
                    System.Diagnostics.Debug.WriteLine("[UPDATE] Draw page known");
                    var obs = DrawPageobservers[drawPage];
                    obs.Update();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[UPDATE] Draw page NOT known !!!");
                    OoDrawPageObserver dpobs = new OoDrawPageObserver(drawPage, this);
                    RegisterDrawPage(dpobs);
                }
            });
        }

        #endregion

        #region utils

        private void unregisterDocumentListener()
        {
            try
            {
                if (_doc != null && _doc.AccComp != null && _doc.AccComp is XAccessibleEventBroadcaster)
                {
#if LIBRE
                    ((XAccessibleEventBroadcaster)_doc.AccComp).removeAccessibleEventListener(eventForwarder);
#else
                    ((XAccessibleEventBroadcaster)_doc.AccComp).removeEventListener(eventForwarder);
#endif
                }
            }
            catch
            {
                try
                {
                    if (_doc.AccComp != null && _doc.AccComp is XAccessibleEventBroadcaster)
                    {
#if LIBRE
                        ((XAccessibleEventBroadcaster)_doc.AccComp).removeAccessibleEventListener(eventForwarder);
#else
                        ((XAccessibleEventBroadcaster)_doc.AccComp).removeEventListener(eventForwarder);
#endif
                    }
                }
                catch { }
            }
        }

        private void addDocumentListeners()
        {
            if (_doc.AccComp != null)
            {
                if (_doc.AccComp is XAccessibleEventBroadcaster)
                {
                    try
                    {
#if LIBRE
                        ((XAccessibleEventBroadcaster)_doc.AccComp).removeAccessibleEventListener(eventForwarder);
#else
                        ((XAccessibleEventBroadcaster)_doc.AccComp).removeEventListener(eventForwarder);
#endif

                    }
                    catch
                    {
                        try
                        {
#if LIBRE
                            ((XAccessibleEventBroadcaster)_doc.AccComp).removeAccessibleEventListener(eventForwarder);
#else
                            ((XAccessibleEventBroadcaster)_doc.AccComp).removeEventListener(eventForwarder);
#endif
                        }
                        catch (Exception e) { Logger.Instance.Log(LogPriority.ALWAYS, this, "can't remove document listener", e); }
                    }

                    try
                    {
#if LIBRE
                        ((XAccessibleEventBroadcaster)_doc.AccComp).addAccessibleEventListener(eventForwarder);
#else
                        ((XAccessibleEventBroadcaster)_doc.AccComp).addEventListener(eventForwarder);
#endif
                    }
                    catch (System.Exception e)
                    {
                        Logger.Instance.Log(LogPriority.ALWAYS, this, "can't add document listener: ", e);
                    }
                }
            }
        }

        #region Unique Identifier

        private readonly ConcurrentBag<String> shapeIds = new ConcurrentBag<String>();

        /// <summary>
        /// Try to generate an unique id.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns></returns>
        public string GetUniqueId(OoShapeObserver shape)
        {
            int uidCount = 0;
            String uid = String.Empty;
            if (shape != null && shape.Shape != null)
            {
                if (OoUtils.ElementSupportsService(shape.Shape, OO.Services.DRAW_SHAPE_TEXT))
                {
                    uid = shape.Name;
                    if (String.IsNullOrWhiteSpace(uid)
                        || uid.Contains('\'')
                        || uid.Contains(' '))
                    {
                        uid = shape.UINamePlural;
                        System.Diagnostics.Debug.WriteLine("______the start name for the text shape is now : '" + uid + "'");
                    }
                }
                else
                {
                    uid = shape.Name;
                }
                if (String.IsNullOrWhiteSpace(uid))
                {
                    uid = shape.UINameSingular + "_" + (++uidCount);
                }

                while (shapeIds.Contains(uid))
                {
                    int i_ = uid.LastIndexOf('_');
                    if (i_ >= 0) { uid = uid.Substring(0, i_ + 1); }
                    else { uid += "_"; }
                    uid += ++uidCount;
                }
                shapeIds.Add(uid);
            }
            Logger.Instance.Log(LogPriority.DEBUG, this, "new Shape with name: " + uid + " registered");
            return uid;
        }

        /// <summary>
        /// Registers a unique shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>the unique identifier (Name) to which the shape is registered</returns>
        public string RegisterUniqueShape(OoShapeObserver shape)
        {
            String uid = String.Empty;
            OoShapeObserver trash;
            //TODO: check if name already exists
            if (shape != null && shape.IsValid())
            {
                // register for disposing element
                shape.ObserverDisposing += new EventHandler(shape_ObserverDisposing);

                uid = OoUtils.GetStringProperty(shape, "Name");
                if (!String.IsNullOrWhiteSpace(uid))
                {
                    // check if name exists
                    if (shapes.ContainsKey(uid))
                    {
                        OoShapeObserver sObs = shapes[uid];
                        if (sObs != null)
                        {
                            if (sObs.IsValid())
                            {
                                // already valid registered
                                if (sObs.Shape == shape.Shape)
                                {
                                    return uid;
                                }
                                else // valid different observer with the same name exist
                                {
                                    if (!uid.Equals(sObs.Name)) // update the registered name in the list if the name has changed
                                    {
                                        shapes.TryRemove(uid, out trash);
                                        uid = sObs.Name;
                                        shapes.TryAdd(uid, shape);
                                    }
                                }
                            }
                            else
                            {
                                sObs.Dispose();
                            }
                        }
                        else
                        {
                            shapes.TryRemove(uid, out trash);
                        }
                    }

                    // check if shape is already registered
                    if (domshapes.ContainsKey(shape.Shape))
                    {
                        OoShapeObserver sObs = domshapes[shape.Shape];
                        if (sObs != null)
                        {
                            if (sObs.IsValid())
                            {
                                domshapes[shape.Shape] = shape;
                                return uid;
                            }
                            else
                            {
                                sObs.Dispose();
                            }
                        }
                        else
                        {
                            domshapes.TryRemove(shape.Shape, out trash);
                        }
                    }
                }
                else
                {
                    uid = GetUniqueId(shape);
                }

                while (shapes.ContainsKey(uid))
                {
                    uid += "*";
                }
                shape.Name = uid;
                shapes[uid] = shape;
                if (shape.Shape != null) domshapes[shape.Shape] = shape;
                if (shape.AcccessibleCounterpart != null) accshapes[shape.AcccessibleCounterpart.getAccessibleContext()] = shape;

            }
            return uid;
        }

        /// <summary>
        /// Determine if the shapes is already registered.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns><c>true</c> if the shape is allready known otherwise <c>false</c> </returns>
        public bool ShapeAlreadyRegistered(XShape shape, OoDrawPageObserver page = null)
        {
            if (shape != null)
            {
                if (domshapes.ContainsKey(shape))
                {
                    OoShapeObserver sobs = domshapes[shape];
                    if (sobs != null)
                    {
                        if (sobs.Shape == shape) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Get a registered shape observer.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>the already registerd shape observer to the shape or <c>null</c></returns>
        internal OoShapeObserver GetRegisteredShapeObserver(XShape shape, OoDrawPageObserver page)
        {
            if (domshapes.ContainsKey(shape))
            {
                OoShapeObserver sobs = domshapes[shape];
                if (sobs != null && sobs.Shape == shape)
                {
                    return sobs;
                }
            }

            String name = OoUtils.GetStringProperty(shape, "Name");
            OoShapeObserver sObs = getRegisteredShapeObserver(name);
            if (sObs == null)
            {
                //TODO: handle this
                if (page == null)
                {

                    XDrawPage pageShape = OoDrawUtils.GetPageForShape(shape);
                    if (pageShape == null)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[EROR] Can't get page to requested NEW shape");
                        page = this.DocWnd.GetActivePage();

                    }
                    else
                    {
                        page = GetRegisteredPageObserver(pageShape as XDrawPage);
                    }
                }

                if (page != null)
                {
                    sObs = new OoShapeObserver(shape, page);
                    RegisterUniqueShape(sObs);
                }
            }

            //TODO: check if valid
            //TODO: check if the same
            if (sObs.Shape != shape)
            {
                sObs = RegisterNewShape(shape, page);
            }

            return sObs;
        }

        /// <summary>
        /// Registers a new shape.
        /// </summary>
        /// <param name="shape">The dom shape.</param>
        /// <returns>a registered <see cref="OoShapeObserver"/> for this shape</returns>
        internal OoShapeObserver RegisterNewShape(XShape shape, OoDrawPageObserver pObs = null)
        {
            OoShapeObserver sobs = null;

            if (shape != null)
            {
                // get the page to this shape
                var page = OoDrawUtils.GetPageForShape(shape);
                if (page != null)
                {
                    if(pObs == null || !page.Equals(pObs.DrawPage)) 
                        pObs = GetRegisteredPageObserver(page);

                    if (pObs != null)
                    {
                        sobs = new OoShapeObserver(shape, pObs);
                        RegisterUniqueShape(sobs);
                    }
                }
            }

            return sobs;
        }

        /// <summary>
        /// Gets the registered shape observer.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns></returns>
        internal OoShapeObserver GetRegisteredShapeObserver(unoidl.com.sun.star.accessibility.XAccessibleContext shape)
        {
            if (shape == null) return null;
            if (accshapes.ContainsKey(shape))
                return accshapes[shape];

            String name = OoAccessibility.GetAccessibleName(shape);
            String desc = OoAccessibility.GetAccessibleDesc(shape);

            OoShapeObserver sObs = searchRegisterdSapeObserverByAccessibleName(name, shape);

            if (sObs != null)
            {
                
            }
            else
            {
                //shape is unknown or the name was changed!!

                Logger.Instance.Log(LogPriority.IMPORTANT, this, "Asking for a not registered Element '" + name + "' ");
            }
            return sObs;
        }

        /// <summary>
        /// Gets the registered shape observer.
        /// </summary>
        /// <param name="comp">The comp.</param>
        /// <returns></returns>
        public OoShapeObserver GetRegisteredShapeObserver(OoAccComponent comp)
        {
            return GetRegisteredShapeObserver(comp.AccCont);
        }

        /// <summary>
        /// Gets the registered shape observer.
        /// </summary>
        /// <param name="comp">The comp.</param>
        /// <returns></returns>
        internal OoShapeObserver GetRegisteredShapeObserver(unoidl.com.sun.star.accessibility.XAccessibleComponent comp)
        {
            return GetRegisteredShapeObserver(comp as unoidl.com.sun.star.accessibility.XAccessibleContext);
        }

        private OoShapeObserver getRegisteredShapeObserver(String name)
        {
            if (shapes.ContainsKey(name))
            {
                return shapes[name];
            }
            return null;
        }

        private OoShapeObserver searchRegisterdSapeObserverByAccessibleName(string accName, XAccessibleContext cont, bool searchUnregistered = true)
        {
            if (accshapes.ContainsKey(cont))
                return accshapes[cont];

            // if the name contains spaces - e.g. for text farmes etc.
            if (accName.Contains(" "))
            {
                System.Diagnostics.Debug.WriteLine("problem in finding the correct shape - name contains space characters");
                var tokens = accName.Split(' ');
                string sName = String.Empty;
                foreach (var token in tokens)
                {
                    sName += (String.IsNullOrEmpty(sName) ? "" : " ") + token;
                    OoShapeObserver _so = getRegisteredShapeObserver(sName);
                    if (AccCorrespondsToShapeObserver(cont as XAccessible, _so))
                    {
                        return _so;
                    }
                }
            }

            // if name is already registered
            OoShapeObserver so = getRegisteredShapeObserver(accName);
            if (AccCorrespondsToShapeObserver(cont as XAccessible, so))
                return so;

            // search through all registered shapes
            so = searchRegisteredShapeByAccessible(accName, cont as XAccessible);
            if (AccCorrespondsToShapeObserver(cont as XAccessible, so))
                return so;


            if (searchUnregistered) so = searchUnregisteredShapeByAccessible(accName, cont);

            return null;
        }

        /// <summary>
        /// Check if the accessible object is corresponding to the  given shape observer.
        /// </summary>
        /// <param name="acc">The accessible object to test.</param>
        /// <param name="sObs">The shape observer to test.</param>
        /// <returns><c>true</c> if the acc is related to changes in the observer. If this is true, 
        /// the AccessibleCounterpart field of the OoShapeObserver is updated as well</returns>
        internal static bool AccCorrespondsToShapeObserver(XAccessible acc, OoShapeObserver sObs)
        {
            bool result = false;
            if (acc != null && sObs != null)
            {
                String hash = sObs.GetHashCode().ToString() + "_";
                if (sObs.IsText)
                {
                    // use description
                    string oldDescription = sObs.Description;
                    sObs.Description = hash + oldDescription;
                    if (OoAccessibility.GetAccessibleDesc(acc).StartsWith(hash))
                    {
                        result = true;
                    }
                    sObs.Description = oldDescription;
                }
                else
                {
                    // use name
                    string oldName = sObs.Name;
                    sObs.Name = hash + oldName;
                    if (OoAccessibility.GetAccessibleName(acc).StartsWith(hash))
                    {
                        result = true;
                    }
                    sObs.Name = oldName;
                }
            }
            if (result)
            {
                sObs.AcccessibleCounterpart = acc;
            }
            return result;
        }

        //TODO: unfinished
        private OoShapeObserver searchUnregisteredShapeByAccessible(string accName, XAccessibleContext cont)
        {
            Update();
            var obs = searchRegisterdSapeObserverByAccessibleName(accName, cont, false);

            if (obs != null)
                return obs;

            XDrawPages pages = PagesSupplier.getDrawPages();
            if (pages != null)
            {
                try
                {
                    for (int i = 0; i < pages.getCount(); i++)
                    {
                        var anyPage = pages.getByIndex(i);
                        if (anyPage.hasValue())
                        {
                            var page = anyPage.Value;
                            if (page != null && page is XShapes)
                            {
                                for (int j = 0; j < ((XShapes)page).getCount(); j++)
                                {
                                    var anyShape = ((XShapes)page).getByIndex(j);
                                    if (anyShape.hasValue())
                                    {
                                        var shape = anyShape.Value;
                                        string name = OoUtils.GetStringProperty(shape, "Name");
                                        string UIname = OoUtils.GetStringProperty(shape, "UINameSingular");
                                        string Title = OoUtils.GetStringProperty(shape, "Title");


                                        if (accName.Equals(UIname))
                                        {

                                        }
                                        else if (accName.Equals(name))
                                        {

                                        }
                                        else if (accName.Equals(name + " " + Title))
                                        {

                                        }
                                        else
                                        {

                                        }

                                    }
                                }
                            }
                        }
                    }
                }
                catch (System.Exception){}
            }
            return null;
        }

        private OoShapeObserver searchRegisteredShapeByAccessible(String accName, XAccessible acc)
        {
            List<String> names = new List<String>();

            foreach (KeyValuePair<String, OoShapeObserver> shapeKeyValuePair in shapes)
            {
                OoShapeObserver _so = shapeKeyValuePair.Value;

                if(_so != null && !_so.Disposed)
                {
                    if (_so.AcccessibleCounterpart != null)
                    {
                        if (_so.AcccessibleCounterpart.Equals(acc))
                        {
                            return _so;
                        }
                    }
                    else
                    {
                        if (AccCorrespondsToShapeObserver(acc, _so))
                        {
                            return _so;
                        }
                    }
                }
            }
            return null;
        }

        private void renewRegistrationOfShape(String key, String newKey, OoShapeObserver _so, XAccessible acc)
        {
            OoShapeObserver _trash;
            shapes.TryRemove(key, out _trash);
            if (!shapes.ContainsKey(newKey)) { shapes[newKey] = _so; }
            else
            {
                //TODO: renew the name to a unique one?
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "Shape with name '" + newKey + "' already registered. Cannot add this shape to list of Shapes");
            }
        }

        #endregion

        #region XAccessible Counterpart

        private void setAccessibleCounterpart(OoShapeObserver shape)
        {
            if (shape != null)
            {
                string oldName = shape.Name;
                string newName = shape.Name + "_" + shape.GetHashCode() + " ";
                shape.Name = newName;
                unoidl.com.sun.star.accessibility.XAccessible counterpart = null;
                if (shape.Parent != null && shape.Parent.AcccessibleCounterpart != null)
                {
                    counterpart = OoAccessibility.GetAccessibleCounterpartFromHash(shape.Shape, shape.Parent.AcccessibleCounterpart as unoidl.com.sun.star.accessibility.XAccessibleContext);
                }
                else
                {
                    counterpart = OoAccessibility.GetAccessibleCounterpartFromHash(shape.Shape, Document.AccComp as unoidl.com.sun.star.accessibility.XAccessibleContext);
                }
                if (counterpart != null) shape.AcccessibleCounterpart = counterpart;

                shape.Name = oldName;
            }
        }

        #endregion

        #endregion

        #region XAccessibleEventListener

        protected override void notifyEvent(AccessibleEventObject aEvent)
        {
            var id = OoAccessibility.GetAccessibleEventIdFromShort(aEvent.EventId);

            switch (id)
            {
                case tud.mci.tangram.Accessibility.AccessibleEventId.CHILD:
                    handleChild(aEvent.Source, aEvent.NewValue, aEvent.OldValue);
                    break;
                default:
                    break;
            }
        }

        #endregion

        #region XModifyListener - makes it possible to receive events when a model object changes.

        //happens on child adding, deleting, title changing, name changing 
        // the draw page supplier will be returned not the modified object :-(
        protected override void modified(unoidl.com.sun.star.lang.EventObject aEvent)
        {
            if (aEvent != null && aEvent.Source == PagesSupplier)
            {
                //System.Diagnostics.Debug.WriteLine("\t\tModify event happened in DrawPageObserver");
            }
        }

        #endregion

        #region XPropertyChangeListener

        /// <summary>
        /// Gets the current zoom value of this document in percent.
        /// </summary>
        /// <value>
        /// The zoom value.
        /// </value>
        public int ZoomValue { get; private set; }
        /// <summary>
        /// Gets the view offset.
        /// Defines the offset from the top left position of the displayed page to the 
        /// top left position of the view area in 100th/mm.
        /// </summary>
        /// <value>
        /// The view offset.
        /// </value>
        public System.Drawing.Point ViewOffset { get; private set; }

        /// <summary>
        /// Gets the horizontal resolution in pixel per meter.
        /// </summary>
        /// <value>
        /// The pixel per meter in x direction.
        /// </value>
        public static double PixelPerMeterX { get; private set; }
        /// <summary>
        /// Gets the vertical resolution in pixel per meter.
        /// </summary>
        /// <value>
        /// The pixel per meter in y direction.
        /// </value>
        public static double PixelPerMeterY { get; private set; }

        private void addVisibleAreaPropertyChangeListener()
        {
            // property listeners for zoom and view offset 
            if (this.PagesSupplier != null && this.PagesSupplier is XModel)
            {
                //var controller = ((XModel)this.PagesSupplier).getCurrentController();
                if (Controller != null && Controller is XPropertySet)
                {
                    // It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
                    try
                    {
                        ((XPropertySet)Controller).removePropertyChangeListener("VisibleArea", eventForwarder);
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Can't remove property listener for 'VisibleArea'", ex);
                    }
                    try
                    {
                        ((XPropertySet)Controller).addPropertyChangeListener("VisibleArea", eventForwarder);
                    }
                    catch (unoidl.com.sun.star.uno.RuntimeException) { Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Can't add property listener for 'VisibleArea' - Listener already registered"); }
                    catch (Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Can't add property listener for 'VisibleArea' " + ex);
                    }
                }
            }
        }

        private void removeVisibleAreaPropertyChangeListener()
        {
            // property listeners for zoom and view offset 
            if (this.PagesSupplier != null && this.PagesSupplier is XModel)
            {
                //var controller = ((XModel)this.PagesSupplier).getCurrentController();
                if (Controller != null && Controller is XPropertySet)
                {

                    try
                    {
                        // It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
                        ((XPropertySet)Controller).removePropertyChangeListener("VisibleArea", eventForwarder);
                    }
                    catch (Exception){}
                }
            }
        }

        /// <summary>
        /// implementing XPropertyChangeListener method
        /// </summary>
        /// <param name="evt">The evt.</param>
        protected override void propertyChange(PropertyChangeEvent evt)
        {
            if (evt.Source == ((XModel)this.PagesSupplier).getCurrentController())
            {
                if (evt.PropertyName.Equals("VisibleArea"))
                {
                    refreshDrawViewProperties((XPropertySet)evt.Source);
                    OnViewOrZoomChangeEvent();
                }
            }
        }

        /// <summary>
        /// get current ZoomValue and view ViewOffset properties
        /// </summary>
        /// <param name="drawViewProperties">The draw view properties.</param>
        private void refreshDrawViewProperties(XPropertySet drawViewProperties)
        {
            if (drawViewProperties == null) return;
            // get the current zoom factor in percent, as displayed in openoffice
            ZoomValue = (System.Int16)drawViewProperties.getPropertyValue("ZoomValue").Value;
            // get the view offset: from the top left position of the displayed page to the top left position of the view area in 100th/mm
            unoidl.com.sun.star.awt.Point vOffset = drawViewProperties.getPropertyValue("ViewOffset").Value as unoidl.com.sun.star.awt.Point;
            if (vOffset != null)
            {
                ViewOffset = new System.Drawing.Point(vOffset.X, vOffset.Y);
            }

        }

        private void OnViewOrZoomChangeEvent()
        {
            if (ViewOrZoomChangeEventHandlers != null) ViewOrZoomChangeEventHandlers();
        }

        /// <summary>
        /// delegate for Bounding Box change event handling
        /// </summary>
        public delegate void ViewOrZoomChangeEventHandler();
        /// <summary>
        /// Occurs when view or zoom changed.
        /// </summary>
        public event ViewOrZoomChangeEventHandler ViewOrZoomChangeEventHandlers;

        #endregion

        #region Children

        //TODO check the child an build a off screen model

        private void handleChild(object sender, uno.Any newValue, uno.Any oldValue)
        {

            System.Diagnostics.Debug.WriteLine("[INFO] OoDrawPagesObserver: handle accessible child ID");

            if (newValue.hasValue())
            {
                if (oldValue.hasValue()) // change
                {
                    //System.Diagnostics.Debug.WriteLine("child changed");
                }
                else // added
                {
                    addNewShape(newValue.Value as XAccessible);
                }

            }
            else if (oldValue.hasValue() && oldValue.Value is XAccessible)
            {

                // FIXME: this does not work properly because sometimes the page itself do this and there is no way to check if this is the page.
                //          So we have to leave the corresponding shape and its observer as a dead body in the lists ...
                //          Maybe the OoShapeObeservers can detect their disposing by their own and handle a clean up of the lists

                //// child deleted;
                ////System.Diagnostics.Debug.WriteLine("child deleted");

                //XAccessible oldValAcc = oldValue.Value as XAccessible;

                //if (oldValAcc != null)
                //{
                //    if (accshapes.ContainsKey((oldValAcc.getAccessibleContext())))
                //    {
                //        //TODO: remove ....
                //        System.Diagnostics.Debug.WriteLine("have to remove observer");
                //    }
                //    else
                //    {
                //        string Name = OoAccessibility.GetAccessibleName(oldValAcc);
                //        if (shapes.ContainsKey(Name))
                //        {
                //            removeChild(shapes[Name]);
                //        }
                //    }
                //}

            }
        }

        /// <summary>
        /// Removes the child and his children from the shapes list.
        /// </summary>
        /// <param name="sObs">The s obs.</param>
        void removeChild(OoShapeObserver sObs)
        {
            if (sObs != null && shapes.ContainsKey(sObs.Name))
            {
                OoShapeObserver trash;
                shapes.TryRemove(sObs.Name, out trash);

                String trashName = sObs.Name;
                shapeIds.TryTake(out trashName);

                if (trash != null)
                {
                    foreach (var item in trash.GetChilderen())
                    {
                        removeChild(item);
                    }
                }
            }
        }

        void addNewShape(XAccessible acc)
        {
            if (acc != null)
            {
                foreach (var item in DrawPageobservers)
                {
                    item.Value.Update();
                }
            }
        }

        internal void UpdateObserverLists(OoShapeObserver obs)
        {
            String name = obs.Name;
            XShape shape = obs.Shape;
            XAccessible acc = obs.AcccessibleCounterpart;
            XAccessibleContext cont = acc != null ? acc.getAccessibleContext() : null;

            //TODO: maybe do this softer?!

            try { shapes[name] = obs; }
            catch (Exception) { }
            if (cont != null)
            {
                try { accshapes[cont] = obs; }
                catch (Exception) { }
            }
            else
            {
                // search for a relation?
            }
            try { domshapes[shape] = obs; }
            catch (Exception) { }

        }

        /// <summary>
        /// Handles the ObserverDisposing event of a shape observer.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        void shape_ObserverDisposing(object sender, EventArgs e)
        {
            if (sender != null && sender is OoShapeObserver)
            {
                OoShapeObserver sO = sender as OoShapeObserver;
                string name = sO.Name;
                Logger.Instance.Log(LogPriority.OFTEN, this, "OoShapeObserver disposed : " + name);

                OoShapeObserver trash;
                try
                {
                    if (shapes.TryRemove(name, out trash))
                    {
                        if (trash != sO)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Deleted OoShapeObserver dos not match the observer to delete (by Name)");
                        }
                    }
                }
                catch { }
                try
                {
                    if (sO.AccComponent.AccCont != null && accshapes.TryRemove(sO.AccComponent.AccCont, out trash))
                    {
                        if (trash != sO)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Deleted OoShapeObserver dos not match the observer to delete (by Accessible)");
                        }
                    }
                }
                catch { }
                try
                {
                    if (sO.Shape != null && domshapes.TryRemove(sO.Shape, out trash))
                    {
                        if (trash != sO)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Deleted OoShapeObserver dos not match the observer to delete (by XShape)");
                        }
                    }
                }
                catch { }
            }
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
