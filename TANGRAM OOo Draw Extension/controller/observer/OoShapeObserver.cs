﻿using System;
using System.Threading;
using System.Threading.Tasks;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.document;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.io;
using unoidl.com.sun.star.lang;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Observes a XShape or an XShapes and his children
    /// </summary>
    public partial class OoShapeObserver : PropertiesEventForwarderBase, IUpdateable, IDisposable, IDisposingObserver, INameBuilder
    {
        #region Members

        /// <summary>
        /// The corresponding XShaoe DOM-Object
        /// </summary>
        internal XShape Shape { get; private set; }
        /// <summary>
        /// returns the anonymous XShape object from the DOM.
        /// </summary>
        public Object DomShape { get { return Shape as Object; } }

        private XAccessible _acccessibleCounterpart;
        /// <summary>
        /// Gets or sets the accessible counterpart to the DOM-object.
        /// </summary>
        /// <value>The accessible counterpart.</value>
        internal XAccessible AcccessibleCounterpart
        {
            get { return _acccessibleCounterpart; }
            set
            {
                _acccessibleCounterpart = value;
                if (_acccessibleCounterpart != null && Page != null && Page.PagesObserver != null)
                {
                    Page.PagesObserver.UpdateObserverLists(this);
                    registerAccessibleEvents();
                }
            }
        }

        /// <summary>
        /// Gets or sets the accessible component - a wrapper for the OO-api.
        /// </summary>
        /// <value>The accessible component.</value>
        public OoAccComponent AccComponent
        {
            get { return new OoAccComponent(AcccessibleCounterpart); }
            set { AcccessibleCounterpart = value.AccComp as XAccessible; }
        }
        /// <summary>
        /// The observer for the page the shape is located on
        /// </summary>
        public readonly OoDrawPageObserver Page;
        /// <summary>
        /// the observer for the parent shape if available
        /// </summary>
        public OoShapeObserver Parent { get; private set; }

        ///// <summary>
        ///// a list of observers for the children - the order must not corresponding to the order in the DOM
        ///// </summary>
        //public ConcurrentBag<OoShapeObserver> Children = new ConcurrentBag<OoShapeObserver>();


        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoShapeObserver"/> class.
        /// </summary>
        /// <param name="s">The XShape to observe.</param>
        /// <param name="page">The observer for the page the shape is located on.</param>
        public OoShapeObserver(XShape s, OoDrawPageObserver page) : this(s, page, null) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="OoShapeObserver"/> class.
        /// </summary>
        /// <param name="s">The XShape to observe.</param>
        /// <param name="page">The observer for the page the shape is located on.</param>
        /// <param name="parent">The observer for the parent shape.</param>
        public OoShapeObserver(XShape s, OoDrawPageObserver page, OoShapeObserver parent)
            : base()
        {

            Shape = s;
            Page = page;
            Parent = parent;
            if (Shape != null)
            {
                //if (IsGroup)
                //{
                //    util.Debug.GetAllInterfacesOfObject(Shape);
                //}

                if (Page != null && Page.PagesObserver != null) Page.PagesObserver.RegisterUniqueShape(this);

                handleChildren();
                registerListeners();

            }

            registerShapeOnPageShapeList();
        }

        private void registerShapeOnPageShapeList()
        {
            // register this shape to the list of elements of the page
            if (Page != null && Page.shapeList != null)
            {
                try
                {
                    //Page.shapeList.RemoveAll((sibling) =>
                    //{
                    //    return sibling.Disposed 
                    //        || sibling.Shape == Shape
                    //        || sibling.Shape == null
                    //        || sibling.Name == Name
                    //        || String.IsNullOrWhiteSpace(sibling.Name)
                    //        || !sibling.IsValid();                        
                    //});

                    Page.shapeList.Remove(this);

                    Page.shapeList.Add(this);

                }
                catch (Exception) { }
            }
        }

        private void registerListeners()
        {
            // register for property changes
            XPropertySet shapePropertySet = (XPropertySet)Shape;
            if (shapePropertySet != null)
            {

                // init value for bounding box
                Rectangle propertyBoundRect = (Rectangle)shapePropertySet.getPropertyValue("BoundRect").Value;

                currentBoundRect = propertyBoundRect != null ?
                    new System.Drawing.Rectangle(propertyBoundRect.X, propertyBoundRect.Y, propertyBoundRect.Width, propertyBoundRect.Height)
                    : new System.Drawing.Rectangle();

                try
                {
                    shapePropertySet.removePropertyChangeListener("Size", eventForwarder);
                    shapePropertySet.removePropertyChangeListener("Position", eventForwarder);
                    shapePropertySet.addPropertyChangeListener("Size", eventForwarder);
                    shapePropertySet.addPropertyChangeListener("Position", eventForwarder);
                }
                catch { }
            }

            if (Shape is XEventBroadcaster)
            {
                try
                {
                    ((XEventBroadcaster)Shape).addEventListener(eventForwarder as unoidl.com.sun.star.document.XEventListener);
                }
                catch { }
            }
        }
        private void registerAccessibleEvents()
        {
            util.Debug.GetAllInterfacesOfObject(AcccessibleCounterpart);
            if (AcccessibleCounterpart is XAccessibleEventBroadcaster)
            {
                try
                {
                    ((XAccessibleEventBroadcaster)AcccessibleCounterpart).addAccessibleEventListener(eventForwarder);
                }
                catch { }
            }
        }

        #endregion

        #region Property Listeners

        #region XPropertyChangeListener

        private System.Drawing.Rectangle currentBoundRect;
        protected override void propertyChange(PropertyChangeEvent evt)
        {
            //System.Diagnostics.Debug.WriteLine("property changed: " + evt.PropertyName);
            if (evt.Source == Shape &&
                (evt.PropertyName.Equals("Position") || evt.PropertyName.Equals("Size")) &&
                Shape is XPropertySet)
            {
                XPropertySet shapePropertySet = (XPropertySet)Shape;
                {
                    // Rectangle propertyBoundRect = (Rectangle)shapePropertySet.getPropertyValue("BoundRect").Value;
                    Rectangle propertyBoundRect = (Rectangle)shapePropertySet.getPropertyValue("FrameRect").Value;
                    currentBoundRect.X = propertyBoundRect.X;
                    currentBoundRect.Y = propertyBoundRect.Y;
                    currentBoundRect.Width = propertyBoundRect.Width;
                    currentBoundRect.Height = propertyBoundRect.Height;

                    OnBoundRectChangeEvent();
                }
            }
        }

        private void OnBoundRectChangeEvent()
        {
            if (BoundRectChangeEventHandlers != null) BoundRectChangeEventHandlers();
        }

        // delegate for Bounding Box change event handling
        public delegate void BoundRectChangeEventHandler();
        // event object
        public event BoundRectChangeEventHandler BoundRectChangeEventHandlers;

        #endregion

        #region XVetoableChangeListener

        protected override void vetoableChange(PropertyChangeEvent aEvent)
        {
            //TODO: check if the new name is unique
        }

        #endregion

        #region XPropertiesChangeListener
        protected override void propertiesChange(PropertyChangeEvent[] aEvent) { }
        #endregion

        protected override void notifyEvent(AccessibleEventObject aEvent)
        {
            if (aEvent != null)
            {
                System.Diagnostics.Debug.WriteLine("-----> Event in ShapeObserver " + aEvent.EventId);
            }
        }

        #endregion

        #region General Object Overwrites

        // override object.Equals
        public override bool Equals(object obj)
        {
            //       
            // See the full list of guidelines at
            //   http://go.microsoft.com/fwlink/?LinkID=85237  
            // and also the guidance for operator== at
            //   http://go.microsoft.com/fwlink/?LinkId=85238
            //

            if (obj == null || GetType() != obj.GetType())
            {
                if (obj is XShape)
                {
                    return OoUtils.AreOoObjectsEqual(Shape, obj);
                }
                return false;
            }

            if (base.Equals(obj)) return true;

            try
            {
                return OoUtils.AreOoObjectsEqual(Shape, ((OoShapeObserver)obj).Shape);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        #endregion

        #region IUpdateable

        bool isUpdating = false;
        DateTime lastUpdateTime = new DateTime(1970, 1, 1);
        virtual public void Update()
        {
            if (isUpdating || (DateTime.Now - lastUpdateTime).TotalMilliseconds < 100)
                return;

            lock (SynchLock)
            {
                try
                {
                    isUpdating = true;
                    if (!IsValid(true)) return;

                    //TODO: update informations
                    //throw new NotImplementedException();
                    if (AcccessibleCounterpart == null)
                    {
                        if (Page != null && Page.PagesObserver != null && Page.PagesObserver.Document != null)
                        {
                            AcccessibleCounterpart = OoAccessibility.GetAccessibleCounterpartFromHash(Shape, Page.PagesObserver.Document.AccCont);
                        }
                        //{
                        //    var childs = OoAccessibility.GetAllChildrenOfAccessibleObject(Page.PagesObserver.Document.AccCont as XAccessible);

                        //    string current_name = String.Empty;
                        //    bool success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(1000, () => { current_name = Name; });
                        //    if (String.IsNullOrEmpty(current_name))
                        //    {
                        //        isUpdating = false;
                        //        lastUpdateTime = DateTime.Now;
                        //        return;
                        //    }

                        //    // foreach (var c in childs) { System.Diagnostics.Debug.WriteLine(OoAccessibility.GetAccessibleNamePart(c as XAccessible)); }

                        //    foreach (var child in childs)
                        //    {
                        //        if (child is XAccessible)
                        //        {
                        //            if (OoUtils.ElementSupportsService(child, OO.Services.DRAWING_ACCESSIBLE_SHAPE))
                        //            {
                        //                if (child is XAccessibleComponent)
                        //                {
                        //                    var size = ((XAccessibleComponent)child).getSize();
                        //                    if (size.Width <= 0 || size.Height <= 0)
                        //                        continue;
                        //                }
                        //                // get AccessibleName leads to hang ons
                        //                var name = OoAccessibility.GetAccessibleNamePart(child as XAccessible);
                        //                if (name.Equals(current_name))
                        //                {
                        //                    // TODO: check other/better possibilities

                        //                    try // change name and check if the name is changed in the accessible view too
                        //                    {
                        //                        if (Shape != null)
                        //                        {
                        //                            string oldName = OoUtils.GetStringProperty(Shape, "Name");
                        //                            string newName = this.GetHashCode() + " ";
                        //                            OoUtils.SetStringProperty(Shape, "Name", newName);
                        //                            System.Threading.Thread.Sleep(5);
                        //                            string accName = OoAccessibility.GetAccessibleName(child as XAccessible);
                        //                            if (accName.StartsWith(newName))
                        //                            {
                        //                                AcccessibleCounterpart = child;
                        //                                OoUtils.SetStringProperty(Shape, "Name", oldName);
                        //                                break;
                        //                            }
                        //                            else
                        //                            {
                        //                                OoUtils.SetStringProperty(Shape, "Name", oldName);
                        //                                continue;
                        //                            }
                        //                        }

                        //                    }
                        //                    catch (Exception)
                        //                    {
                        //                    }

                        //                    AcccessibleCounterpart = child;
                        //                    break;
                        //                }
                        //            }
                        //        }
                        //    }
                        //}
                    }

                    if (_ppObs != null) { _ppObs.Update(); } // update polypolygon points
                }
                catch { }
                finally
                {
                    // TODO: update children
                    isUpdating = false;
                    lastUpdateTime = DateTime.Now;
                }
            }
        }

        #endregion

        #region Validation

        /// <summary>
        /// A flag for disabling the validation request. The request for validation collides sometimes with the 
        /// change of properties such as polygon points.
        /// </summary>
        internal static volatile bool LockValidation = false;
        /// <summary>
        /// The lock object to synchronize parallel access to this object. 
        /// </summary>
        public static Object SynchLock = new Object();

        bool _valid = true;
        private const long _validationTreshold = 80000000; // 10.000.000 ~ 1 sec.
        long _lastValidation = DateTime.UtcNow.Ticks;

        /// <summary>
        /// Determines whether this instance is valid.
        /// </summary>
        /// <param name="force">forces a revalidation; otherwise a revalidation is only applied every ~5 sec. to save performance - in such a case the last validation value will be returned. </param>
        /// <returns>
        ///   <c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </returns>
        /// <remarks>Parts of this function are time limited to 100 ms.</remarks>
        virtual public bool IsValid(bool force = true)
        {
            if (!_valid) return false;
            if (LockValidation) return _valid;

            long ts = DateTime.UtcNow.Ticks;
            if (!force && (ts - _lastValidation < _validationTreshold)) return _valid;

            _valid = true;
            lock (SynchLock)
            {
                try
                {
                    if (!Disposed && Shape != null && Page != null)
                    {
                        TimeLimitExecutor.WaitForExecuteWithTimeLimit(100, () =>
                        {
                            try
                            {

                                var test1 = Shape.GetHashCode();
                                string uiName = UINameSingular;
                                if (test1 == 0 || String.IsNullOrEmpty(uiName))
                                {
                                    _valid = false;
                                    Dispose();
                                    return;
                                }

                                if (AcccessibleCounterpart != null) AcccessibleCounterpart.GetHashCode();
                                else
                                {
                                    Update();
                                }

                                if (!OoUtils.ElementSupportsService(Shape, OO.Services.DRAW_SHAPE))
                                {
                                    _valid = false;
                                    return;
                                }

                                //var services = util.Debug.GetAllServicesOfObject(Shape, false);
                                //if (services.Length == 0)
                                //{
                                //    valid = false;
                                //    return;
                                //}

                                //// test if a parent exists
                                //XShapes parent;
                                //bool parentSuccess = tryGetParentByXShape(out parent, Shape);
                                //if (parentSuccess && parent == null)
                                //{
                                //    valid = false;
                                //    // should bee disposed?!
                                //    Dispose();
                                //    return;
                                //}

                                _valid = true;

                            }
                            catch (System.Threading.ThreadAbortException) { DisableValidationTemporary(); }
                            catch (System.Threading.ThreadInterruptedException) { DisableValidationTemporary(); }
                        }, "ShapeValidation Shape " + _lastName);
                    }
                    else
                    {
                        _valid = false;
                    }
                }
                catch (unoidl.com.sun.star.lang.DisposedException)
                {
                    Dispose();
                }
                catch { }
            }

            _lastValidation = ts;
            return _valid;
        }

        /// <summary>
        /// Disables the validation of all shape observers for a short term.
        /// This is to prevent deadlocks caused by many parallel requests or by 
        /// massive changes of Shape or page properties.
        /// </summary>
        internal static void DisableValidationTemporary()
        {
            Task t = new Task(new Action(() =>
            {
                try
                {
                    OoShapeObserver.LockValidation = true;
                    Thread.Sleep(100);
                }
                catch { }
                finally { OoShapeObserver.LockValidation = false; }
            }));
        }

        /// <summary>
        /// Determines whether this instance is visible on the screen.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance is visible; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool IsVisible()
        {
            try
            {
                if (IsValid(false))
                {
                    var s = Size;

                    if (Visible && !GetAbsoluteScreenBoundsByDom().IsEmpty)
                        return true;

                    if (AcccessibleCounterpart != null && s.Height > 0 && s.Width > 0)
                    {
                        return true;
                    }
                }
            }
            catch (unoidl.com.sun.star.lang.DisposedException) { Dispose(); }
            catch { }
            return false;
        }

        #endregion

        #region IDispose

        virtual public void Dispose()
        {
            if (this.Shape != null)
            {
                try
                {
                    // register for property changes
                    XPropertySet shapePropertySet = (XPropertySet)Shape;
                    if (shapePropertySet != null)
                    {
                        shapePropertySet.removePropertyChangeListener("Size", eventForwarder);
                        shapePropertySet.removePropertyChangeListener("Position", eventForwarder);
                    }
                }
                catch { }
            }

            //this.Shape = null;
            //this.AcccessibleCounterpart = null;

            // remove from the pages' shape list
            if (Page != null && Page.shapeList != null && Page.shapeList.Contains(this))
            {
                Page.shapeList.Remove(this);
            }

            Disposed = true;
            fire_DisposingEvent();
        }
        #endregion

        #region Helper

        /// <summary>
        /// Gets the document (SERVICE com.sun.star.document.OfficeDocument) this 
        /// XShape is related to.
        /// </summary>
        /// <returns>the document object or null</returns>
        internal object GetDocument()
        {
            var p = this.Page;
            if (p != null)
            {
                var dps = p.PagesObserver;
                if (dps != null)
                {
                    var doc = dps.PagesSupplier;
                    if (doc != null)
                    {
                        if (util.OoUtils.ElementSupportsService(doc, OO.Services.DOCUMENT))
                        {
                            return doc;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// To get the shape (may also be a group...) as a translucent png in a byte array.
        /// </summary>
        /// <param name="pngFileData">The png file data as byte array.</value>
        /// <returns>the png file size (only > 0 if successful)</returns>
        virtual public int GetShapeAsPng(out byte[] pngFileData)
        {
            pngFileData = new byte[0];
            object doc = GetDocument();
            // see https://blog.oio.de/2010/10/27/copy-and-paste-without-clipboard-using-openoffice-org-api/
            if (IsValid(false) && doc != null && Shape != null)
            {
                System.Drawing.Rectangle rectDom = GetRelativeScreenBoundsByDom();

                // See <http://www.oooforum.org/forum/viewtopic.phtml?t=50783> 
                // properties for the png export filter:
                PropertyValue[] aFilterData = new PropertyValue[5];
                aFilterData[0] = new PropertyValue
                {
                    Name = "PixelWidth",
                    Value = tud.mci.tangram.models.Any.Get(rectDom.Width) // bounding box width
                };
                aFilterData[1] = new PropertyValue
                {
                    Name = "PixelHeight",
                    Value = tud.mci.tangram.models.Any.Get(rectDom.Height)// bounding box height
                };
                aFilterData[2] = new PropertyValue
                {
                    Name = "Translucent",
                    Value = new uno.Any(true)    // png with translucent background
                };
                aFilterData[3] = new PropertyValue
                {
                    Name = "Compression",
                    Value = new uno.Any(0)    // png compression level 0..9 (set to 0 to be fastest, smallest files would be produced by 9)
                };
                aFilterData[4] = new PropertyValue
                {
                    Name = "Interlaced",
                    Value = new uno.Any(0)    // png interlacing (off=0)
                };
                /* create com.sun.star.comp.MemoryStream, Debug.GetAllInterfacesOfObject(memoryOutStream) gets:
                    unoidl.com.sun.star.io.XStream
                    unoidl.com.sun.star.io.XSeekableInputStream
                    unoidl.com.sun.star.io.XOutputStream
                    unoidl.com.sun.star.io.XTruncate
                    unoidl.com.sun.star.lang.XTypeProvider
                    unoidl.com.sun.star.uno.XWeak
                 */
                //XMultiServiceFactory xmsf = OO.GetMultiServiceFactory(OO.GetContext(), OO.GetMultiComponentFactory(OO.GetContext()));
                XStream memoryStream = (XStream)OO.GetContext().getServiceManager().createInstanceWithContext("com.sun.star.comp.MemoryStream", OO.GetContext());
                //XStream memoryStream = (XStream)xmsf.createInstance("com.sun.star.comp.MemoryStream");

                // the filter media descriptor: media type, destination, and filterdata containing the export settings from above
                PropertyValue[] aArgs = new PropertyValue[3];
                aArgs[0] = new PropertyValue
                {
                    Name = "MediaType",
                    Value = tud.mci.tangram.models.Any.Get("image/png")
                };
                aArgs[1] = new PropertyValue
                {
                    Name = "OutputStream",
                    Value = tud.mci.tangram.models.Any.Get(memoryStream)   // filter to our memory stream
                };
                aArgs[2] = new PropertyValue
                {
                    Name = "FilterData",
                    Value = tud.mci.tangram.models.Any.GetAsOne(aFilterData)
                };
                // create exporter service
                XExporter exporter = (XExporter)OO.GetContext().getServiceManager().createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", OO.GetContext());
                exporter.setSourceDocument((XComponent)Shape);

                // call the png export filter
                ((XFilter)exporter).filter(aArgs);

                // read all bytes from stream into byte array
                int pngFileSize = ((XInputStream)memoryStream).available();
                ((XInputStream)memoryStream).readBytes(out pngFileData, pngFileSize);
                return pngFileSize;
            }
            return 0;
        }

        /// <summary>
        /// tries to calculate the onscreen bounding box in px relative to the DRAW view area, based on dom position, openoffice zoom etc.
        /// the accessibility interface is not used by this method
        /// </summary>
        /// <returns>Bounding box in pixels as screen position relative to the DRAW view area.</returns>
        virtual public System.Drawing.Rectangle GetRelativeScreenBoundsByDom()
        {
            System.Drawing.Rectangle result = new System.Drawing.Rectangle(-1, -1, 0, 0);
            if (Shape != null && Page != null && Page.PagesObserver != null)
            {
                // get page properties like zoom and offset
                try
                {
                    // get the current zoom factor in percent, as displayed in openoffice
                    int zoomInPercent = Page.PagesObserver.ZoomValue;
                    // get the view offset: from the top left position of the displayed page to the top left position of the view area in 100th/mm
                    System.Drawing.Point viewOffset = Page.PagesObserver.ViewOffset;

                    // get dom bounds
                    //Rectangle domBounds = (Rectangle)GetProperty("BoundRect");
                    result = new System.Drawing.Rectangle(currentBoundRect.X, currentBoundRect.Y, currentBoundRect.Width, currentBoundRect.Height); //new Rectangle(domBounds.X, domBounds.Y, domBounds.Width, domBounds.Height);

                    // subtract page view offset
                    result.X = result.X - viewOffset.X;
                    result.Y = result.Y - viewOffset.Y;
                    // convert to pixel coords
                    result = util.OoDrawUtils.convertToPixel(result, zoomInPercent, OoDrawPagesObserver.PixelPerMeterX, OoDrawPagesObserver.PixelPerMeterY);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, ex);
                }
            }
            return result;
        }

        private System.Drawing.Rectangle _lastBounds = new System.Drawing.Rectangle();
        private volatile bool _updatingBounds = false;
        /// <summary>
        /// tries to calculate the absolute on screen bounding box in px of some given rectangle or if none is provided, based on the shape observer dom position, openoffice zoom etc.
        /// the accessibility interface is used for getting the controller main window position on the screen
        /// </summary>
        /// <param name="relativeBounds">the given relative bounds are used for transformation into absolute px coordinates.
        /// if the given value is null, getRelativeScreenBoundsByDom() of this shape observer will be called
        /// </param>
        /// <returns>Bounding box in pixels as absolute screen position.</returns>
        /// <remarks>Parts of this function are time limited to 100 ms.</remarks>
        virtual public System.Drawing.Rectangle GetAbsoluteScreenBoundsByDom(System.Drawing.Rectangle relativeBounds)
        {
            if (_updatingBounds) return _lastBounds;

            _updatingBounds = true;
            System.Drawing.Rectangle result = new System.Drawing.Rectangle();
            try
            {
                result = (relativeBounds.Height * relativeBounds.Width <= 0) ? GetRelativeScreenBoundsByDom() : relativeBounds;
                //XModel docModel = (XModel)GetDocument();
                if (result.Width * result.Height > 0
                    //&& docModel != null
                    )
                {
                    // get page properties like zoom and offset
                    //docModel.lockControllers();
                    try
                    {
                        TimeLimitExecutor.WaitForExecuteWithTimeLimit(100, () =>
                        {
                            try
                            {
                                // add offset for draw view
                                XAccessibleComponent xaCmp = (XAccessibleComponent)Page.PagesObserver.DocWnd.Document.getAccessibleContext();
                                Point drawViewWindowPos = (xaCmp != null) ? xaCmp.getLocationOnScreen() : new Point(0, 0);
                                result.X = result.X + drawViewWindowPos.X;
                                result.Y = result.Y + drawViewWindowPos.Y;
                            }
                            catch (System.Threading.ThreadAbortException) { }
                        }, "GetAbsoluteScreenBoundsByDom [" + Name + "]");
                    }
                    catch (System.Threading.ThreadAbortException) { }
                    catch (Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, this, ex);
                    }
                    finally
                    {
                        // docModel.unlockControllers();
                    }
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, this, "[NOTICE] can't get relative screen bounds from shape ");
                }

                _lastBounds = result;
            }
            finally
            {
                _updatingBounds = false;
            }
            return result;
        }

        /// <summary>
        /// tries to calculate the absolute on screen bounding box in px of some given rectangle or if none is provided, based on the shape observer dom position, openoffice zoom etc.
        /// the accessibility interface is used for getting the controller main window position on the screen
        /// </summary>
        /// <returns>
        /// Bounding box in pixels as absolute screen position.
        /// </returns>
        virtual public System.Drawing.Rectangle GetAbsoluteScreenBoundsByDom()
        {
            return GetAbsoluteScreenBoundsByDom(new System.Drawing.Rectangle(0, 0, 0, 0));
        }

        #endregion

        #region IDisposingObserver

        /// <summary>
        /// Gets a value indicating whether this <see cref="AbstractDisposingBase"/> is disposed.
        /// </summary>
        /// <value><c>true</c> if disposed; otherwise, <c>false</c>.</value>
        virtual public bool Disposed { get; private set; }

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

        #region INameBuilder

        string _uid = String.Empty;
        string _baseUid = String.Empty;
        /// <summary>
        /// Builds a unique, human-readable name for the object.
        /// </summary>
        /// <returns>
        /// A unique name for the object.
        /// </returns>
        string INameBuilder.BuildName()
        {
            _uid = Name;
            if (String.IsNullOrWhiteSpace(_uid))
            {
                _uid = UINameSingular;
            }
            _baseUid = _uid;
            return _uid;
        }

        int _tryies = 0;
        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <returns>
        /// A new try for a unique name for the object.
        /// </returns>
        string INameBuilder.RebuildName()
        {
            if (String.IsNullOrWhiteSpace(_baseUid))
                ((INameBuilder)this).BuildName();
            _uid = _baseUid + "_" + ++_tryies;
            return _uid;
        }

        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <param name="startIndex">The start index for the next try - e.g. 4 objects of the same type already exists.</param>
        /// <returns>
        /// A new try for a unique name for the object.
        /// </returns>
        string INameBuilder.RebuildName(int startIndex)
        {
            _tryies = startIndex;
            return ((INameBuilder)this).RebuildName();
        }

        #endregion

        #region DELETE

        /// <summary>
        /// Deletes the Shape from the DrawPage and disposes this instance.
        /// </summary>
        /// <returns><c>true</c> if the delete call was handled successfully.</returns>
        virtual public bool Delete()
        {
            bool success = false;
            try
            {
                lock (SynchLock)
                {
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[DELETE] shape: " + this.Name);

                    // FIXME: dont know why but sometimes the page is deleted as well with the object

                    if (this.Page != null)
                    {
                        var pObs = this.Page.PagesObserver;
                        if (pObs != null)
                        {
                            var contrl = pObs.Controller;
                            if (contrl != null && contrl is XDispatchProvider && contrl is unoidl.com.sun.star.view.XSelectionSupplier)
                            {
                                OoDispatchHelper.ActionWithChangeAndResetSelection(
                                    new Action(() =>
                                    {
                                        try
                                        {
                                            success = OoDispatchHelper.CallDispatch(
                                                DispatchURLs.SID_DELETE,
                                                contrl as XDispatchProvider
                                                );
                                        }
                                        catch (Exception ex) { Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] Can't finish DELET dispatch:", ex); }
                                    }),
                                    contrl as unoidl.com.sun.star.view.XSelectionSupplier,
                                    Shape
                                );
                                Thread.Sleep(100);                                
                                success = !IsValid();
                            }
                        }
                    }
                    return success;
                }
            }
            finally
            {
                // Dispose();
            }
        }

        #endregion

    }
}