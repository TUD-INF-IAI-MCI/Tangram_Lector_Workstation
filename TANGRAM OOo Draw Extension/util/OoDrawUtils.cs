using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using uno;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.view;
using unodrawing = unoidl.com.sun.star.drawing;

namespace tud.mci.tangram.util
{
    /*
     * OOoDrawUtils.java
     *
     * Created on February 19, 2003, 2:37 PM
     *
     * Copyright 2003 Danny Brewer
     * Anyone may run this code.
     * If you wish to modify or distribute this code, then
     *  you are granted a license to do so only under the terms
     *  of the Gnu Lesser General Public License.
     * See:  http://www.gnu.org/licenses/lgpl.html
     */

    /*
     * Highly modified by Jens Bornschein Copyright 2013 - 2016
     * under BSD license. 
     */

    /// <summary>
    /// General static helper functions for handling DRAW objects and documents
    /// @author jens bornschein, danny brewer 
    /// </summary>
    public static class OoDrawUtils
    {
        #region Get Draw PageSupplieirs
        /// <summary>
        /// Try to get all available XDrawPagesSuppliers.
        /// </summary>
        /// <param name="xContext">The XComponentContext.</param>
        /// <returns>A list of all available XDrawPagesSupplier</returns>
        public static List<XDrawPagesSupplier> GetDrawPageSuppliers(XComponentContext xContext)
        {
            //lock (_dpsLock)
            {
                return GetDrawPageSuppliers(OO.GetMultiComponentFactory(xContext), xContext);
            }
        }

        /// <summary>
        /// Try to get all available XDrawPagesSuppliers.
        /// </summary>
        /// <param name="factory">The XMultiComponentFactory.</param>
        /// <returns>A list of all available XDrawPagesSupplier</returns>
        public static List<XDrawPagesSupplier> GetDrawPageSuppliers(XMultiComponentFactory factory)
        {
            //lock (_dpsLock)
            {
                return GetDrawPageSuppliers(factory, OO.GetContext());
            }
        }
        /// <summary>
        /// Try to get all available XDrawPagesSuppliers.
        /// </summary>
        /// <param name="factory">The XMultiComponentFactory.</param>
        /// <param name="xContext">The XComponentContext.</param>
        /// <returns>A list of all available XDrawPagesSupplier</returns>
        public static List<XDrawPagesSupplier> GetDrawPageSuppliers(XMultiComponentFactory factory, XComponentContext xContext)
        {
            //lock (_dpsLock)
            {
                return GetDrawPageSuppliers(OO.GetDesktop(factory, xContext));
            }
        }

        private static readonly object _dpsLock = new Object();
        /// <summary>
        /// Try to get all available XDrawPagesSuppliers.
        /// </summary>
        /// <param name="xDesktop">The XDesktop.</param>
        /// <returns>A list of all available XDrawPagesSupplier</returns>
        public static List<XDrawPagesSupplier> GetDrawPageSuppliers(XDesktop xDesktop)
        {
            //lock (_dpsLock)
            {
                var dpsList = new List<XDrawPagesSupplier>();
                //cause by OpenOffice Version 3
                if (xDesktop == null)
                {
                    return dpsList;
                }

                XEnumeration enummeraration = null;

                TimeLimitExecutor.WaitForExecuteWithTimeLimit(200, () =>
                {
                    try
                    {
                        XEnumerationAccess xEnumerationAccess = xDesktop.getComponents();
                        enummeraration = xEnumerationAccess.createEnumeration();
                    }
                    catch { }
                }, "GetOjectsFromDesktop");

                while (enummeraration != null && enummeraration.hasMoreElements())
                {
                    Any anyelemet = enummeraration.nextElement();
                    XComponent element = (XComponent)anyelemet.Value;

                    if (element is XDrawPagesSupplier)
                    {
                        dpsList.Add(element as XDrawPagesSupplier);

                        XModel model = element as XModel;
                        if (model != null)
                        {
                            int i = 0;
                            //model.lockControllers();
                            try
                            {
                                while (!TimeLimitExecutor.WaitForExecuteWithTimeLimit
                                    (1000,
                                        () =>
                                        {
                                            var controller = model.getCurrentController();
                                            if (controller != null && controller is XSelectionSupplier)
                                            {
                                                tud.mci.tangram.controller.OoSelectionObserver.Instance.RegisterListenerToElement(controller);
                                                //try { ((XSelectionSupplier)controller).removeSelectionChangeListener(tud.mci.tangram.controller.OoSelectionObserver.Instance); }
                                                //catch { }
                                                //try { ((XSelectionSupplier)controller).addSelectionChangeListener(tud.mci.tangram.controller.OoSelectionObserver.Instance); }
                                                //catch
                                                //{
                                                //    try { ((XSelectionSupplier)controller).addSelectionChangeListener(tud.mci.tangram.controller.OoSelectionObserver.Instance); }
                                                //    catch { }
                                                //}
                                            }
                                        },
                                        "AddSelectionListenerToDrawPagesSupplier"
                                    )
                                    && i++ < 5) { Thread.Sleep(100); }
                            }
                            catch (System.Exception ex)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, "OoDrawUtils", "[ERROR] Can't add Selection listener to draw pages supplier: " + ex);
                            }
                            //model.unlockControllers();

                        }
                    }
                }
                return dpsList;
            }
        }
        #endregion

        #region DOM Tree functions

        internal static XDrawPage GetPageForShape(XShape shape)
        {
            XDrawPage page = null;
            if (shape != null && shape is XChild)
            {
                XShapes lastParent = null;
                XShapes par = GetParentShape(shape);
                while (true)
                {
                    if (par == null || lastParent == par) break;
                    //check if par is page shape
                    if (par is XDrawPage)
                    {
                        page = par as XDrawPage;
                        break;
                    }

                    lastParent = par;
                    par = GetParentShape(lastParent as XShape);
                }
            }
            return page;
        }

        internal static XShapes GetParentShape(XShape shape)
        {
            XShapes parent = null;
            if (shape != null && shape is XChild)
            {
                TimeLimitExecutor.WaitForExecuteWithTimeLimit(500, () =>
                {
                    parent = ((XChild)shape).getParent() as XShapes;
                }, "GetParentShape");
            }
            return parent;
        }

        #endregion

        //----------------------------------------------------------------------
        //  Sugar coated access to pages on a drawing document.
        //   The first page of a drawing is page zero.
        //----------------------------------------------------------------------

        #region Page Operations

        /// <summary>
        /// Gets the count of draw pages inside the draw doc. How many pages are on a drawing?
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <returns>Count of pages of the draw doc</returns>
        public static int GetNumDrawPages(Object drawDoc)
        {
            var drawDocXDrawPages = DrawDocGetXDrawPages(drawDoc);
            return drawDocXDrawPages != null ? drawDocXDrawPages.getCount() : 0;
        }
        /// <summary>
        /// Gets the count of draw pages.
        /// </summary>
        /// <param name="drawDocXDrawPages">The draw pages container.</param>
        /// <returns></returns>
        public static int GetNumDrawPages(XDrawPages drawDocXDrawPages)
        {
            return drawDocXDrawPages != null ? drawDocXDrawPages.getCount() : 0;
        }

        /// <summary>
        /// Obtain a page from a drawing.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="pageIndex">Index of the page.</param>
        /// <returns>the XDrawPage at the given index or null if not successful.</returns>
        public static XDrawPage GetDrawPageByIndex(Object drawDoc, int pageIndex)
        {
            var drawDocXDrawPages = DrawDocGetXDrawPages(drawDoc);
            return drawDocXDrawPages != null ? GetDrawPageByIndex(drawDocXDrawPages, pageIndex) : null;
        }
        /// <summary>
        /// Obtain a page from a drawing.
        /// </summary>
        /// <param name="drawDocXDrawPages">The draw doc X draw pages.</param>
        /// <param name="pageIndex">Index of the page.</param>
        /// <returns>the XDrawPage at the given index or null if not successful.</returns>
        public static XDrawPage GetDrawPageByIndex(XDrawPages drawDocXDrawPages, int pageIndex)
        {
            // Now ask the XIndexAccess interface to the drawPages object to get a certain page.
            Object drawPage = drawDocXDrawPages != null ? drawDocXDrawPages.getByIndex(pageIndex) : new Object();

            // Get the right interface to the page.
            var drawPageXDrawPage = drawPage as XDrawPage;

            return drawPageXDrawPage;
        }

        /// <summary>
        /// Create a new page on a drawing at the given index position.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="pageIndex">Index of the page.</param>
        /// <returns>the new page or null if not successful.</returns>
        public static XDrawPage InsertNewPageByIndex(Object drawDoc, int pageIndex)
        {
            var drawDocXDrawPages = DrawDocGetXDrawPages(drawDoc);
            var xDrawPage = drawDocXDrawPages != null ? drawDocXDrawPages.insertNewByIndex(pageIndex) : null;
            return xDrawPage;
        }
        /// <summary>
        /// Create a new page on a drawing at the given index position.
        /// </summary>
        /// <param name="drawDocXDrawPages">The draw doc X draw pages.</param>
        /// <param name="pageIndex">Index of the page.</param>
        /// <returns>the new page or null if not successful.</returns>
        public static XDrawPage InsertNewPageByIndex(XDrawPages drawDocXDrawPages, int pageIndex)
        {
            XDrawPage xDrawPage = drawDocXDrawPages != null ? drawDocXDrawPages.insertNewByIndex(pageIndex) : null;
            return xDrawPage;
        }

        /// <summary>
        /// Removes the page at the given index from the drawing.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="pageIndex">Index of the page.</param>
        public static void RemovePageFromDrawing(Object drawDoc, int pageIndex)
        {
            var drawDocXDrawPages = DrawDocGetXDrawPages(drawDoc);
            if (drawDocXDrawPages != null)
            {
                var xDrawPage = GetDrawPageByIndex(drawDocXDrawPages, pageIndex);
                if (xDrawPage != null)
                {
                    drawDocXDrawPages.remove(xDrawPage);
                }
            }
        }
        /// <summary>
        /// Removes the page at the given index from the drawing.
        /// </summary>
        /// <param name="drawDocXDrawPages">The draw doc X draw pages.</param>
        /// <param name="pageIndex">Index of the page.</param>
        public static void RemovePageFromDrawing(XDrawPages drawDocXDrawPages, int pageIndex)
        {
            var xDrawPage = GetDrawPageByIndex(drawDocXDrawPages, pageIndex);
            if (xDrawPage != null && drawDocXDrawPages != null)
            {
                drawDocXDrawPages.remove(xDrawPage);
            }
        }
        /// <summary>
        /// Removes the page from the drawing.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="xDrawPage">The x draw page.</param>
        public static void RemovePageFromDrawing(Object drawDoc, XDrawPage xDrawPage)
        {
            XDrawPages xDrawPages = DrawDocGetXDrawPages(drawDoc);
            if (xDrawPages != null && xDrawPage != null)
            {
                xDrawPages.remove(xDrawPage);
            }
        }
        /// <summary>
        /// Removes the page from drawing.
        /// </summary>
        /// <param name="drawDocXDrawPages">The draw doc X draw pages.</param>
        /// <param name="xDrawPage">The x draw page.</param>
        public static void RemovePageFromDrawing(XDrawPages drawDocXDrawPages, XDrawPage xDrawPage)
        {
            if (drawDocXDrawPages != null && xDrawPage != null)
            {
                drawDocXDrawPages.remove(xDrawPage);
            }
        }

        #region  Properties of Pages

        /// <summary>
        /// Gets the name of the page.
        /// </summary>
        /// <param name="drawPage">The draw page.</param>
        /// <returns></returns>
        public static String GetPageName(Object drawPage)
        {
            // Get a different interface to the drawDoc.
            //XNamed drawPage_XNamed = QI.XNamed( drawPage );       
            //return drawPage_XNamed.getName();
            return OoUtils.XNamedGetName(drawPage);
        }

        /// <summary>
        /// Sets the name of the page.
        /// </summary>
        /// <param name="drawPage">The draw page.</param>
        /// <param name="pageName">Name of the page.</param>
        public static void SetPageName(Object drawPage, String pageName)
        {
            // Get a different interface to the drawDoc.
            //XNamed drawPage_XNamed = QI.XNamed( drawPage );
            //drawPage_XNamed.setName( pageName );
            OoUtils.XNamedSetName(drawPage, pageName);
        }

        #endregion

        #endregion

        //----------------------------------------------------------------------
        //  Sugar coated access to layers of a drawing document.
        //   The first layer of a drawing is page zero.
        //----------------------------------------------------------------------

        #region Operations on Draw Document

        /// <summary>
        /// Gets the count of available draw layers.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <returns>The count of available layers in this draw document</returns>
        public static int GetNumDrawLayers(Object drawDoc)
        {
            XLayerManager xLayerManager = DrawDocGetXLayerManager(drawDoc);
            return xLayerManager != null ? xLayerManager.getCount() : 0;
        }

        /// <summary>
        /// Get one of the useful interfaces from a drawing document.
        /// XDrawPages gives you...
        ///      XDrawPage insertNewByIndex( int pageIndex )
        ///      void remove( XDrawPage drawPage )
        /// Since XDrawPages includes XIndexAccess, you also get...
        ///      int getCount()
        ///      Object getByIndex( long index )
        /// Since XIndexAccess includes XElementAccess, you also get...
        ///      type getElementType()
        ///      boolean hasElements()
        /// </summary>
        /// <param name="drawDoc">The draw doc (XDrawPagesSupplier).</param>
        /// <returns></returns>
        public static XDrawPages DrawDocGetXDrawPages(Object drawDoc)
        {
            // Get a different interface to the drawDoc.
            // The parameter passed in to us is the wrong interface to the drawDoc.
            XDrawPagesSupplier drawDocXDrawPagesSupplier = drawDoc as XDrawPagesSupplier;

            if (drawDocXDrawPagesSupplier != null)
            {
                // Ask the drawing document to give us it's draw pages object.
                Object drawPages = drawDocXDrawPagesSupplier.getDrawPages();

                // Get the XDrawPages interface to the object.
                XDrawPages drawPagesXDrawPages = drawPages as XDrawPages;
                return drawPagesXDrawPages;
            }

            return null;
        }

        /// <summary>
        ///  Gets a List of all available draw pages.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <returns></returns>
        public static List<XDrawPage> DrawDocGetXDrawPageList(Object drawDoc)
        {
            List<XDrawPage> dpL = new List<XDrawPage>();
            XDrawPages dps = DrawDocGetXDrawPages(drawDoc);

            if (dps != null)
            {
                for (int i = 0; i < dps.getCount(); i++)
                {
                    try
                    {
                        var p = dps.getByIndex(i);
                        if (p.hasValue() && p.Value is XDrawPage)
                        {
                            dpL.Add(p.Value as XDrawPage);
                        }
                    }
                    catch { }
                }
            }
            return dpL;
        }

        /// <summary>
        /// Get one of the useful interfaces from a drawing document.
        /// XLayerManager gives you...
        ///      XLayer insertNewByIndex( int layerIndex )
        ///      void remove( XLayer layer )
        ///      void attachShapeToLayer( XShape shape, XLayer layer )
        ///      XLayer getLayerForShape( XShape shape )
        /// Since XLayerManager includes XIndexAccess, you also get...
        ///      int getCount()
        ///      Object getByIndex( long index )
        /// Since XIndexAccess includes XElementAccess, you also get...
        ///      type getElementType()
        ///      boolean hasElements()
        /// QueryInterface can also be used to get an XNameAccess from this object.
        /// XNameAccess gives you...
        ///      Object getByName( String name )
        ///      String[] getElementNames()
        ///      boolean hasByName( String name )
        /// </summary>
        /// <param name="drawDoc">The draw doc (XLayerSupplier).</param>
        /// <returns></returns>
        public static XLayerManager DrawDocGetXLayerManager(Object drawDoc)
        {
            // Get a different interface to the drawDoc.
            // The parameter passed in to us is the wrong interface to the drawDoc.
            XLayerSupplier drawDocXLayerSupplier = drawDoc as XLayerSupplier;

            if (drawDocXLayerSupplier != null)
            {
                // Ask the drawing document to give us it's layer manager object.
                Object layerManager = drawDocXLayerSupplier.getLayerManager();

                // Get the XLayerManager interface to the object.
                var layerManagerXLayerManager = layerManager as XLayerManager;

                return layerManagerXLayerManager;
            }

            return null;
        }

        #endregion

        //----------------------------------------------------------------------
        //  Operations on Shapes
        //----------------------------------------------------------------------

        #region Shape Creation

        /// <summary>
        /// Creates a rectangle shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateRectangleShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_RECT, x, y, width, height);
        }
        /// <summary>
        /// Creates a rectangle shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape[XShape] or null</returns>
        public static Object CreateRectangleShape_anonymous(Object drawDoc, int x, int y, int width, int height)
        { return CreateRectangleShape(drawDoc, x, y, width, height); }
        /// <summary>
        /// Creates a ellipse shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateEllipseShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_ELLIPSE, x, y, width, height);
        }
        /// <summary>
        /// Creates a ellipse shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape[XShape] or null</returns>
        public static Object CreateEllipseShape_anonymous(Object drawDoc, int x, int y, int width, int height)
        { return CreateEllipseShape(drawDoc, x, y, width, height); }
        /// <summary>
        /// Creates a line shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateLineShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_LINE, x, y, width, height);
        }
        /// <summary>
        /// Creates a line shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static Object CreateLineShape_anonymous(Object drawDoc, int x, int y, int width, int height)
        { return CreateLineShape(drawDoc, x, y, width, height); }
        /// <summary>
        /// Creates a text shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateTextShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_TEXT, x, y, width, height);
        }
        /// <summary>
        /// Creates a text shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape[XShape] or null</returns>
        public static Object CreateTextShape_anonymous(Object drawDoc, int x, int y, int width, int height)
        { return CreateTextShape(drawDoc, x, y, width, height); }

        #region Freeforms

        /// <summary>
        /// Creates a Bezier shape.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <param name="closed">if set to <c>true</c> the form will be closed automatically.</param>
        /// <returns>The Bezier shape or <c>null</c></returns>
        internal static XShape CreateBezierShape(Object drawDoc, bool closed = true)
        {
            return closed ?
                CreateShape(drawDoc, OO.Services.DRAW_SHAPE_BEZIER_CLOSED) :
                CreateShape(drawDoc, OO.Services.DRAW_SHAPE_BEZIER_OPEN);
        }
        // <summary>
        /// Creates a Bezier shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <param name="closed">if set to <c>true</c> the form will be closed automatically.</param>
        /// <returns>The Bezier shape[XShape] or <c>null</c></returns>
        public static Object CreateBezierShape_anonymous(Object drawDoc, bool closed = true)
        { return CreateBezierShape(drawDoc, closed); }

        /// <summary>
        /// Creates a polyline shape.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <returns>The Polyline shape or <c>null</c></returns>
        internal static XShape CreatePolylineShape(Object drawDoc)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_POLYLINE);
        }
        /// <summary>
        /// Creates a polyline shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <returns>The Polyline shape[XShape] or <c>null</c></returns>
        public static Object CreatePolylineShape_anonymous(Object drawDoc)
        { return CreatePolylineShape(drawDoc); }

        /// <summary>
        /// Creates a polygon shape.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <param name="closed">if set to <c>true</c> the form will be closed automatically.</param>
        /// <returns>A PolyPolygon if closed, a Polyline shape if open or <c>null</c></returns>
        internal static XShape CreatePolygonShape(Object drawDoc, bool closed = true)
        {
            return closed ?
                CreateShape(drawDoc, OO.Services.DRAW_SHAPE_POLYPOLYGON) :
                CreatePolylineShape(drawDoc);
        }
        /// <summary>
        /// Creates a polygon shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <param name="closed">if set to <c>true</c> the form will be closed automatically.</param>
        /// <returns>A PolyPolygon if closed, a Polyline shape[XShape] if open or <c>null</c></returns>
        public static Object CreatePolygonShape_anonymous(Object drawDoc, bool closed = true)
        { return CreatePolygonShape(drawDoc, closed); }

        #endregion

        /// <summary>
        /// Creates a shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateShape(Object drawDoc, String service)
        {
            // We need the XMultiServiceFactory interface.
            XMultiServiceFactory drawDocXMultiServiceFactory;
            if (drawDoc is XMultiServiceFactory)
            {
                // If the right interface was passed in, just typecaset it.
                drawDocXMultiServiceFactory = drawDoc as XMultiServiceFactory;
            }
            else
            {
                // Get a different interface to the drawDoc.
                // The parameter passed in to us is the wrong interface to the drawDoc.
                drawDocXMultiServiceFactory = OO.GetMultiServiceFactory();
            }

            if (drawDocXMultiServiceFactory != null)
            {
                // Ask MultiServiceFactory to create a shape.
                // Yuck, it gives back an Object with no specific interface.
                Object shapeNoInterface = drawDocXMultiServiceFactory.createInstance(service);

                // Get a more useful interface to the shape object.
                var shape = shapeNoInterface as XShape;

                return shape;
            }
            else
            {
                XMultiComponentFactory xMcf = OO.GetMultiComponentFactory();
                if (xMcf != null)
                {
                    // Ask MultiServiceFactory to create a shape.
                    // Yuck, it gives back an Object with no specific interface.
                    Object shapeNoInterface = xMcf.createInstanceWithContext(service, OO.GetContext());

                    // Get a more useful interface to the shape object.
                    var shape = shapeNoInterface as XShape;

                    return shape;
                }
            }
            return null;
        }
        /// <summary>
        /// Creates a shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <returns>The resulting shape[XShape] or null</returns>
        public static Object CreateShape_anonymous(Object drawDoc, String service)
        { return CreateShape(drawDoc, service); }

        /// <summary>
        /// Creates a shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        internal static XShape CreateShape(Object drawDoc, String service, int x, int y, int width, int height)
        {
            XShape shape = CreateShape(drawDoc, service);
            SetShapePositionAndSize(shape, x, y, width, height);
            return shape;
        }
        /// <summary>
        /// Creates a shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape[XShape] or null</returns>
        public static Object CreateShape_anonymous(Object drawDoc, String service, int x, int y, int width, int height)
        { return CreateShape(drawDoc, service, x, y, width, height); }

        #endregion

        #region Shape to Page Relation
        /// <summary>
        /// Adds a shape to a draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape [XShape].</param>
        /// <param name="drawDoc">The draw document [XDrawPagesSupplier].</param>
        public static void AddShapeToDrawPage(Object shape, Object drawDoc) { AddShapeToDrawPage(shape as XShape, drawDoc as XDrawPagesSupplier); }
        /// <summary>
        /// Adds a shape to a draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="drawDoc">The draw document [XDrawPagesSupplier].</param>
        internal static void AddShapeToDrawPage(XShape shape, Object drawDoc) { AddShapeToDrawPage(shape, drawDoc as XDrawPagesSupplier); }
        /// <summary>
        /// Adds a shape to a draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="page">The page.</param>
        internal static void AddShapeToDrawPage(XShape shape, XDrawPagesSupplier drawDoc)
        {
            if (shape != null && drawDoc != null)
            {
                try
                {
                    AddShapeToDrawPage(shape, GetCurrentPage(drawDoc));
                }
                catch { }
            }
        }
        /// <summary>
        /// Adds a shape to a draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="page">The page.</param>
        internal static void AddShapeToDrawPage(XShape shape, XShapes page)
        {
            if (shape != null && page != null)
            {
                try
                {
                    ((XShapes)page).add(shape);
                }
                catch { }
            }
        }

        /// <summary>
        /// Removes the shape from the draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="page">The page.</param>
        internal static void RemoveShapeFromDrawPage(XShape shape, XShapes page)
        {
            if (shape != null && page != null)
            {
                try
                {
                    ((XShapes)page).remove(shape);
                }
                catch { }
            }
        }
        /// <summary>
        /// Removes the shape from the draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape [XShape].</param>
        /// <param name="page">The page [XShapes].</param>
        public static void RemoveShapeFromDrawPage(Object shape, Object page) { RemoveShapeFromDrawPage(shape as XShape, page as XShapes); }

        /// <summary>
        /// Gets the current page.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <returns></returns>
        internal static XDrawPage GetCurrentPage(XDrawPagesSupplier drawDoc)
        {
            if (drawDoc != null)
            {
                XModel xModel = drawDoc as XModel;
                if (xModel != null)
                {
                    XController xController = xModel.getCurrentController();
                    if (xController != null)
                    {
                        XPropertySet xPropSet = xController as XPropertySet;
                        if (xPropSet != null)
                        {
                            try
                            {
                                XDrawPage xDrawPg = xPropSet.getPropertyValue("CurrentPage").Value as XDrawPage;
                                return xDrawPg;
                            }
                            catch
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, "OoDrawUtils", "can't get current page");
                            }
                        }
                    }
                }
            }
            return null;
        }
        /// <summary>
        /// Gets the current page.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="drawDoc">The draw doc [XDrawPagesSupplier].</param>
        /// <returns>the currently active draw page [XDrawPage]</returns>
        public static Object GetCurrentPage_anonymous(Object drawDoc) { return GetCurrentPage(drawDoc as XDrawPagesSupplier); }
        #endregion

        #region Shape Size and Position

        /// <summary>
        /// Sets the size and position of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        internal static void SetShapePositionAndSize(XShape shape, int x, int y, int width, int height)
        {
            SetShapePosition(shape, x, y);
            SetShapeSize(shape, width, height);
        }
        /// <summary>
        /// Sets the size and position of the shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="shape">The shape. [XShape]</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        public static void SetShapePositionAndSize(Object shape, int x, int y, int width, int height)
        { SetShapePositionAndSize(shape as XShape, x, y, width, height); }

        /// <summary>
        /// Sets the position of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        internal static void SetShapePosition(XShape shape, int x, int y)
        {
            if (shape != null)
            {
                var position = new Point(x, y);
                shape.setPosition(position);
            }
        }
        /// <summary>
        /// Sets the position of the shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="shape">The shape. [XShape]</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        public static void SetShapePosition(Object shape, int x, int y)
        { SetShapePosition(shape as XShape, x, y); }

        /// <summary>
        /// Sets the size of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        internal static void SetShapeSize(XShape shape, int width, int height)
        {
            if (shape != null)
            {
                var size = new Size(width, height);
                shape.setSize(size);
            }
        }
        /// <summary>
        /// Sets the size of the shape.
        /// ANONYMOUS: UNO type interfaces will be covered. Only Objects are taken and given.
        /// </summary>
        /// <param name="shape">The shape. [XShape]</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        public static void SetShapeSize(Object shape, int width, int height)
        { SetShapeSize(shape as XShape, width, height); }

        #endregion

        #region Shape Property Manipulation

        /// <summary>
        /// Sets the height property.
        /// </summary>
        /// <param name="obj">The obj whos hight should be set.</param>
        /// <param name="height">The height.</param>
        public static void SetHeight(Object obj, int height)
        {
            OoUtils.SetIntProperty(obj, "Height", height);
        }
        /// <summary>
        /// Gets the height property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>the height of an element</returns>
        public static int GetHeight(Object obj)
        {
            return OoUtils.GetIntProperty(obj, "Height");
        }

        /// <summary>
        /// Sets the width property.
        /// </summary>
        /// <param name="obj">The obj who's width should be set.</param>
        /// <param name="width">The width.</param>
        public static void SetWidth(Object obj, int width)
        {
            OoUtils.SetIntProperty(obj, "Width", width);
        }
        /// <summary>
        /// Gets the width property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>the width of an element</returns>
        public static int GetWidth(Object obj)
        {
            return OoUtils.GetIntProperty(obj, "Width");
        }

        /// <summary>
        /// Sets the FillColor property of an element.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="rgbFillColor">RGB Color to fill.</param>
        public static void SetFillColor(Object obj, int rgbFillColor)
        {
            OoUtils.SetIntProperty(obj, "FillColor", rgbFillColor);
            OoUtils.SetProperty(obj, "FillStyle", unoidl.com.sun.star.drawing.FillStyle.SOLID);
        }

        /// <summary>
        /// Gets theFillColor property of an element.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>integer representing the fill color</returns>
        public static int GetFillColor(Object obj)
        {
            return OoUtils.GetIntProperty(obj, "FillColor");
        }

        #endregion

        //public static void convertIntoUnit(int size, unoidl.com.sun.star.util.MeasureUnit targetUnit)
        //{
        //    //    unoidl.com.sun.star.util      
        //}

        #region General Functions

        /// <summary>
        /// Gets the XAccessible from a draw pages supplier.
        /// </summary>
        /// <param name="dps">The DPS.</param>
        /// <returns>the corresponding XAccessible if possible otherwise <c>null</c></returns>
        internal static XAccessible GetXAccessibleFromDrawPagesSupplier(XDrawPagesSupplier dps)
        {
            XAccessible acc = null;
            if (dps != null && dps is XModel2)
            {
                var controller = ((XModel2)dps).getCurrentController();
                if (controller != null)
                {
                    Object componenetWindow = ((XController2)controller).ComponentWindow;
                    // is unoidl.com.sun.star.accessibility.XAccessible :D

                    if (componenetWindow != null && componenetWindow is XAccessible)
                        acc = componenetWindow as XAccessible;
                }
            }
            return acc;
        }

        /// <summary>
        /// Determines whether a point is inside a rect.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="rect">The rect.</param>
        /// <returns>
        /// 	<c>true</c> if the point is inside the rect; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsPointInRect(Point point, Rectangle rect)
        {
            return IsPointInRect(point.X, point.Y, rect);
        }
        /// <summary>
        /// Determines whether a point is inside a rect.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <param name="rect">The rect.</param>
        /// <returns>
        ///   <c>true</c> if the point is inside the rect; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsPointInRect(int x, int y, Rectangle rect)
        {
            return IsPointInRect(x, y, rect.X, rect.Y, rect.Width, rect.Height);
        }
        /// <summary>
        /// Determines whether a point is inside a rect.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <param name="rX">The x of the rectangle.</param>
        /// <param name="rY">The y of the rectangle.</param>
        /// <param name="rWidth">Width of the rectangle.</param>
        /// <param name="rHeight">Height of the rectangle.</param>
        /// <returns>
        ///   <c>true</c> if the point is inside the rect; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsPointInRect(int x, int y, int rX, int rY, int rWidth, int rHeight)
        {
            if (x < rX)
                return false;
            if (y < rY)
                return false;
            if (x > (rX + rWidth))
                return false;
            if (y > (rY + rHeight))
                return false;

            return true;
        }

        #endregion

        #region 100thmm to pixel conversion

        /// <summary>
        /// Calculates pixel value from a length value in 100th/mm and given zoom (in %) and given dpi.
        /// Dots per mm = DPI / 25.4
        /// </summary>
        /// <param name="valueIn100thMm">A length value in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pixelPerMeter">The pixels per meter. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// The length in pixels.
        /// </returns>
        public static int ConvertToPixel(int valueIn100thMm, int zoomInPercent, double pixelPerMeter = (96.0 / 25.4) * 1000.0)
        {
            return (int)(zoomInPercent * valueIn100thMm * pixelPerMeter / 10000000.0);
        }

        /// <summary>
        /// Converts a coordinate from 100th/mm to pixel coordinate.
        /// </summary>
        /// <param name="pos">Coordinate in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Coordinate in px.
        /// </returns>
        internal static Point convertToPixel(Point pos, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new Point(ConvertToPixel(pos.X, zoomInPercent, pxPerMeterX), ConvertToPixel(pos.Y, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a coordinate from 100th/mm to pixel coordinate.
        /// </summary>
        /// <param name="pos">Coordinate in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Coordinate in px.
        /// </returns>
        public static System.Drawing.Point convertToPixel(System.Drawing.Point pos, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Point(ConvertToPixel(pos.X, zoomInPercent, pxPerMeterX), ConvertToPixel(pos.Y, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a size from 100th/mm to pixel size.
        /// </summary>
        /// <param name="size">Size in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Size in px.
        /// </returns>
        internal static Size convertToPixel(Size size, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new Size(ConvertToPixel(size.Width, zoomInPercent, pxPerMeterX), ConvertToPixel(size.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a size from 100th/mm to pixel size.
        /// </summary>
        /// <param name="size">Size in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Size in px.
        /// </returns>
        public static System.Drawing.Size convertToPixel(System.Drawing.Size size, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Size(ConvertToPixel(size.Width, zoomInPercent, pxPerMeterX), ConvertToPixel(size.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a rectangle from 100th/mm coordinates to pixel coordinates and size.
        /// </summary>
        /// <param name="rect">Rectangle in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Rectangle in px.
        /// </returns>
        internal static unoidl.com.sun.star.awt.Rectangle convertToPixel(unoidl.com.sun.star.awt.Rectangle rect, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new unoidl.com.sun.star.awt.Rectangle(ConvertToPixel(rect.X, zoomInPercent, pxPerMeterX), ConvertToPixel(rect.Y, zoomInPercent, pxPerMeterY),
                ConvertToPixel(rect.Width, zoomInPercent, pxPerMeterX), ConvertToPixel(rect.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a rectangle from 100th/mm coordinates to pixel coordinates and size.
        /// </summary>
        /// <param name="rect">Rectangle in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="pxPerMeterX">The px per meter x. Default is (96.0 / 25.4) * 1000.0</param>
        /// <param name="pxPerMeterY">The px per meter y. Default is (96.0 / 25.4) * 1000.0</param>
        /// <returns>
        /// Rectangle in px.
        /// </returns>
        public static System.Drawing.Rectangle convertToPixel(System.Drawing.Rectangle rect, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Rectangle(ConvertToPixel(rect.X, zoomInPercent, pxPerMeterX), ConvertToPixel(rect.Y, zoomInPercent, pxPerMeterY),
                ConvertToPixel(rect.Width, zoomInPercent, pxPerMeterX), ConvertToPixel(rect.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts pixel into the internal document unit of 100/mm.
        /// ATTENTION: the return internal document units have to be corrected with the current applied zoom level,
        /// if you want to transform pixel to units! You have to do this by your own afterwards!! 
        /// </summary>
        /// <param name="value">The pixel value.</param>
        /// <returns> value / (96.0 / 25.4) * 100 = value in 100/mm on a resolution of 96 dpi</returns>
        public static int ConvertPixelTo100thmm(double value)
        {
            double v = value / (96.0 / 25.4);
            v = v * 100;
            return (int)Math.Round(v, 0, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Applies an affine transformation, given by the transformation matrix to a given coordinate
        /// </summary>
        /// <param name="pos">The original coordinate.</param>
        /// <param name="homogenMatrix">The transformation matrix.</param>
        /// <returns>The transformed matrix.</returns>
        internal static Point Transform(Point pos, HomogenMatrix3 homogenMatrix)
        {
            // x' = a11 * x + a12 * y + a13
            // y' = a21 * x + a22 * y + a23
            //      a31       a32       a33   // 3rd line is only to allow matrix inversion, has no additional information, should always be [0, 0, 1]
            return new Point(
                (int)(homogenMatrix.Line1.Column1 * pos.X + homogenMatrix.Line1.Column2 * pos.Y + homogenMatrix.Line1.Column3),
                (int)(homogenMatrix.Line2.Column1 * pos.X + homogenMatrix.Line2.Column2 * pos.Y + homogenMatrix.Line2.Column3)
            );
        }

        /// <summary>
        /// Applies an affine transformation, given by the transformation matrix to a given rectangle.
        /// The resulting bounding box might be larger than the original rectangle if the object is rotated.
        /// </summary>
        /// <param name="rect">The original bounding rectangle.</param>
        /// <param name="homogenMatrix">The transformation matrix.</param>
        /// <returns>The bounding box of the transformed rectangle.</returns>
        internal static Rectangle TransformBoundingBox(Rectangle rect, HomogenMatrix3 homogenMatrix)
        {
            // transform all corners
            // p1       p2
            // ┌────────┐
            // │        │
            // └────────┘
            // p3       p4
            Point p1 = Transform(new Point(rect.X, rect.Y), homogenMatrix);
            Point p2 = Transform(new Point(rect.X + rect.Width, rect.Y), homogenMatrix);
            Point p3 = Transform(new Point(rect.X + rect.Width, rect.Y + rect.Height), homogenMatrix);
            Point p4 = Transform(new Point(rect.X, rect.Y + rect.Height), homogenMatrix);
            int minX = Math.Min(Math.Min(Math.Min(p1.X, p2.X), p3.X), p4.X);
            int maxX = Math.Max(Math.Max(Math.Max(p1.X, p2.X), p3.X), p4.X);
            int minY = Math.Min(Math.Min(Math.Min(p1.Y, p2.Y), p3.Y), p4.Y);
            int maxY = Math.Max(Math.Max(Math.Max(p1.Y, p2.Y), p3.Y), p4.Y);
            return new Rectangle(minX, minY, maxX - minX, maxY - minY);
        }
        #endregion

        #region Screen Position To Document Position
        /// <summary>
        /// Gets the document position from a screen position.
        /// </summary>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="currentPage">The current page [XDrawPage].</param>
        /// <param name="currentView">The current view [XDrawView].</param>
        /// <returns>The position on the page based on document page position and current zoom level.</returns>
        public static System.Drawing.Point GetDocumentPositionFromScreenPosition(double x, double y, Object currentPage, Object currentView)
        {
            Point p = GetDocumentPositionFromScreenPosition(x, y, currentPage as XDrawPage, currentView as XDrawView);
            return new System.Drawing.Point(p.X, p.Y);
        }

        /// <summary>
        /// Gets the document position from a screen position.
        /// </summary>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="currentPage">The current page [XDrawPage].</param>
        /// <param name="currentView">The current view [XDrawView].</param>
        /// <returns>
        /// The position on the page based on document page position and current zoom level.
        /// </returns>
        internal static Point GetDocumentPositionFromScreenPosition(double x, double y, XDrawPage currentPage, XDrawView currentView)
        {
            if (currentPage != null && currentView != null)
            {
                var zoom = OoUtils.GetProperty(currentView, "ZoomValue");
                var viewOffset = OoUtils.GetProperty(currentView, "ViewOffset");

                int zoomInPercent = Int32.Parse(zoom.ToString());

                Point viewOffsetAsPoint = viewOffset is Point ? (Point)viewOffset : new Point(0, 0);
                double zoomFactor = ((double)zoomInPercent / 100.0);

                double xOffsetZoomed = zoomFactor * viewOffsetAsPoint.X;
                double yOffsetZoomed = zoomFactor * viewOffsetAsPoint.Y;

                double xIn100thmm = ConvertPixelTo100thmm(x) + xOffsetZoomed;
                double yIn100thmm = ConvertPixelTo100thmm(y) + yOffsetZoomed;

                int xZoomed = (int)Math.Round(xIn100thmm / zoomFactor, 0, MidpointRounding.AwayFromZero);
                int yZoomed = (int)Math.Round(yIn100thmm / zoomFactor, 0, MidpointRounding.AwayFromZero);

                Point p = new Point(xZoomed, yZoomed);
                return p;
            }
            return null;
        }

        /// <summary>
        /// Gets the document position without screen position.
        /// </summary>
        /// <param name="page">The page[XDrawPage].</param>
        /// <param name="drawView">The draw view[XDrawView].</param>
        /// <param name="coords1">The coords1.</param>
        /// <param name="coords2">The coords2.</param>
        public static void GetDocumentPositionWithoutScreenPosition(Object page, Object drawView, out System.Drawing.Point coords1, out System.Drawing.Point coords2)
        {
            System.Drawing.Point p1, p2;
            GetDocumentPositionWithoutScreenPosition(page as XDrawPage, drawView as XDrawView, out p1, out p2);
            coords1 = p1;
            coords2 = p2;
        }

        /// <summary>
        /// Gets the document position without screen position.
        /// </summary>
        /// <param name="page">The page.</param>
        /// <param name="drawView">The draw view.</param>
        /// <param name="coords1">The coords1.</param>
        /// <param name="coords2">The coords2.</param>
        internal static void GetDocumentPositionWithoutScreenPosition(XDrawPage page, XDrawView drawView, out Point coords1, out Point coords2)
        {
            var zoom = OoUtils.GetProperty(drawView, "ZoomValue");
            var viewOffset = OoUtils.GetProperty(drawView, "ViewOffset");
            var width = OoUtils.GetProperty(page, "Width");
            var height = OoUtils.GetProperty(page, "Height");

            double zoomInPercent = Int32.Parse(zoom.ToString());
            double pageHeight = Int32.Parse(height.ToString());
            double pageWidth = Int32.Parse(width.ToString());

            Point viewOffsetAsPoint = viewOffset is Point ? (Point)viewOffset : new Point(0, 0);
            double zoomFactor = (zoomInPercent / 100.0);

            double xZoomedOffset = zoomFactor * viewOffsetAsPoint.X;
            double yZoomedOffset = zoomFactor * viewOffsetAsPoint.Y;

            double x1 = Math.Abs(xZoomedOffset) + pageWidth / 3;
            double y1 = Math.Abs(yZoomedOffset) + pageHeight / 3;
            double x2 = Math.Abs(xZoomedOffset) + pageWidth * 2 / 3;
            double y2 = Math.Abs(yZoomedOffset) + pageHeight * 2 / 3;

            int xOffsetZoomed = (int)Math.Round(x1 + xZoomedOffset, 0, MidpointRounding.AwayFromZero);
            int yOffsetZoomed = (int)Math.Round(y1 + yZoomedOffset, 0, MidpointRounding.AwayFromZero); ;
            int x2OffsetZoomed = (int)Math.Round(x2 + xZoomedOffset, 0, MidpointRounding.AwayFromZero); ;
            int y2OffsetZoomed = (int)Math.Round(y2 + yZoomedOffset, 0, MidpointRounding.AwayFromZero); ;

            coords1 = new Point(xOffsetZoomed, yOffsetZoomed);
            coords2 = new Point(x2OffsetZoomed, y2OffsetZoomed);
        }

        #endregion

        #region Homogen Matrix Conversion

        internal static double[,] ConvertHomogenMatrix3ToMatrix(HomogenMatrix3 transformProp)
        {
            // initialization of a two dimensional array
            // matrix [m,n] = array[rows,cols] or array[y,x]
            // concatenation of row initializations
            // build the standard basis for a transformation matrix
            //  1   0   0
            //  0   1   0
            //  0   0   1
            double[,] matrix = new double[,] { { 1.0, 0.0, 0.0 }, { 0.0, 1.0, 0.0 }, { 0.0, 0.0, 1.0 } };

            if (transformProp != null)
            {
                if (transformProp.Line1 != null)
                {
                    matrix[0, 0] = transformProp.Line1.Column1;
                    matrix[0, 1] = transformProp.Line1.Column2;
                    matrix[0, 2] = transformProp.Line1.Column3;
                }
                if (transformProp.Line2 != null)
                {
                    matrix[1, 0] = transformProp.Line2.Column1;
                    matrix[1, 1] = transformProp.Line2.Column2;
                    matrix[1, 2] = transformProp.Line2.Column3;
                }
                if (transformProp.Line3 != null)
                {
                    matrix[2, 0] = transformProp.Line3.Column1;
                    matrix[2, 1] = transformProp.Line3.Column2;
                    matrix[2, 2] = transformProp.Line3.Column3;
                }
            }

            return matrix;
        }

        internal static HomogenMatrix3 ConvertMatrixToHomogenMatrix3(double[,] matrix)
        {
            HomogenMatrix3 m = new HomogenMatrix3();

            if (matrix != null)
            {
                try
                {
                    var HomogenMatrix3Type = typeof(HomogenMatrix3);
                    var HomogenMatrixLine3Type = typeof(HomogenMatrixLine3);

                    for (int i = 0; i < 3 && i < matrix.GetLength(0); i++)
                    {
                        var lineField = HomogenMatrix3Type.GetField("Line" + (i + 1));
                        if (lineField != null)
                        {
                            HomogenMatrixLine3 line = lineField.GetValue(m) as HomogenMatrixLine3;

                            for (int j = 0; j < 3 && j < matrix.GetLength(1); j++)
                            {
                                var colField = HomogenMatrixLine3Type.GetField("Column" + (j + 1));
                                if (colField != null)
                                {
                                    colField.SetValue(line, matrix[i, j]);
                                }
                            }
                        }
                    }
                }
                catch (System.Exception) { }
            }
            return m;
        }

        #endregion

    }

    /// <summary>
    /// Class with helper functions for handling polygons, polylines etc.
    /// </summary>
    public static class PolygonHelper
    {
        /// <summary>
        /// Determines whether this shape is a closed free form or an open one.
        /// </summary>
        /// <param name="shape">The shape [XShape].</param>
        /// <returns>
        ///   <c>true</c> if this is a closed free form; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsClosedFreeForm(Object shape)
        {
            return IsClosedFreeForm(shape as XShape);
        }
        /// <summary>
        /// Determines whether this shape is a closed free form or an open one.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>
        ///   <c>true</c> if this is a closed free form; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsClosedFreeForm(XShape shape)
        {
            return shape != null && (
                   OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_BEZIER_CLOSED)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_POLYPOLYGON));
        }

        /// <summary>
        /// Determines whether the specified shape is freeform.
        /// Freeforms are polygons, polylines and Bezier curves.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>
        ///   <c>true</c> if the specified shape is a freeform; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsFreeform(XShape shape)
        {
            return OoUtils.ElementSupportsService(shape, OO.Services.DRAW_POLY_POLYGON_DESCRIPTOR)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_POLY_POLYGON_BEZIER_DESCRIPTOR)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_BEZIER_CLOSED)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_BEZIER_OPEN)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_POLYLINE)
                || OoUtils.ElementSupportsService(shape, OO.Services.DRAW_SHAPE_POLYPOLYGON)
                ;
        }
        /// <summary>
        /// Determines whether the specified shape is freeform.
        /// Freeforms are polygons, polylines and Bezier curves.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>
        ///   <c>true</c> if the specified shape is a freeform; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsFreeform(Object shape) { return IsFreeform(shape as XShape); }

        /// <summary>
        /// Builds a free form such as a polygon or a Bezier curve etc.
        /// </summary>
        /// <param name="drawDoc">The draw document.</param>
        /// <param name="kind">The kind of freeform you want to create.</param>
        /// <param name="coordinates">The coordinates of the form(s). By adding more coordinate lists in the List, it will create a poly kind of the requested kind.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <returns>XShape</returns>
        public static Object BuildFreeForm(Object drawDoc, PolygonKind kind, List<List<PolyPointDescriptor>> coordinates = null, bool geometry = false)
        {
            XShape shape = null;

            switch (kind)
            {
                case PolygonKind.LINE:
                case PolygonKind.PLIN: // polygon open
                    shape = OoDrawUtils.CreatePolylineShape(drawDoc);
                    break;

                case PolygonKind.PATHPLIN:
                case PolygonKind.FREELINE:
                case PolygonKind.PATHLINE: // Freehand open, Curve
                    shape = OoDrawUtils.CreateBezierShape(drawDoc, false);
                    break;

                case PolygonKind.FREEFILL:
                case PolygonKind.PATHFILL: // Freehand filled, Curve filled
                    shape = OoDrawUtils.CreateBezierShape(drawDoc);
                    break;

                case PolygonKind.PATHPOLY:
                case PolygonKind.POLY: // polygon filled
                    shape = OoDrawUtils.CreatePolygonShape(drawDoc);
                    break;

                default:
                    break;
            }

            SetPolyPoints(shape, coordinates, geometry);
            return shape;
        }

        /// <summary>
        /// Sets the poly points.
        /// </summary>
        /// <param name="shape">The shape to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        internal static bool SetPolyPoints(XShape shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        {
            if (shape != null && coordinates != null && coordinates.Count > 0)
            {
                #region Polygon
                if (OoUtils.ElementSupportsService(shape, OO.Services.DRAW_POLY_POLYGON_DESCRIPTOR))
                {
                    return AddPointsToPolygon(shape, coordinates, geometry, doc);
                }
                #endregion

                #region Bezier
                // Bezier
                else if (OoUtils.ElementSupportsService(shape, OO.Services.DRAW_POLY_POLYGON_BEZIER_DESCRIPTOR))
                {
                    return AddPointsToPolyPolygonBezierDescriptor(shape, coordinates, geometry, doc);
                }
                #endregion
                else
                {
                } // not a polygon or bezier 
            }

            return false;
        }

        /// <summary>
        /// Sets the poly points.
        /// </summary>
        /// <param name="shape">The shape to change the points. [XShape]</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        public static bool SetPolyPoints(Object shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        { return SetPolyPoints(shape as XShape, coordinates, geometry, doc); }

        /// <summary>
        /// Adds the points to a poly polygon bezier descriptor.
        /// </summary>
        /// <param name="shape">The shape to change the points.</param>
        /// <param name="coordinates">The coordinates.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon).
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns>
        ///   <c>true</c> if the property could be set successfully
        /// </returns>
        internal static bool AddPointsToPolyPolygonBezierDescriptor(XShape shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        {
            // Geometry           .drawing.PolyPolygonBezierCoords  -STRUCT-
            // PolyPolygonBezier  .drawing.PolyPolygonBezierCoords  -STRUCT-
            //      Coordinates   [].drawing.PointSequence          -Sequence- 
            //                          [].awt.Point                -Sequence- 
            //      Flags         [].drawing.FlagSequence           -Sequence-
            //                           [].drawing.Flag            -Sequence-

            PolyPolygonBezierCoords ppc = BuildPolyPolygonBezierCoords(coordinates);
            return OoUtils.SetPropertyUndoable(shape, geometry ? "Geometry" : "PolyPolygonBezier", ppc, doc as unoidl.com.sun.star.document.XUndoManagerSupplier);
        }
        /// <summary>
        /// Adds the points to a poly polygon bezier descriptor.
        /// </summary>
        /// <param name="shape">The shape[XShape] to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        public static bool AddPointsToPolyPolygonBezierDescriptor(Object shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        { return AddPointsToPolyPolygonBezierDescriptor(shape as XShape, coordinates, geometry, doc); }

        /// <summary>
        /// Adds the points to a poly polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape[XShape] to change the points.</param>
        /// <param name="coordinates">The coordinates.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon).
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns>
        ///   <c>true</c> if the property could be set successfully
        /// </returns>
        internal static bool AddPolyPointsToPolyPolygonDescriptor(XShape shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        {
            int i = 0;
            Point[][] ps = new Point[coordinates.Count][];
            unodrawing.PolygonFlags[] f;
            foreach (var polygon in coordinates)
            {
                Point[] p;
                bool success = TransformToCoordinateAndFlagSequence(coordinates[0], out p, out f, true);
                ps[i] = p;
                i++;
            }
            if (geometry)
            {
                return OoUtils.SetPropertyUndoable(shape, "Geometry", ps, doc as unoidl.com.sun.star.document.XUndoManagerSupplier);
            }
            else
            {
                return OoUtils.SetPropertyUndoable(shape, "PolyPolygon", ps, doc as unoidl.com.sun.star.document.XUndoManagerSupplier);
            }
        }
        /// <summary>
        /// Adds the points to a poly polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape[XShape] to change the points.</param>
        /// <param name="coordinates">The coordinates.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon).
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns>
        ///   <c>true</c> if the property could be set successfully
        /// </returns>
        public static bool AddPolyPointsToPolyPolygonDescriptor(Object shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        { return AddPolyPointsToPolyPolygonDescriptor(shape as XShape, coordinates, geometry, doc); }

        /// <summary>
        /// Adds the points to a polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        internal static bool AddPointsToPolyPolygonDescriptor(XShape shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        {
            if (coordinates.Count > 0)
            {
                List<Point[]> polygonList = new List<Point[]>();
                unodrawing.PolygonFlags[] f;
                foreach (var pointList in coordinates)
                {
                    if (pointList != null && pointList.Count > 0)
                    {
                        Point[] p;
                        bool success = TransformToCoordinateAndFlagSequence(pointList, out p, out f, true);
                        if (success)
                        {
                            polygonList.Add(p);
                        }
                    }
                }

                Point[][] val = polygonList.ToArray<Point[]>();

                return OoUtils.SetPropertyUndoable(shape,
                    geometry ? "Geometry" : "PolyPolygon", 
                    val, doc as unoidl.com.sun.star.document.XUndoManagerSupplier); 
            }

            return false;
        }
        /// <summary>
        /// Adds the points to a polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape[XShape] to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        public static bool AddPointsToPolyPolygonDescriptor(Object shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        { return AddPointsToPolyPolygonDescriptor(shape as XShape, coordinates, geometry, doc); }

        /// <summary>
        /// Adds the points to a poly polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        internal static bool AddPointsToPolygon(XShape shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        {
            // Geometry         [][].awt.Point     -Sequence-   NOTE: is empty if not a poly!
            // PolyPolygon      [][].awt.Point     -Sequence-   NOTE: is empty if not a poly!  
            // Polygon          [].awt.Point       -Sequence-     

            if (coordinates.Count > 1) // poly polygon
            {
                return AddPolyPointsToPolyPolygonDescriptor(shape, coordinates, geometry, doc);
            }
            else // polygon
            {
                if (coordinates[0] != null)
                {
                    return AddPointsToPolyPolygonDescriptor(shape, coordinates, geometry, doc);
                }
                return false;
            }
        }
        /// <summary>
        /// Adds the points to a poly polygon descriptor.
        /// </summary>
        /// <param name="shape">The shape[XShape] to change the points.</param>
        /// <param name="coodinates">The complete coordinates to set.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// Normally you should use the absolute coordinates on the page! So use the default value <c>false</c>.</param>
        /// <param name="doc">The document [unoidl.com.sun.star.document.XUndoManagerSupplier]. If this is set, the manipulation will be added to the undo/redo history</param>
        /// <returns><c>true</c> if the property could be set successfully</returns>
        public static bool AddPointsToPolygon(Object shape, List<List<PolyPointDescriptor>> coordinates, bool geometry = false, object doc = null)
        { return AddPointsToPolygon(shape as XShape, coordinates, geometry, doc); }

        #region PolyPointHandling

        /// <summary>
        /// Gets the poly points of a free form shape.
        /// </summary>
        /// <param name="shape">The shape [XShape].</param>
        /// <param name="getGeometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// <returns>List of lists containing the point descriptors for polygon points</returns>
        public static List<List<PolyPointDescriptor>> GetPolyPoints(Object shape, bool getGeometry = false)
        { return GetPolyPoints(shape as XShape, getGeometry); }


        internal static System.Drawing.Point _toPoint(unoidl.com.sun.star.awt.Point point)
        {
            return new System.Drawing.Point(point.X, point.Y);
        }


        /// <summary>
        /// Gets the poly points of a free form shape.
        /// </summary>
        /// <param name="shape">The shape [XShape].</param>
        /// <param name="getGeometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed bezier coordinates of the polygon). 
        /// <returns>List of lists containing the point descriptors for polygon points</returns>
        internal static List<List<PolyPointDescriptor>> GetPolyPoints(XShape shape, bool getGeometry = false)
        {
            //System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            //watch.Restart();

            //try
            //{
            if (shape != null && IsFreeform(shape))
            {
                try
                {
                    // check for the right type
                    if (getGeometry)
                    {
                        // Geometry         [][].awt.Point      -Sequence-
                        Point[][] geometry = OoUtils.GetProperty(shape, "Geometry") as Point[][];
                    }
                    else
                    {
                        // PolyPolygon      [][].awt.Point      -Sequence-
                        Point[][] points = new Point[0][];
                        points = OoUtils.GetProperty(shape, "PolyPolygon") as Point[][];

                        // PolyPolygonBezier    .drawing.PolyPolygonBezierCoords  -STRUCT-
                        if (points == null)
                        {
                            PolyPolygonBezierCoords coords = OoUtils.GetProperty(shape, "PolyPolygonBezier") as PolyPolygonBezierCoords;
                            return GetPolyPoints(coords);
                        }
                        else
                        {
                            return GetPolyPoints(points);
                        }
                    }
                }
                catch { }
            }
            return null;
            //}
            //finally
            //{
            //    watch.Stop();
            //    System.Diagnostics.Debug.WriteLine("Ticks for handling Polygon getting event:" + watch.ElapsedTicks + " / Milliseconds:" + watch.ElapsedMilliseconds);
            //}
        }

        /// <summary>
        /// Gets a List of Lists of PolyPoint descriptors from a property value.
        /// </summary>
        /// <param name="coordinates">The coordinates.</param>
        /// <returns>
        /// List of lists containing the point descriptors for polygon points
        /// </returns>
        internal static List<List<PolyPointDescriptor>> GetPolyPoints(Point[] coordinates)
        {
            List<List<PolyPointDescriptor>> pointLists = new List<List<PolyPointDescriptor>>();
            if (coordinates != null && coordinates.Length > 0)
            {
                List<PolyPointDescriptor> polyPoints = new List<PolyPointDescriptor>();
                if (coordinates != null)
                {
                    foreach (Point p in coordinates)
                    {
                        polyPoints.Add(new PolyPointDescriptor(p));
                    }
                }
                pointLists.Add(polyPoints);
            }
            return pointLists;
        }

        /// <summary>
        /// Gets a List of Lists of PolyPoint descriptors from a property value.
        /// </summary>
        /// <param name="coordinates">The coordinates.</param>
        /// <returns>
        /// List of lists containing the point descriptors for polygon points
        /// </returns>
        internal static List<List<PolyPointDescriptor>> GetPolyPoints(Point[][] coordinates)
        {
            List<List<PolyPointDescriptor>> pointLists = new List<List<PolyPointDescriptor>>();
            if (coordinates != null && coordinates.Length > 0)
            {
                foreach (Point[] polygon in coordinates)
                {
                    List<PolyPointDescriptor> polyPoints = new List<PolyPointDescriptor>();
                    if (polygon != null)
                    {
                        foreach (Point p in polygon)
                        {
                            polyPoints.Add(new PolyPointDescriptor(p));
                        }
                    }
                    pointLists.Add(polyPoints);
                }
            }
            return pointLists;
        }

        /// <summary>
        /// Gets a List of Lists of PolyPoint descriptors from a property value.
        /// </summary>
        /// <param name="coordinates">The coordinates.</param>
        /// <returns>List of lists containing the point descriptors for bezier points</returns>
        internal static List<List<PolyPointDescriptor>> GetPolyPoints(unoidl.com.sun.star.drawing.PolyPolygonBezierCoords coordinates)
        {
            List<List<PolyPointDescriptor>> pointLists = null;
            if (coordinates != null)
            {
                Point[][] coords = coordinates.Coordinates;
                unodrawing.PolygonFlags[][] flags = coordinates.Flags;

                if (coords != null && coords.Length > 0 && coords[0].Length > 0)
                {
                    pointLists = new List<List<PolyPointDescriptor>>(coords.Length);

                    for (int i = 0; i < coords.Length; i++)
                    {
                        Point[] coordList = coords[i];
                        unodrawing.PolygonFlags[] flaglist = flags.Length > i ? flags[i] : null;

                        if (coordList != null && coordList.Length > 0)
                        {
                            List<PolyPointDescriptor> points = new List<PolyPointDescriptor>(coordList.Length);

                            for (int j = 0; j < coordList.Length; j++)
                            {
                                points.Add(new PolyPointDescriptor(coordList[j],
                                    flaglist != null && flaglist.Length > j ? (PolygonFlags)flaglist[j] : (PolygonFlags)unodrawing.PolygonFlags.NORMAL
                                    ));
                            }
                            pointLists.Add(new List<PolyPointDescriptor>(points));
                        }
                    }
                    pointLists.TrimExcess();
                }
            }

            return pointLists != null ? pointLists : new List<List<PolyPointDescriptor>>();
        }

        /// <summary>
        /// Builds the poly polygon bezier coords property value.
        /// </summary>
        /// <param name="points">The points.</param>
        /// <returns>a PolyPolygonBezierCoords, which multidimensional arrays are filled with the given coordinates and flags</returns>
        internal static PolyPolygonBezierCoords BuildPolyPolygonBezierCoords(List<List<PolyPointDescriptor>> points)
        {
            PolyPolygonBezierCoords property = new PolyPolygonBezierCoords();

            if (points != null && points.Count > 0)
            {
                Point[][] AllCoordinates = new Point[points.Count][];
                unodrawing.PolygonFlags[][] AllFlags = new unodrawing.PolygonFlags[points.Count][];

                int i = 0;
                foreach (var item in points)
                {
                    Point[] Coordinates;
                    unodrawing.PolygonFlags[] Flags;

                    bool success = TransformToCoordinateAndFlagSequence(item, out Coordinates, out Flags);
                    if (success)
                    {
                        AllCoordinates[i] = Coordinates;
                        AllFlags[i] = Flags;
                        i++;
                    }
                }

                property.Coordinates = AllCoordinates;
                property.Flags = AllFlags;
            }

            return property;
        }

        /// <summary>
        /// Gets the coordinate and flag sequence arrays to be set as property values.
        /// </summary>
        /// <param name="points">The points to be transformed.</param>
        /// <param name="Coordinates">The coordinate sequence array.</param>
        /// <param name="Flags">The flag sequence array.</param>
        /// <param name="filterControlPoints">if set to <c>true</c> control points will be ignored during translation.</param>
        /// <returns>
        ///   <c>true</c> if successfully transformed; otherwise, <c>false</c>.
        /// </returns>
        internal static bool TransformToCoordinateAndFlagSequence(List<PolyPointDescriptor> points, out Point[] Coordinates, out unodrawing.PolygonFlags[] Flags, bool filterControlPoints = false)
        {
            bool success = false;
            Coordinates = null;
            Flags = null;

            if (points != null)
            {
                Coordinates = new Point[points.Count];
                Flags = new unodrawing.PolygonFlags[points.Count];

                if (points.Count > 0)
                {
                    try
                    {
                        int i = 0;
                        foreach (var item in points)
                        {
                            if (!filterControlPoints || item.Flag == PolygonFlags.NORMAL)
                            {
                                // System.Linq.Enumerable.ElementAt(points, i);
                                Coordinates[i] = new Point(item.X, item.Y);
                                Flags[i] = (unodrawing.PolygonFlags)item.Flag;
                                i++;
                            }
                        }

                        if (i < points.Count)
                        {
                            Coordinates = Coordinates.Take(i).ToArray();
                            Flags = Flags.Take(i).ToArray();
                            //TODO: check if this works
                        }
                    }
                    catch { return false; }
                }
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Builds a list of poly point descriptors from a comma ',' separated list of short describing string.
        /// Each of the point description-string must be of the structure 'FLAG X Y [V[.vv]]' the different parts have to be separated by a free space.
        /// The FLAG can be the first letter of the type (C will turn into CORTOL and S into SMOOTH) or the whole word. The value parameter 'V' is optional
        /// </summary>
        /// <param name="ppDescr">The poly point description.</param>
        /// <returns>A List of PolyPointDescriptor objects filled with the specified values.</returns>
        public static List<PolyPointDescriptor> ParsePolyPointDescriptorsFromString(string ppDescrs)
        {
            List<PolyPointDescriptor> points = new List<PolyPointDescriptor>();
            if (!String.IsNullOrWhiteSpace(ppDescrs))
            {
                string[] components = ppDescrs.Split(new char[]{','}, StringSplitOptions.RemoveEmptyEntries);
                if (components.Length > 0)
                {
                    foreach (var item in components)
                    {
                        try
                        {
                            points.Add(ParsePolyPointDescriptorFromString(item));
                        }
                        catch (System.ArgumentNullException) { }
                        catch (System.Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }
            return points;
        }


        /// <summary>
        /// Builds the poly point descriptor from a short describing string.
        /// The string must be of the structure 'FLAG X Y [V[.vv]]' the different parts have to be separated by a free space.
        /// The FLAG can be the first letter of the type (C will turn into CORTOL and S into SMOOTH) or the whole word. The value parameter 'V' is optional
        /// </summary>
        /// <param name="ppDescr">The poly point description.</param>
        /// <returns>A PolyPointDescriptor object filled with the specified values.</returns>
        /// <exception cref="System.ArgumentException">string to parse does not contain enough number of parameter. - ppDescr</exception>
        /// <exception cref="System.ArgumentNullException">ppDescr - String to parse can't be empty</exception>
        public static PolyPointDescriptor ParsePolyPointDescriptorFromString(string ppDescr)
        {
            if (!String.IsNullOrWhiteSpace(ppDescr))
            {
                PolyPointDescriptor ppD = new PolyPointDescriptor();

                // split the string
                string[] components = ppDescr.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);

                if (components.Length < 3)
                {
                    throw new ArgumentException("string to parse does not contain enough number of parameter.", "ppDescr");
                }

                int x = 0;
                int y = 0;

                if (!int.TryParse(components[1], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out x)) x = 0;
                if (!int.TryParse(components[2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out y)) y = 0;


                ppD.X = x;
                ppD.Y = y;
                ppD.Flag = ParseFlagFromString(components[0]);

                // parse value
                if (components.Length > 3)
                {
                    object value = null;

                    int intVal = 0;
                    Double doubleVal = 0.0;
                    if (!int.TryParse(components[3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out intVal))
                    {
                        if (!Double.TryParse(components[3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out doubleVal))
                        {
                            value = components[3];
                        }
                        else { value = doubleVal; }
                    }
                    else { value = intVal; }

                    ppD.Value = value;
                }

                return ppD;
            }
            else
            {
                throw new ArgumentNullException("ppDescr", "String to parse can't be empty");
            }
        }


        /// <summary>
        /// Parses a string into a PolygonFlags if possible.
        /// </summary>
        /// <param name="_flag">The flag string, e.g. the whole Flag as name or even the first letter (case doesn't matter).</param>
        /// <returns>The corresponding PolygonFlags or PolygonFlags.CUSTOM as default or in case of error.</returns>
        internal static PolygonFlags ParseFlagFromString(string _flag)
        {

            if (!String.IsNullOrWhiteSpace(_flag))
            {
                _flag = _flag.ToUpper();

                try
                {
                    if (Enum.IsDefined(typeof(PolygonFlags), _flag))
                    {
                        return (PolygonFlags)Enum.Parse(typeof(PolygonFlags), _flag, true);
                    }
                }
                catch (System.Exception) { }


                switch (_flag)
                {
                    case "C":
                    case "CONTROL":
                        return PolygonFlags.CONTROL;
                    case "CUSTOM":
                        return PolygonFlags.CUSTOM;
                    case "N":
                    case "NORMAL":
                        return PolygonFlags.NORMAL;
                    case "S":
                    case "SMOOTH":
                        return PolygonFlags.SMOOTH;
                    case "SYMMETRIC":
                        return PolygonFlags.SYMMETRIC;
                    default:
                        break;
                }
            }

            return PolygonFlags.CUSTOM;
        }


        #endregion

    }

    #region Enums

    /// <summary>
    /// This is a set of properties to describe the style for rendering an area. 
    /// </summary>
    public enum FillStyle : short
    {
        /// <summary>
        /// the area is not filled.  
        /// </summary>
        NONE = unoidl.com.sun.star.drawing.FillStyle.NONE,
        /// <summary>
        /// use a solid color to fill the area. 
        /// </summary>
        SOLID = unoidl.com.sun.star.drawing.FillStyle.SOLID,
        /// <summary>
        /// use a gradient color to fill the area.
        /// </summary>
        GRADIENT = unoidl.com.sun.star.drawing.FillStyle.GRADIENT,
        /// <summary>
        /// use a hatch to fill the area.
        /// </summary>
        HATCH = unoidl.com.sun.star.drawing.FillStyle.HATCH,
        /// <summary>
        /// use a bitmap to fill the area.
        /// </summary>
        BITMAP = unoidl.com.sun.star.drawing.FillStyle.BITMAP
    }

    /// <summary>
    /// The BitmapMode selects an algorithm for filling an area with a bitmap.
    /// </summary>
    public enum BitmapMode
    {
        /// <summary>
        /// the bitmap is repeated over the fill area.
        /// </summary>
        REPEAT = unoidl.com.sun.star.drawing.BitmapMode.REPEAT,
        /// <summary>
        /// the bitmap is stretched to fill the area. 
        /// </summary>
        STRETCH = unoidl.com.sun.star.drawing.BitmapMode.STRETCH,
        /// <summary>
        /// the bitmap is painted in its original or selected size.  
        /// </summary>
        NO_REPEAT = unoidl.com.sun.star.drawing.BitmapMode.NO_REPEAT
    }

    /// <summary>
    /// specifies the appearance of the lines of a shape.
    /// </summary>
    public enum LineStyle
    {
        /// <summary>
        ///     the linestyle is unknown. Maybe this element has not this property
        /// </summary>
        UNKNOWN = 0,
        /// <summary>
        /// 	the line is hidden. 
        /// </summary>
        NONE = unoidl.com.sun.star.drawing.LineStyle.NONE,
        /// <summary>
        /// the line is solid. 
        /// </summary>
        SOLID = unoidl.com.sun.star.drawing.LineStyle.SOLID,
        /// <summary>
        /// the line use dashes. 
        /// </summary>
        DASH = unoidl.com.sun.star.drawing.LineStyle.DASH
    }

    /// <summary>
    /// This enumeration defines the style of a dash on a line
    /// </summary>
    public enum DashStyle
    {
        /// <summary>	
        /// the dash is a rectangle  
        /// </summary>
        RECT = unoidl.com.sun.star.drawing.DashStyle.RECT,
        /// <summary>
        /// the dash is a point
        /// </summary>
        ROUND = unoidl.com.sun.star.drawing.DashStyle.ROUND,
        /// <summary>
        /// the dash is a rectangle, with the size of the dash given in relation to the length of the line  
        /// </summary>
        RECTRELATIVE = unoidl.com.sun.star.drawing.DashStyle.RECTRELATIVE,
        /// <summary>
        /// the dash is a point, with the size of the dash given in relation to the length of the line  
        /// </summary>
        ROUNDRELATIVE = unoidl.com.sun.star.drawing.DashStyle.ROUNDRELATIVE
    }

    /// <summary>
    /// Defines how a bezier curve goes through a point.
    /// </summary>
    public enum PolygonFlags
    {
        /// <summary>
        /// the point is normal, from the curve discussion view.  
        /// </summary>
        NORMAL = unoidl.com.sun.star.drawing.PolygonFlags.NORMAL,
        /// <summary>
        /// the point is smooth, the first derivation from the curve discussion view.  
        /// </summary>
        SMOOTH = unoidl.com.sun.star.drawing.PolygonFlags.SMOOTH,
        /// <summary>
        ///  the point is a control point, to control the curve from the user interface.  
        /// </summary>
        CONTROL = unoidl.com.sun.star.drawing.PolygonFlags.CONTROL,
        /// <summary>
        /// the point is symmetric, the second derivation from the curve discussion view.
        /// </summary>
        SYMMETRIC = unoidl.com.sun.star.drawing.PolygonFlags.SYMMETRIC,

        //unoidl.com.sun.star.drawing.
        /// <summary>
        /// for customs
        /// </summary>
        CUSTOM = int.MaxValue
    }

    /// <summary>
    /// This enumeration defines the type of polygon.
    /// Property 'PolygonKind'.
    /// </summary>
    public enum PolygonKind
    {
        /// <summary>
        /// This is the PolygonKind for a LineShape. 
        /// </summary>
        LINE = unoidl.com.sun.star.drawing.PolygonKind.LINE,
        /// <summary>
        ///  This is the PolygonKind for a PolyPolygonShape.  
        /// </summary>
        POLY = unoidl.com.sun.star.drawing.PolygonKind.POLY,
        /// <summary>
        /// This is the PolygonKind for a PolyLineShape.  
        /// </summary>
        PLIN = unoidl.com.sun.star.drawing.PolygonKind.PLIN,
        /// <summary>
        /// This is the PolygonKind for an OpenBezierShape.
        /// </summary>
        PATHLINE = unoidl.com.sun.star.drawing.PolygonKind.PATHLINE,
        /// <summary>
        /// This is the PolygonKind for a ClosedBezierShape. 
        /// </summary>
        PATHFILL = unoidl.com.sun.star.drawing.PolygonKind.PATHFILL,
        /// <summary>
        /// This is the PolygonKind for an OpenFreeHandShape.
        /// </summary>
        FREELINE = unoidl.com.sun.star.drawing.PolygonKind.FREELINE,
        /// <summary>
        /// This is the PolygonKind for a ClosedFreeHandShape.
        /// </summary>
        FREEFILL = unoidl.com.sun.star.drawing.PolygonKind.FREEFILL,
        /// <summary>
        /// This is the PolygonKind for a PolyPolygonPathShape.
        /// </summary>
        PATHPOLY = unoidl.com.sun.star.drawing.PolygonKind.PATHPOLY,
        /// <summary>
        /// This is the PolygonKind for a PolyLinePathShape.
        /// </summary>
        PATHPLIN = unoidl.com.sun.star.drawing.PolygonKind.PATHPLIN,
    }

    #endregion

    #region Structs

    /// <summary>
    /// A Struct to define Points of polypoint structures.
    /// </summary>
    public struct PolyPointDescriptor
    {
        /// <summary>
        /// specifies the x-coordinate. 
        /// </summary>
        public int X;
        /// <summary>
        /// specifies the y-coordinate. 
        /// </summary>
        public int Y;
        /// <summary>
        /// defines how a bezier curve goes through a point.
        /// </summary>
        public PolygonFlags Flag;
        /// <summary>
        /// Optional Value for special treatments or special kinds of points.
        /// </summary>
        public object Value;

        public PolyPointDescriptor(int x, int y, PolygonFlags flag = PolygonFlags.NORMAL)
        {
            X = x;
            Y = y;
            Flag = flag;
            Value = null;
        }

        internal PolyPointDescriptor(Point p, PolygonFlags flag = PolygonFlags.NORMAL)
            : this(p != null ? p.X : 0, p != null ? p.Y : 0, flag) { }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return this.GetType().ToString() + " - " + Flag.ToString() + " " + X + " " + Y + (Value != null ? " " + Value.ToString() : "");
        }
    }

    #endregion

}