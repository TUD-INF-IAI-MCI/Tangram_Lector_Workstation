using System;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.lang;
using System.Collections.Generic;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.container;
using uno;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.view;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.accessibility;
using System.Threading;

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


    /**
     *
     * @author danny brewer, jens bornschein
     */

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


        //----------------------------------------------------------------------
        //  Sugar coated access to pages on a drawing document.
        //   The first page of a drawing is page zero.
        //----------------------------------------------------------------------

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


        //----------------------------------------------------------------------
        //  Sugar coated access to layers of a drawing document.
        //   The first layer of a drawing is page zero.
        //----------------------------------------------------------------------

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


        //----------------------------------------------------------------------
        //  Get useful interfaces to a drawing document.
        //----------------------------------------------------------------------


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


        //----------------------------------------------------------------------
        //  Operations on Pages
        //----------------------------------------------------------------------

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

        //  Sugar Coated property manipulation.

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
        /// <param name="obj">The obj whos width should be set.</param>
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


        //----------------------------------------------------------------------
        //  Operations on Shapes
        //----------------------------------------------------------------------


        /// <summary>
        /// Creates a rectangle shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateRectangleShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_RECT, x, y, width, height);
        }
        /// <summary>
        /// Creates a ellipse shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateEllipseShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_ELLIPSE, x, y, width, height);
        }
        /// <summary>
        /// Creates a line shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateLineShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_LINE, x, y, width, height);
        }
        /// <summary>
        /// Creates a text shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateTextShape(Object drawDoc, int x, int y, int width, int height)
        {
            return CreateShape(drawDoc, OO.Services.DRAW_SHAPE_TEXT, x, y, width, height);
        }

        /// <summary>
        /// Creates a shape.
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateShape(Object drawDoc, String service)
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
        /// </summary>
        /// <param name="drawDoc">The draw doc.</param>
        /// <param name="service">The service identifier of the shape to create.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        /// <returns>The resulting shape or null</returns>
        public static XShape CreateShape(Object drawDoc, String service, int x, int y, int width, int height)
        {
            XShape shape = CreateShape(drawDoc, service);
            SetShapePositionAndSize(shape, x, y, width, height);
            return shape;
        }

        /// <summary>
        /// Adds a shape to a draw page or an other shape group.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="page">The page.</param>
        public static void AddShapeToDrawPage(XShape shape, XShapes page)
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
        public static void RemoveShapeFromDrawPage(XShape shape, XShapes page)
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
        /// Sets the size and position of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        public static void SetShapePositionAndSize(XShape shape, int x, int y, int width, int height)
        {
            SetShapePosition(shape, x, y);
            SetShapeSize(shape, width, height);
        }

        /// <summary>
        /// Sets the position of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="x">The x position.</param>
        /// <param name="y">The y position.</param>
        public static void SetShapePosition(XShape shape, int x, int y)
        {
            if (shape != null)
            {
                var position = new Point(x, y);
                shape.setPosition(position);
            }
        }

        /// <summary>
        /// Sets the size of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="width">The width of the shape.</param>
        /// <param name="height">The height of the shape.</param>
        public static void SetShapeSize(XShape shape, int width, int height)
        {
            if (shape != null)
            {
                var size = new Size(width, height);
                shape.setSize(size);
            }
        }

        /// <summary>
        /// Determines whether a point is inside a rect.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="rect">The rect.</param>
        /// <returns>
        /// 	<c>true</c> if the point is inside the rect; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsPointInRect(Point point, Rectangle rect)
        {
            if (point.X < rect.X)
                return false;
            if (point.Y < rect.Y)
                return false;
            if (point.X > (rect.X + rect.Width))
                return false;
            if (point.Y > (rect.Y + rect.Height))
                return false;

            return true;
        }

        //  Sugar Coated property manipulation.

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

        public static void convertIntoUnit(int size, unoidl.com.sun.star.util.MeasureUnit targetUnit)
        {
            //    unoidl.com.sun.star.util      
        }

        /// <summary>
        /// Gets the XAccessible from a draw pages supplier.
        /// </summary>
        /// <param name="dps">The DPS.</param>
        /// <returns>the coresponding XAccessible if possible otherwise <c>null</c></returns>
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

        #region 100thmm to pixel conversion
        /// <summary>
        /// Calculates pixel value from a length value in 100th/mm and given zoom (in %) and given dpi.
        /// Dots per mm = DPI / 25.4
        /// </summary>
        /// <param name="valueIn100thMm">A length value in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <param name="dpi">A dpi value, 96 by default.</param>
        /// <returns>The length in pixels.</returns>
        public static int convertToPixel(int valueIn100thMm, int zoomInPercent, double pixelPerMeter = (96.0 / 25.4) * 1000.0)
        {
            return (int)(zoomInPercent * valueIn100thMm * pixelPerMeter / 10000000.0);
            //return dpi * zoomInPercent * valueIn100thMm / 254000;
        }

        /// <summary>
        /// Converts a coordinate from 100th/mm to pixel coordinate.
        /// </summary>
        /// <param name="pos">Coordinate in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Coordinate in px.</returns>
        internal static Point convertToPixel(Point pos, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new Point(convertToPixel(pos.X, zoomInPercent, pxPerMeterX), convertToPixel(pos.Y, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a coordinate from 100th/mm to pixel coordinate.
        /// </summary>
        /// <param name="pos">Coordinate in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Coordinate in px.</returns>
        public static System.Drawing.Point convertToPixel(System.Drawing.Point pos, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Point(convertToPixel(pos.X, zoomInPercent, pxPerMeterX), convertToPixel(pos.Y, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a size from 100th/mm to pixel size.
        /// </summary>
        /// <param name="size">Size in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Size in px.</returns>
        internal static Size convertToPixel(Size size, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new Size(convertToPixel(size.Width, zoomInPercent, pxPerMeterX), convertToPixel(size.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a size from 100th/mm to pixel size.
        /// </summary>
        /// <param name="size">Size in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Size in px.</returns>
        public static System.Drawing.Size convertToPixel(System.Drawing.Size size, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Size(convertToPixel(size.Width, zoomInPercent, pxPerMeterX), convertToPixel(size.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Converts a rectangle from 100th/mm coordinates to pixel coordinates and size.
        /// </summary>
        /// <param name="rect">Rectangle in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Rectangle in px.</returns>
        internal static unoidl.com.sun.star.awt.Rectangle convertToPixel(unoidl.com.sun.star.awt.Rectangle rect, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new unoidl.com.sun.star.awt.Rectangle(convertToPixel(rect.X, zoomInPercent, pxPerMeterX), convertToPixel(rect.Y, zoomInPercent, pxPerMeterY),
                convertToPixel(rect.Width, zoomInPercent, pxPerMeterX), convertToPixel(rect.Height, zoomInPercent, pxPerMeterY));
        }


        /// <summary>
        /// Converts a rectangle from 100th/mm coordinates to pixel coordinates and size.
        /// </summary>
        /// <param name="rect">Rectangle in 100th/mm.</param>
        /// <param name="zoomInPercent">A zoom value in percent (e.g. 102 for 102%).</param>
        /// <returns>Rectangle in px.</returns>
        public static System.Drawing.Rectangle convertToPixel(System.Drawing.Rectangle rect, int zoomInPercent, double pxPerMeterX = (96.0 / 25.4) * 1000.0, double pxPerMeterY = (96.0 / 25.4) * 1000.0)
        {
            return new System.Drawing.Rectangle(convertToPixel(rect.X, zoomInPercent, pxPerMeterX), convertToPixel(rect.Y, zoomInPercent, pxPerMeterY),
                convertToPixel(rect.Width, zoomInPercent, pxPerMeterX), convertToPixel(rect.Height, zoomInPercent, pxPerMeterY));
        }

        /// <summary>
        /// Applies an affine transformation, given by the transformation matrix to a given coordinate
        /// </summary>
        /// <param name="pos">The original coordinate.</param>
        /// <param name="homogenMatrix">The transformation matrix.</param>
        /// <returns>The transformed matrix.</returns>
        internal static Point transform(Point pos, HomogenMatrix3 homogenMatrix)
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
        internal static Rectangle transformBoundingBox(Rectangle rect, HomogenMatrix3 homogenMatrix)
        {
            // transform all corners
            // p1       p2
            // ┌────────┐
            // │        │
            // └────────┘
            // p3       p4
            Point p1 = transform(new Point(rect.X, rect.Y), homogenMatrix);
            Point p2 = transform(new Point(rect.X + rect.Width, rect.Y), homogenMatrix);
            Point p3 = transform(new Point(rect.X + rect.Width, rect.Y + rect.Height), homogenMatrix);
            Point p4 = transform(new Point(rect.X, rect.Y + rect.Height), homogenMatrix);
            int minX = Math.Min(Math.Min(Math.Min(p1.X, p2.X), p3.X), p4.X);
            int maxX = Math.Max(Math.Max(Math.Max(p1.X, p2.X), p3.X), p4.X);
            int minY = Math.Min(Math.Min(Math.Min(p1.Y, p2.Y), p3.Y), p4.Y);
            int maxY = Math.Max(Math.Max(Math.Max(p1.Y, p2.Y), p3.Y), p4.Y);
            return new Rectangle(minX, minY, maxX - minX, maxY - minY);
        }
        #endregion

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
                catch (System.Exception){}
            }
            return m;
        }
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

    #endregion

}