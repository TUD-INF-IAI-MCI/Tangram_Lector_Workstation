// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extension
// Author           : Admin
// Created          : 09-13-2012
//
// Last Modified By : Admin
// Last Modified On : 09-17-2012
// ***********************************************************************
// <copyright file="AbstarctDrawDocBase.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using tud.mci.tangram.util;
using System.Collections.Generic;



namespace tud.mci.tangram.classes
{

    /** This is a helper class for OpenOffice.org Draw.
     * It connects to a running office.
     * Additionally it contains various helper functions.
     */
    public abstract class AbstarctDrawDocBase : AbstractDocBase 
    {

        //protected unoidl.com.sun.star.sheet.XSpreadsheetDocument mxDocument;


        /// <summary>
        /// Initializes a new instance of the <see cref="AbstarctDrawDocBase"/> class.
        /// </summary>
        /// <param name="args">The args.</param>
        protected AbstarctDrawDocBase(String[] args) : base(args)
        {
        }
        

        protected XDrawPagesSupplier UseDraw()
        {
            try
            {
                //create new draw document and insert rectangle shape
                XComponent xDrawComponent = NewDocComponent("sdraw");
                XDrawPagesSupplier xDrawPagesSupplier = xDrawComponent as XDrawPagesSupplier;

                Object drawPages = xDrawPagesSupplier.getDrawPages();
                XIndexAccess xIndexedDrawPages = drawPages as XIndexAccess;

                Object drawPage = xIndexedDrawPages.getByIndex(0).Value;

                System.Diagnostics.Debug.WriteLine(xIndexedDrawPages.getCount());

                if (drawPage is XDrawPage)
                {
                    XDrawPage xDrawPage = (XDrawPage)drawPage;

                    if (xDrawPage is XComponent)
                    {
                        (xDrawPage as XComponent).addEventListener(new TestOOoEventListerner());
                    }

                    // get internal service factory of the document
                    XMultiServiceFactory xDrawFactory = xDrawComponent as XMultiServiceFactory;

                    Object drawShape = xDrawFactory.createInstance(
                        "com.sun.star.drawing.RectangleShape");
                    XShape xDrawShape = drawShape as XShape;
                    xDrawShape.setSize(new Size(10000, 20000));
                    xDrawShape.setPosition(new Point(5000, 5000));
                    xDrawPage.add(xDrawShape);

                    // XText xShapeText = (XText)drawShape // COMMENTED BY CODEIT.RIGHT;
                    XPropertySet xShapeProps = (XPropertySet)drawShape;

                    // wrap text inside shape
                    xShapeProps.setPropertyValue("TextContourFrame", new uno.Any(true));
                    return xDrawPagesSupplier;
                }
                else
                {
                    //TODO: handle if no drwapage was found
                    System.Diagnostics.Debug.WriteLine("no XDrawPage found");
                    System.Diagnostics.Debug.WriteLine(drawPage);

                }

            }
            catch (unoidl.com.sun.star.lang.DisposedException e)
            { //works from Patch 1
                MXContext = null;
                throw e;
            }

            return null;
        }

        /// <summary>
        /// Copies the text.
        /// </summary>
        /// <param name="xDialog">The x dialog.</param>
        /// <param name="aEvent">A event object.</param>
        public void CopyText(XDialog xDialog, Object aEvent)
        {
            XControlContainer xControlContainer = (XControlContainer)xDialog;

            String aTextPropertyStr = "Text";
            String aText = "";
            XControl xTextField1Control = xControlContainer.getControl("TextField1");
            XControlModel xControlModel1 = xTextField1Control.getModel();
            XPropertySet xPropertySet1 = (XPropertySet)xControlModel1;
            try
            {
                aText = (String)xPropertySet1.getPropertyValue(aTextPropertyStr).Value;
            }
            catch (unoidl.com.sun.star.uno.Exception e)
            {
                Console.WriteLine("copyText caught exception! " + e);
            }

            XControl xTextField2Control = xControlContainer.getControl("TextField2");
            XControlModel xControlModel2 = xTextField2Control.getModel();
            XPropertySet xPropertySet2 = (XPropertySet)xControlModel2;
            try
            {
                xPropertySet2.setPropertyValue(aTextPropertyStr, new uno.Any(aText));
            }
            catch (unoidl.com.sun.star.uno.Exception e)
            {
                Console.WriteLine("copyText caught exception! " + e);
            }

            global::System.Windows.Forms.MessageBox.Show("DialogComponent", "copyText() called");
        }

        #region Helper Functions

        #region GetAllObjectsFromXIndexAccess
        protected static List<XDrawPage> GetAllDrawPages(XDrawPagesSupplier xDrawPagesSupplier)
        {
            return GetAllObjectsFromXIndexAccess<XDrawPage>(xDrawPagesSupplier != null ? xDrawPagesSupplier.getDrawPages() as XIndexAccess : null) as List<XDrawPage>;
        }

        protected static List<Object> GetAllDrawShapesOnPage(XDrawPage xDrawPage)
        {
            return GetAllObjectsFromXIndexAccess(xDrawPage as XIndexAccess) as List<Object>;
        }
        #endregion


        #endregion



    }

    public class TestOOoEventListerner : XContainerListener
    {
        public TestOOoEventListerner() { }

        public void elementRemoved(ContainerEvent event1)
        {
            //System.Windows.Forms.MessageBox.Show("Event Raised - REMOVEED");
        }

        public void elementReplaced(ContainerEvent event1)
        {
            //System.Windows.Forms.MessageBox.Show("Event Raised - REPLACED");
        }

        public void disposing(EventObject source)
        {
            //System.Windows.Forms.MessageBox.Show("Event Raised - DISPOSING");
        }

        public void elementInserted(ContainerEvent event1)
        {
            //System.Windows.Forms.MessageBox.Show("Event Raised - INSERT");
        }
    }
}
