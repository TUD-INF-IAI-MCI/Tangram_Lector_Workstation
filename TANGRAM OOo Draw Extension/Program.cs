// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extension
// Author           : Admin
// Created          : 09-12-2012
//
// Last Modified By : Admin
// Last Modified On : 10-11-2012
// ***********************************************************************
// <copyright file="Program.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using tud.mci.tangram.classes;
using tud.mci.tangram.util;
using unoidl.com.sun.star.lang;

namespace tud.mci.tangram
{
    public static class Tangram
    {

        internal static XMultiComponentFactory xMcf;

        public static bool ConnectToOO()
        {
            bool success = false;

            var cont = OO.GetContext();
            //if (cont == null)
            //{
            //    OO.CheckConnection();
            //}

            xMcf = OO.GetMultiComponentFactory(cont);

            success = xMcf != null;

            if (!success)
            {
                cont = OO.GetContext(true);
            }

            return success;
        }

        public static void GetDrawDocuments()
        {
            //List<XDrawPagesSupplier> drawPageSuppliers = new List<XDrawPagesSupplier>();

            //// get the open Draw Document
            //var dps = OOoDrawUtils.GetDrawPageSuppliers(xMcf);

            //if (dps.Count > 0)
            //{

            //    Console.WriteLine("Some DrawPageSupplieres available");

            //    foreach (XDrawPagesSupplier dp in dps)
            //    {
            //        if (dp != null) drawPageSuppliers.Add(dp);

            //        //// EXTRACT TITLE NAME AN DESCRIPTION FROM DRAW DOCUMENT

            //        DescriptionMapper dm = new DescriptionMapper(dp);
            //        var tad = dm.GetAllTitleAndDescriptions();

            //        Console.WriteLine("\tTry to getActiveLayer");
            //        //TestExamples.TryToGetTheActiveLayer(dp);

            //        //XModel model = dp as XModel;

            //        //LayerContextMenuInterceptor mi = new LayerContextMenuInterceptor();
            //        //mi.RegisterContextMenuControler(model);
            //    }
            //}
            //else
            //{
            //    XDrawPagesSupplier dp = UseDraw();
            //    drawPageSuppliers.Add(dp);
            //}


            //if (drawPageSuppliers.Count > 0)
            //{

            //    var dps0 = drawPageSuppliers[0];

            //    //get all draw pages
            //    util.Debug.GetAllInterfacesOfObject(dps0);
            //    util.Debug.GetAllServicesOfObject(dps0);

            //    System.Diagnostics.Debug.WriteLine("test");
            //    XModel model = dps0 as XModel;
            //    if (model != null)
            //    {
            //        var controller = model.getCurrentController();
            //        if (controller != null && controller is XSelectionSupplier)
            //        {
            //            try { ((XSelectionSupplier)controller).removeSelectionChangeListener(OoSelectionObserver.Instance); }
            //            catch { }
            //            try { ((XSelectionSupplier)controller).addSelectionChangeListener(OoSelectionObserver.Instance); }
            //            catch { }
            //        }
            //    }
            //    else
            //    {

            //    }

            //}



        }


        public static void Main(string[] args)
        {
            try
            {
                var xMcf = OO.GetMultiComponentFactory();

                String available = (xMcf != null ? "available" : "not available");
                //Console.WriteLine("\nremote ServiceManager is " + available);


                //Console.WriteLine("Try to start search thread");
                //var docSearcher = new DocumentListener(xMcf);
                //    docSearcher.startSearch();



                //// get the open Draw Document
                //var dps = OOoDrawUtils.GetDrawPageSuppliers(xMcf);
                //if (dps.Count > 0)
                //{

                //    Console.WriteLine("Some DrawPageSupplieres available");

                //    foreach (var dp in dps)
                //    {
                //        //// EXTRACT TITLE NAME AN DESCRIPTION FROM DRAW DOCUMENT
                //        //DescriptionMapper dm = new DescriptionMapper(dp);
                //        //var tad = dm.GetAllTitleAndDescriptions();


                //        Console.WriteLine("\tTry to getActiveLayer");
                //        TestExamples.TryToGetTheActiveLayer(dp);

                //        XModel model = dp as XModel;

                //        LayerContextMenuInterceptor mi = new LayerContextMenuInterceptor();
                //        mi.RegisterContextMenuControler(model);
                //    }
                //}




                /** opens a new draw document and print a rectangle in it **/
                //tangram.useDraw();



                /** dialog witha several kind of components **/
                //UnoDialogSample.ShowDialog();
                //var test =  tud.mci.tangram.Examples.dialogs.NonModalDialog.createDialog(OOo.GetContext(), 100, 100, 200, 400, "testDialog");

                //Console.ReadLine();
                //UnoMenu.main(null);

                //tangram.test();
                // ContextMenuInterceptor.main(null);
                // tangram.ToolBarAccesTest();

                // First: Retrieve the component context


                /** show the sample Dialog **/
                //SampleDialog dialog = new SampleDialog(xContext);
                //dialog.trigger("execute");

                //tangram.ToolBarTest();
                //Console.ReadLine();
                //TestExamples.ToolBarClassTest();
                //TestExamples.TopWindowAccesTest();

                //TestExamples.DesctopInspector();

                //TestExamples.TryToGetTheActiveLayer();


                //TestExamples.DrawLayerTest();
                //TestExamples.ContextMenuClassTest();

                Console.ReadLine();
                //UnoMenu2.main(null);

                Console.WriteLine("\nSamples done.");
            }
            catch (unoidl.com.sun.star.lang.DisposedException ex)
            {
                /*
                 * The client should resume even if the connection goes down for some reason and comes back later on. 
                 * When the connection fails, a robust, long running client should stop the current work, inform the 
                 * user that the connection is not available and release the references to the remote process. When 
                 * the user tries to repeat the last action, the client should try to rebuild the connection. Do not 
                 * force the user to restart your program just because the connection was temporarily unavailable. 
                 */
                Console.WriteLine("Connection to Open Office ServiceManager get lost with exception: " + ex);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Sample caught exception! " + ex);
            }

            Console.ReadLine();

        }

        public static bool Initialize()
        {
            var xCont = OO.GetContext();
            String c_available = (xCont != null ? "available" : "not available");

            string message = ".NET SITE\r\n\r\nProtocol Handler Initialization Call";
            message += "\nremote Context is " + c_available;

            return true;
        }

    }
}
