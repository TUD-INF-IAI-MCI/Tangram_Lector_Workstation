using System;
using System.Collections.Generic;
using System.Text;
using tud.mci.tangram.models.menus;
using unoidl.com.sun.star.ui;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.container;
using tud.mci.tangram.models;
using tud.mci.tangram.util;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.view;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.drawing;
using tud.mci.tangram.models.documents;

namespace tud.mci.tangram.classes.draw
{
    class LayerContextMenuInterceptor : AbstractContextMenuInterceptorBase
    {
        protected override bool AddMenuEnties(ContextMenuExecuteEvent aEvent, XMultiServiceFactory xMenuElementFactory, XIndexContainer xContextMenu)
        {

            debugEvent(aEvent);

            //if (isSelctionEmpty(aEvent.Selection))
            if(true) // FIND A WAY FOR BETTER IDENTIFING IF NEEDED
            {
                var actLayer = getActiveLayer();
                Debug.GetAllProperties(actLayer);

                var props = new List<PropertyValue>();
                props.Add(new PropertyValue("Name", 0, Any.Get("Test"), PropertyState.DIRECT_VALUE));
                props.Add(new PropertyValue("Name2", 0, Any.Get(343), PropertyState.DIRECT_VALUE));
                props.Add(new PropertyValue("Name#3", 0, Any.Get(6787678.89890), PropertyState.DIRECT_VALUE));
                var url = CommandUrlHelper.getCommandUrl(OO.UNO_COMMAND_URL_PATH, "Function3", props);
                string commandUrl = url.Complete; 
                
                xContextMenu.insertByIndex(0, Any.Get(CreateLineSeperator(xMenuElementFactory)));
                xContextMenu.insertByIndex(0, Any.Get(CreateMenueEntry(xMenuElementFactory, 
                    "testZeug", 
                    commandUrl,
                    "")));
                return true;
            }         
        }


        /// <summary>
        /// Determines whether [is selection empty] [the specified selection].
        /// </summary>
        /// <param name="selection">The selection.</param>
        /// <returns>
        /// 	<c>true</c> if [is selction empty] [the specified selection]; otherwise, <c>false</c>.
        /// </returns>
        private bool isSelectionEmpty(XSelectionSupplier selection)
        {
            if (selection.getSelection().Value != null)
            {
                var selObjType = selection.getSelection().Type;
                var selObj = selection.getSelection().Value;
                System.Diagnostics.Debug.WriteLine("SELECTED Object is of type: '" + selObjType + "'");
                return false;
            }
            return true;
        }


        private XLayer getActiveLayer()
        {
            if (Frame != null)
            {
                var actLayer = DrawLayerManager.getActiveLayer(Frame.getController());
                return actLayer;
            }
            return null;
        }


private void debugEvent(ContextMenuExecuteEvent aEvent)
        {
    //        object selection = aEvent.Selection.getSelection().Value;

    //        System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT SELECTION" + selection);
    ////scheint leer zu sein
    //        Debug.GetAllInterfacesOfObject(selection);
    //        Debug.GetAllProperties(selection);

    //        XWindow test = aEvent.SourceWindow;
    //        System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT Source window '" + test + "'");
    ////scheint immer NULL zu sein
    //        Debug.GetAllInterfacesOfObject(test);
    //        Debug.GetAllProperties(test);
    //        System.Diagnostics.Debug.WriteLine("\t\t _____________ window hash " + test.GetHashCode());


    //        XIndexContainer bla = aEvent.ActionTriggerContainer;
    //        System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT Container " + bla);
    //        System.Diagnostics.Debug.WriteLine("\t\t _____________ Containertype " + bla.getElementType());
    //        Debug.GetAllInterfacesOfObject(bla);
    //        Debug.GetAllProperties(bla);

            ////bla.getCount()
            //for (int i = 0; i < bla.getCount(); i++)
            //{
            //    object ps = bla.getByIndex(i).Value;

            //    System.Diagnostics.Debug.WriteLine("______propertie: " + ps + ": ");

            //    if (ps is XPropertySet)
            //    {
            //        var t = ps as XPropertySet;
            //        Property[] tt = t.getPropertySetInfo().getProperties();
            //        foreach (Property property in tt)
            //        {
            //            try
            //            {
            //                System.Diagnostics.Debug.WriteLine("\t property: " + property.Name + " = " +
            //                                                   t.getPropertyValue(property.Name));
            //            }
            //            catch (System.Exception e)
            //            {
            //                System.Diagnostics.Debug.WriteLine("\t property: " + property.Name + " is not reachable");
            //                Console.WriteLine(e);
            //            }
            //        }
            //    }
            //}

    	try
	{
        var pos = aEvent.ExecutePosition;

            //XWindow oInitialTarget = aEvent.SourceWindow;
            //Debug.GetAllInterfacesOfObject(oInitialTarget);

            //System.Diagnostics.Debug.WriteLine("Window ID is = " + oInitialTarget.GetHashCode());


        #region get the sending document

           


        #endregion



        #region get Child windows
            //var xVlcCont = oInitialTarget as XVclContainer;

            //if (xVlcCont != null)
            //{
            //    var windows = xVlcCont.getWindows();
            //    int i = 0;
            //    System.Diagnostics.Debug.WriteLine(windows.Length + " window in source window");
            //    foreach (var win in windows)
            //    {
            //        System.Diagnostics.Debug.WriteLine("window" + i + " in source window");
            //        Debug.GetAllInterfacesOfObject(win);
                    
            //        i++;
            //    }
            //}
        #endregion

        //    XFramesSupplier xFramesSupplier = (XFramesSupplier) oInitialTarget;
        //XFrames xFrames = xFramesSupplier.getFrames();
		
        //XIndexAccess xIndexAccess = (XIndexAccess) xFrames;

        //XFrame xFrame_5 = (XFrame) xIndexAccess.getByIndex(0).Value;

		
        //XWindow xWindow_2 = xFrame_5.getComponentWindow();
		
        //XVclContainer xVclContainer = (XVclContainer) xWindow_2;
	
        //XWindow[] xWindow_4 = xVclContainer.getWindows();

		
	}
	catch (WrappedTargetException e)
	{
		// getByIndex
		Console.WriteLine(e.Message);
	}
	catch (IndexOutOfBoundsException e)
	{
		// getByIndex
		Console.WriteLine(e.Message);
	}
	catch (RuntimeException e)
	{
		// getTypes
		Console.WriteLine(e.Message);
	}




        }










    }
}
