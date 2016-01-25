// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-18-2012
//
// Last Modified By : Admin
// Last Modified On : 09-19-2012
// ***********************************************************************
// <copyright file="ContextMenueInspector.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using unoidl.com.sun.star.ui;
using unoidl.com.sun.star.frame;
using tud.mci.tangram.util;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.view;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.container;

namespace tud.mci.tangram.models.menu
{
    /**************************************************************
 * 
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 * 
 *************************************************************/

    public class ContextMenuInterceptor__ : XContextMenuInterceptor
    {

        /**
         *Description of the Method
         *
         *@param  args  Description of Parameter
         *@since
         */
        public static void main(String[] args)
        {
            try
            {
//                var aConnect = OOo.GetContext() // COMMENTED BY CODEIT.RIGHT;
//                var factory = OOo.getMCF(aConnect) // COMMENTED BY CODEIT.RIGHT;
                XDesktop xDesktop = OO.GetDesktop();

                // create a new test document
                XComponentLoader xCompLoader = (XComponentLoader)xDesktop;
                XComponent xComponent = xCompLoader.loadComponentFromURL("private:factory/swriter", "_blank", 0, new PropertyValue[0]);

                // intialize the test document
                XFrame xFrame = null;
                {
                    XTextDocument xDoc = (XTextDocument)xComponent;

                    String infoMsg = "All context menus of the created document frame contains now a 'Help' entry with the submenus 'Content', 'Help Agent' and 'Tips'.\n\nPress 'Return' in the shell to remove the context menu interceptor and finish the example!";
                    xDoc.getText().setString(infoMsg);

                    // ensure that the document content is optimal visible
                    XModel xModel = (XModel)xDoc;
                    // get the frame for later usage
                    xFrame = xModel.getCurrentController().getFrame();

                    XViewSettingsSupplier xViewSettings =
                        (XViewSettingsSupplier)xModel.getCurrentController();
                    xViewSettings.getViewSettings().setPropertyValue(
                        "ZoomType", Any.Get((short)0));
                }
                // test document will be closed later

                // reuse the frame
                XController xController = xFrame.getController();
                if (xController != null)
                {
                    XContextMenuInterception xContextMenuInterception =
                        (XContextMenuInterception)xController;
                    if (xContextMenuInterception != null)
                    {
                        ContextMenuInterceptor__ aContextMenuInterceptor = new ContextMenuInterceptor__();
                        XContextMenuInterceptor xContextMenuInterceptor =
                            (XContextMenuInterceptor)aContextMenuInterceptor;
                        xContextMenuInterception.registerContextMenuInterceptor(xContextMenuInterceptor);

                        System.Diagnostics.Debug.WriteLine("\n ... all context menus of the created document frame contains now a 'Help' entry with the\n     submenus 'Content', 'Help Agent' and 'Tips'.\n\nPress 'Return' to remove the context menu interceptor and finish the example!");

                        //xContextMenuInterception.releaseContextMenuInterceptor(
                        //    xContextMenuInterceptor);
                        //System.Diagnostics.Debug.WriteLine(" ... context menu interceptor removed!");
                    }
                }

                //// close test document
                //XCloseable xCloseable = (XCloseable)xComponent ;

                //if (xCloseable != null ) {
                //    xCloseable.close(false);
                //} else
                //{
                //    xComponent.dispose();
                //}
            }
            catch (unoidl.com.sun.star.uno.RuntimeException)
            {
                // something strange has happend!
                //System.out.println( " Sample caught exception! " + ex );
                //System.exit(1);
            }
            catch (System.Exception)
            {
                // catch java exceptions and do something useful
                //System.out.println( " Sample caught exception! " + ex );
                //System.exit(1);
            }

            System.Diagnostics.Debug.WriteLine(" ... exit!\n");

        }

        /**
         *Description of the Method
         *
         *@param  args  Description of Parameter
         *@since
         */
        public ContextMenuInterceptorAction notifyContextMenuExecute(ContextMenuExecuteEvent aEvent)
        {

            try
            {

                XIndexContainer actionContainer = aEvent.ActionTriggerContainer;
                for (int i = 0; i < actionContainer.getCount(); i++)
                {
                    System.Diagnostics.Debug.WriteLine("SELECTED : " + i + " _____________________");
                    uno.Any aCont = actionContainer.getByIndex(i);

                    Debug.GetAllInterfacesOfObject(aCont.Value, true);
                    Debug.GetAllProperties(aCont.Value, true);
                }



                //var selected = aEvent.Selection.getSelection().Value;

                //System.Diagnostics.Debug.WriteLine("SELECTED EVENT CALL : _____________________");
                

                //XIndexAccess selIA = (XIndexAccess)selected;
                //System.Diagnostics.Debug.WriteLine("SELECTED Count : " + selIA.getCount() + " _____________________");

                //for (int i = 0; i < selIA.getCount(); i++)
                //{
                //    System.Diagnostics.Debug.WriteLine("SELECTED : " + i + " _____________________");
                //    uno.Any selItem = selIA.getByIndex(i);

                //    Debug.getAllInterfacesOfObject(selItem.Value, true);
                //    Debug.getAllProperties(selItem.Value, true);
                //}

                // Retrieve context menu container and query for service factory to
                // create sub menus, menu entries and separators
                XIndexContainer xContextMenu = aEvent.ActionTriggerContainer;
                XMultiServiceFactory xMenuElementFactory =
                    (XMultiServiceFactory)xContextMenu;
                if (xMenuElementFactory != null)
                {

                    // create root menu entry and sub menu
                    XPropertySet xRootMenuEntry =
                        (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger");

                    // create a line separator for our new help sub menu
                    XPropertySet xSeparator =
                        (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator");

                    short aSeparatorType = ActionTriggerSeparatorType.LINE;
                    xSeparator.setPropertyValue("SeparatorType", Any.Get((Object)aSeparatorType));

                    // query sub menu for index container to get access
                    XIndexContainer xSubMenuContainer = (XIndexContainer)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer");

                    // intialize root menu entry
                    xRootMenuEntry.setPropertyValue("Text", Any.Get("Help"));
                    xRootMenuEntry.setPropertyValue("CommandURL", Any.Get("slot:5410"));
                    xRootMenuEntry.setPropertyValue("HelpURL", Any.Get("5410"));
                    xRootMenuEntry.setPropertyValue("SubContainer", Any.Get((Object)xSubMenuContainer));

                    // create menu entries for the new sub menu

                    // intialize help/content menu entry
                    XPropertySet xMenuEntry = (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger");

                    xMenuEntry.setPropertyValue("Text", Any.Get("Content"));
                    xMenuEntry.setPropertyValue("CommandURL", Any.Get("slot:5401"));
                    xMenuEntry.setPropertyValue("HelpURL", Any.Get("5401"));

                    // insert menu entry to sub menu
                    xSubMenuContainer.insertByIndex(0, Any.Get(xMenuEntry));

                    // intialize help/help agent
                    xMenuEntry = (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger");
                    xMenuEntry.setPropertyValue("Text", Any.Get("Help Agent"));
                    xMenuEntry.setPropertyValue("CommandURL", Any.Get("slot:5962"));
                    xMenuEntry.setPropertyValue("HelpURL", Any.Get("5962"));

                    // insert menu entry to sub menu
                    xSubMenuContainer.insertByIndex(1, Any.Get(xMenuEntry));

                    // intialize help/tips
                    xMenuEntry = (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger");
                    xMenuEntry.setPropertyValue("Text", Any.Get("Tips"));
                    xMenuEntry.setPropertyValue("CommandURL", Any.Get("slot:5404"));
                    xMenuEntry.setPropertyValue("HelpURL", Any.Get("5404"));

                    // insert menu entry to sub menu
                    xSubMenuContainer.insertByIndex(2, Any.Get(xMenuEntry));

                    // add separator into the given context menu
                    xContextMenu.insertByIndex(0, Any.Get(xSeparator));

                    // add new sub menu into the given context menu
                    xContextMenu.insertByIndex(0, Any.Get(xRootMenuEntry));

                    // The controller should execute the modified context menu and stop notifying other
                    // interceptors.
                    return ContextMenuInterceptorAction.EXECUTE_MODIFIED;
                }
            }
            catch (unoidl.com.sun.star.beans.UnknownPropertyException)
            {
                // do something useful
                // we used a unknown property 
            }
            catch (unoidl.com.sun.star.lang.IndexOutOfBoundsException)
            {
                // do something useful
                // we used an invalid index for accessing a container
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                // something strange has happend!
            }
            catch (System.Exception ex)
            {
                // catch java exceptions and something useful
            }

            return ContextMenuInterceptorAction.IGNORED;
        }
    }
}
