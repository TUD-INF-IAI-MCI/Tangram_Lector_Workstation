using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.script.provider;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.beans;
using tud.mci.tangram.models;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.uno;
using tud.mci.tangram.util;

namespace tud.mci.tangram.Examples.dialogs
{
    public static class NonModalDialog
    {

        public static XControlContainer createDialog(XComponentContext xContext,
            int posX, int posY, int width, int height, string title
            )
        { return createDialog(OO.GetMultiComponentFactory(xContext), xContext, posX, posX, width, height, title); }

        public static XControlContainer createDialog(XMultiComponentFactory xMultiComponentFactory, XComponentContext xContext,
            int posX, int posY, int width, int height, string title
            )
        {

            //dialog model
            Object oDialogModel = xMultiComponentFactory.createInstanceWithContext(
                  "com.sun.star.awt.UnoControlDialogModel", xContext);


            XPropertySet xPSetDialog = (XPropertySet)oDialogModel;
            xPSetDialog.setPropertyValue("PositionX", Any.Get(posX));
            xPSetDialog.setPropertyValue("PositionY", Any.Get(posY));
            xPSetDialog.setPropertyValue("Width", Any.Get(width));
            xPSetDialog.setPropertyValue("Height", Any.Get(height));
            xPSetDialog.setPropertyValue("Title", Any.Get(title));

            // get service manager from  dialog model
            XMultiServiceFactory MXMsfDialogModel = (XMultiServiceFactory)oDialogModel;

            // dialog control model
            Object oUnoDialog = xMultiComponentFactory.createInstanceWithContext(
                  "com.sun.star.awt.UnoControlDialog", xContext);


            XControl MXDialogControl = (XControl)oUnoDialog;
            XControlModel xControlModel = (XControlModel)oDialogModel;
            MXDialogControl.setModel(xControlModel);



            XToolkit xToolkit = (XToolkit)xMultiComponentFactory
                        .createInstanceWithContext("com.sun.star.awt.Toolkit",
                             xContext);

            WindowDescriptor aDescriptor = new WindowDescriptor();
            aDescriptor.Type = WindowClass.TOP;
            aDescriptor.WindowServiceName = "";
            aDescriptor.ParentIndex = -1;
            aDescriptor.Parent = xToolkit.getDesktopWindow();
            aDescriptor.Bounds = new Rectangle(100, 200, 300, 400);

            aDescriptor.WindowAttributes = WindowAttribute.BORDER
                  | WindowAttribute.MOVEABLE | WindowAttribute.SIZEABLE
                  | WindowAttribute.CLOSEABLE;

            XWindowPeer xPeer = xToolkit.createWindow(aDescriptor);

            XWindow xWindow = (XWindow)xPeer;
            xWindow.setVisible(false);
            MXDialogControl.createPeer(xToolkit, xPeer);

            // execute the dialog
            XDialog xDialog = (XDialog)oUnoDialog;
            xDialog.execute();

            // dispose the dialog
            XComponent xComponent = (XComponent)oUnoDialog;
            xComponent.dispose();

            return oUnoDialog as XControlContainer;
        }


    }
}
