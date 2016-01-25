using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.lang;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.frame;

namespace tud.mci.tangram.models.dialogs
{
    class ModalDialog
    {

        public static XControlContainer createDialog(XComponentContext xContext,
           int posX, int posY, int width, int height, string title
           )
       { return createDialog(OO.GetMultiComponentFactory(xContext), xContext, posX, posX, width, height, title); }

        public static XControlContainer createDialog(XMultiComponentFactory xMultiComponentFactory, XComponentContext xContext,
            int posX, int posY, int width, int height, string title
            )
        {
            try
            {
            Object oDialogModel = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", xContext);

                // The XMultiServiceFactory of the dialog model is needed to instantiate the controls...
                var MXMsfDialogModel = (XMultiServiceFactory)oDialogModel;

                // The named container is used to insert the created controls into...
                var MXDlgModelNameContainer = (XNameContainer)oDialogModel;

                // create the dialog...
                Object oUnoDialog = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", xContext);
                var MXDialogControl = (XControl)oUnoDialog;

                // The scope of the control container is public...
                var MXDlgContainer = (XControlContainer)oUnoDialog;

                var MXTopWindow = (XTopWindow)MXDlgContainer;

                // link the dialog and its model...
                XControlModel xControlModel = (XControlModel)oDialogModel;
                MXDialogControl.setModel(xControlModel);
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }


            //xMultiPropertySet.setPropertyValues(PropertyNames, Any.Get(PropertyValues));

            return null;

        }

        

    }
}
