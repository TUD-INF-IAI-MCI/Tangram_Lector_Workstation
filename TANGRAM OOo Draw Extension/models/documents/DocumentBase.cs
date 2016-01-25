// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-17-2012
//
// Last Modified By : Admin
// Last Modified On : 09-19-2012
// ***********************************************************************
// <copyright file="DocumentBase.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using System;
using unoidl.com.sun.star.lang;
using tud.mci.tangram.util;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.uno;
using tud.mci.tangram.classes;

namespace tud.mci.tangram.models.documents
{
    class OOoDocument : AbstractDocBase
    {

        public OOoDocument() : base(null){}

        /// <summary>
        /// News the doc component.
        /// </summary>
        /// <param name="docType">Type of the doc.</param>
        /// <param name="xMsFactory"> </param>
        /// <param name="xContext"> </param>
        /// <returns></returns>
        public static XComponent OpenNewDocumentComponent(String docType, XComponentContext xContext = null, XMultiComponentFactory xMsFactory = null)
        {
            try
            {
                if (xContext == null)
                {
                    xContext = OO.GetContext();
                }

                if (xMsFactory == null)
                {
                    xMsFactory = OO.GetMultiComponentFactory(xContext);
                }

                var desktop = xMsFactory.createInstanceWithContext(OO.Services.FRAME_DESKTOP, xContext);
                var mxCompLoader = desktop as XComponentLoader;

                String loadUrl = OO.DocTypes.DOC_TYPE_BASE + docType;
                var loadProps = new unoidl.com.sun.star.beans.PropertyValue[0];
                if (mxCompLoader != null) 
                    return mxCompLoader.loadComponentFromURL(loadUrl, "_blank", 0, loadProps);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Desktop to XComponentLoader cast Exeption: " + ex);
            }
            return null;
        }
    }
}
