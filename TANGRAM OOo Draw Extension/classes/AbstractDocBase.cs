using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.lang;
using tud.mci.tangram.util;
using unoidl.com.sun.star.frame;

namespace tud.mci.tangram.classes
{
    public abstract class AbstractDocBase : System.IDisposable
    {
        #region Members

        //TODO: make them more stable available. Maybe make them singleton

        private unoidl.com.sun.star.uno.XComponentContext m_xContext;
        public unoidl.com.sun.star.uno.XComponentContext MXContext
        {
            get
            {
                if (m_xContext == null)
                    try
                    {
                        m_xContext = OO.GetContext();
                    }
                    catch
                    {
                    }
                return m_xContext;
            }
            set { m_xContext = value; }
        }

        protected unoidl.com.sun.star.lang.XMultiServiceFactory mxMSFactory;

        #endregion


        #region Constructor / Destructor

        /// <summary>
        /// Initializes a new instance of the <see cref="AbstractDocBase"/> class.
        /// </summary>
        /// <param name="args">The args.</param>
        protected AbstractDocBase(String[] args)
        {
            // Connect to a running office and get the service manager
            //mxMSFactory = Connect(args);
        }

        #endregion


        /// <summary>
        /// Connect to a running office that is accepting connections.
        /// </summary>
        /// <param name="args">The args.</param>
        /// <returns>The ServiceManager to instantiate office components.</returns>
        public XMultiComponentFactory Connect(String[] args)
        {
            if (MXContext == null) return null;
            try
            {
                return (XMultiComponentFactory)MXContext.getServiceManager();
            }
            catch (unoidl.com.sun.star.lang.DisposedException)
            {
                return OO.GetMultiComponentFactory();
            }
        }

        /// <summary>
        /// News the doc component.
        /// </summary>
        /// <param name="docType">Type of the doc.</param>
        /// <returns></returns>
        public XComponent NewDocComponent(String docType)
        {
            var desktop = ((XMultiComponentFactory)mxMSFactory).createInstanceWithContext(
                   OO.Services.FRAME_DESKTOP, m_xContext);

            try
            {
                XComponentLoader mxCompLoader = desktop as XComponentLoader;

                String loadUrl = OO.DocTypes.DOC_TYPE_BASE + docType;
                unoidl.com.sun.star.beans.PropertyValue[] loadProps = new unoidl.com.sun.star.beans.PropertyValue[0];
                return mxCompLoader.loadComponentFromURL(loadUrl, "_blank", 0, loadProps);

            }
            catch (System.Exception ex)
            {
                //TOD: check this and handle if possible
                System.Diagnostics.Debug.WriteLine("Desktop to XComponentLoader cast Exception: " + ex);
            }
            return null;

        }



        #region Helper Functions

        /// <summary>
        /// Gets the frame from document.
        /// </summary>
        /// <param name="doc">The doc.</param>
        /// <returns></returns>
        public static XFrame GetFrameFromDocument(XModel doc)
        {
            try
            {
                return doc.getCurrentController().getFrame();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Can't get Frame form doc: " + ex);
            }

            return null;
        }

        #region GetAllObjectsFromXIndexAccess
        protected static IList<Object> GetAllObjectsFromXIndexAccess(XIndexAccess container)
        {
            return GetAllObjectsFromXIndexAccess<Object>(container);
        }

        protected static IList<T> GetAllObjectsFromXIndexAccess<T>(XIndexAccess container) where T : class
        {
            IList<T> list = new List<T>();
            if (container != null && container.hasElements())
            {
                for (int i = 0; i < container.getCount(); i++)
                {
                    var child = container.getByIndex(i).Value;
                    if (child != null && (child is T))
                        list.Add(child as T);
                }
            }
            return list;
        }
        #endregion

        #endregion



        public void Dispose()
        {
            
        }
    }
}
