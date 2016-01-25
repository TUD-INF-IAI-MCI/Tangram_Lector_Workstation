using System;
using System.Threading;
using tud.mci.tangram.models;
using tud.mci.tangram.models.menus;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.drawing;

using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.ui;
using unoidl.com.sun.star.util;
using Exception = System.Exception;
using unoidl.com.sun.star.xforms;
using unoidl.com.sun.star.frame;
using tud.mci.tangram.util;

namespace tud.mci.tangram.models.documents
{
    public class DrawLayerManager : XModifyListener
    {
        public const string MARKER_LOCKED = "[Ø]";
        public const string MARKER_NOT_PRINTABLE = "[℗]";
        public const string MARKER_INVISIBLE = "[■]";

        private readonly XNameAccess _layerManger;
        public XLayerSupplier XLayerSupplier { get; set; }

        public DrawLayerManager(XLayerSupplier drwaDoc)
        {
            if (drwaDoc == null)
                throw new ArgumentNullException("drwaDoc", "The XLayerSupplier is null");

            XLayerSupplier = drwaDoc;

            _layerManger = drwaDoc.getLayerManager();
            XModifyBroadcaster debc = drwaDoc as XModifyBroadcaster;
            if (debc != null)
                ((XModifyBroadcaster) drwaDoc).addModifyListener(this);
        }

        #region get active layer

        /// <summary>
        /// Gets the active layer.
        /// </summary>
        /// <returns></returns>
        public XLayer getActiveLayer()
        {
            return getActiveLayer(XLayerSupplier);            
        }

        public static XLayer getActiveLayer(XLayerSupplier xLaySup)
        {
            if (xLaySup != null)
            {
                try
                {
                    return getActiveLayer(xLaySup as unoidl.com.sun.star.frame.XModel );
                }
                catch (unoidl.com.sun.star.uno.Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Error while getting active layer:\n" + e);
                }
            }
            return null;
        }

        public static XLayer getActiveLayer(unoidl.com.sun.star.frame.XModel xModel)
        {
            if (xModel != null)
            {
                try
                {
                    return getActiveLayer(xModel.getCurrentController());

                }
                catch (unoidl.com.sun.star.uno.Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Error while getting active layer:\n" + e);
                }
            }
            //Debug.GetAllProperties(bActiveLayer);
            return null;
        }

        public static XLayer getActiveLayer(XController xController)
        {
            XLayer bActiveLayer = null;
            if (xController != null)
            {
                try
                {
                    XPropertySet xPropSet = (XPropertySet)xController;
                    bActiveLayer = xPropSet.getPropertyValue("ActiveLayer").Value as XLayer;
                }
                catch (unoidl.com.sun.star.uno.Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Error while getting active layer:\n" + e);
                }
            }
            return bActiveLayer;
        }
        #endregion

        #region XModifyListener Members

        public void modified(EventObject aEvent)
        {
            ParameterizedThreadStart pts = HandleLayerTitles;
            var thread = new Thread(pts);
            thread.Start(_layerManger);
        }

        public void disposing(EventObject source)
        {
        }

        #endregion

        /// <summary>
        /// Check the layers for modification in properties and diplay them in the name of the layers.
        /// </summary>
        /// <param name="obj">The LayerManager as XIndexAccess used to access the XLayer.</param>
        public static void HandleLayerTitles(Object obj)
        {
            var xIndexAccess = obj as XIndexAccess;
            if (xIndexAccess == null)
                return;

            try
            {
                int c = xIndexAccess.getCount();

                //Add PropertychangeListener to all Layer
                for (int i = 0; i < c; i++)
                {
                    try
                    {
                        var xLayer = (XLayer) xIndexAccess.getByIndex(i).Value;
                        var pSet = xLayer as XPropertySet;
                        if (pSet != null)
                        {
                            string name = pSet.getPropertyValue("Name").Value.ToString();
                            string name2 = name;
                            name = DoTitleString(name, MARKER_INVISIBLE,
                                                 !(bool.Parse(pSet.getPropertyValue("IsVisible").Value.ToString())));
                            name = DoTitleString(name, MARKER_LOCKED,
                                                 (bool.Parse(pSet.getPropertyValue("IsLocked").Value.ToString())));
                            name = DoTitleString(name, MARKER_NOT_PRINTABLE,
                                                 !(bool.Parse(pSet.getPropertyValue("IsPrintable").Value.ToString())));

                            if (!name.Equals(name2))
                                pSet.setPropertyValue("Name", Any.Get(name));
                        }
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine("ERROR while handling layer titles:\n" + e);
                    }
                }
            }
            catch (Exception)
            {}
        }

        /// <summary>
        /// Prepend or remove the state flag to the string.
        /// </summary>
        /// <param name="name">The string to prepare.</param>
        /// <param name="flag">The flag-string to add or remove.</param>
        /// <param name="set">The identifiert if to set or to remove the flag if is set.</param>
        /// <returns>The new string with prepend or removed flag</returns>
        private static string DoTitleString(string name, string flag, bool set)
        {
            if (set)
            {
                if (name.IndexOf(flag, StringComparison.Ordinal) == -1)
                {
                    return flag + name;
                }
            }
            else
            {
                return name.Replace(flag, "");
            }

            return name;
        }
    }

    //public static class DrawLayerManager
    //{
    //    public static XLayer getActiveLayer(XLayerSupplier xLaySup)
    //    {
    //        XLayer bActiveLayer = null;
    //        if (xLaySup != null)
    //        {
    //            try
    //            {
    //                unoidl.com.sun.star.frame.XModel xModel = (unoidl.com.sun.star.frame.XModel)xLaySup;
    //                XController xController = xModel.getCurrentController();

    //                uno.Any oViewData = xController.getViewData();

    //                XPropertySet xPropSet = (XPropertySet)xController;
    //                bActiveLayer = xPropSet.getPropertyValue("ActiveLayer").Value as XLayer;
    //            }
    //            catch (unoidl.com.sun.star.uno.Exception e)
    //            {
    //                System.Diagnostics.Debug.WriteLine("Error while getting active layer:\n" + e);
    //            }
    //        }

    //        //Debug.GetAllProperties(bActiveLayer);
    //        return bActiveLayer;
    //    }
    //}
    
}