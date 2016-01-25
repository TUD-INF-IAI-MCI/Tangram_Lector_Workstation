using System;
using System.Diagnostics;
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

namespace tud.mci.tangram.models.documents
{
    internal class DrawDocLayerManager : XModifyListener
    {
        public const string MARKER_LOCKED = "[Ø]";
        public const string MARKER_NOT_PRINTABLE = "[℗]";
        public const string MARKER_INVISIBLE = "[■]";

        private readonly XNameAccess _layerManger;

        public DrawDocLayerManager(XLayerSupplier drwaDoc)
        {
            if (drwaDoc == null)
                throw new ArgumentNullException("drwaDoc", "The XLayerSupplier is null");

            _layerManger = drwaDoc.getLayerManager();
            XModifyBroadcaster debc = drwaDoc as XModifyBroadcaster;
            if (debc != null)
                ((XModifyBroadcaster) drwaDoc).addModifyListener(this);
        }      

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
                        Debug.WriteLine(e);
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


    class TestContextMenue : AbstractContextMenuInterceptorBase
    {
        protected override bool AddMenuEnties(ContextMenuExecuteEvent aEvent, XMultiServiceFactory xMenuElementFactory, XIndexContainer xContextMenu)
        {
            debugEvent(aEvent);

            xContextMenu.insertByIndex(0, Any.Get(CreateLineSeperator(xMenuElementFactory)));
            xContextMenu.insertByIndex(0, Any.Get(CreateMenueEntry(xMenuElementFactory, "testZeug", ".uno:open", "")));
            return true;
        }


        private void debugEvent(ContextMenuExecuteEvent aEvent)
        {

            var selection = aEvent.Selection.getSelection().Value;

            System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT " + selection);
            util.Debug.GetAllInterfacesOfObject(selection);
            util.Debug.GetAllProperties(selection);

            var test = aEvent.SourceWindow;
            System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT Source window " + test);
            util.Debug.GetAllInterfacesOfObject(test);
            util.Debug.GetAllProperties(test);
            System.Diagnostics.Debug.WriteLine("\t\t _____________ window hash " + test.GetHashCode());



            var bla = aEvent.ActionTriggerContainer;
            System.Diagnostics.Debug.WriteLine("\t\t _____________ CONTEXT MENUE EVENT Container " + bla);
            util.Debug.GetAllInterfacesOfObject(bla);
            util.Debug.GetAllProperties(bla);




        }
    }

}