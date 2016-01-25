using System;
using System.Collections.Generic;
using System.Text;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.ui;

namespace tud.mci.tangram.models.menus
{
    abstract class AbstractContextMenuInterceptorBase : XContextMenuInterceptor, IDisposable
    {
        private XFrame _xFrame;
        protected XFrame Frame { get { return _xFrame; } }

        /// <summary>
        /// Registers the context menu controler.
        /// </summary>
        /// <param name="xModel">The x model.</param>
        public void RegisterContextMenuControler(XModel xModel)
        {
            _xFrame = xModel.getCurrentController().getFrame();
            RegisterContextMenuControler(_xFrame);
        }

        /// <summary>
        /// Registers the context menu controler.
        /// </summary>
        /// <param name="xFrame">The x frame.</param>
        public void RegisterContextMenuControler(XFrame xFrame)
        {
            _xFrame = xFrame;
            var xController = _xFrame.getController();

            if (xController != null)
            {
                var xContextMenuInterception = xController as XContextMenuInterception;
                if (xContextMenuInterception != null)
                {
                    //var aContextMenuInterceptor = new AbstractContextMenuInterceptorBase();
                    //var xContextMenuInterceptor =
                    //    (XContextMenuInterceptor)aContextMenuInterceptor;
                    //xContextMenuInterception.registerContextMenuInterceptor(xContextMenuInterceptor);
                    xContextMenuInterception.registerContextMenuInterceptor(this);

                    System.Diagnostics.Debug.WriteLine("ContextMenue "+this.GetType().Name+" registerd");
                }
            }
        }

        /// <summary>
        /// Unregisters the context menu controler.
        /// </summary>
        /// <param name="xFrame">The x frame.</param>
        public void UnregisterContextMenuControler(XFrame xFrame)
        {
            var xController = xFrame.getController();

            if (xController != null)
            {
                var xContextMenuInterception = xController as XContextMenuInterception;
                if (xContextMenuInterception != null)
                {
                    xContextMenuInterception.releaseContextMenuInterceptor(this);
                    System.Diagnostics.Debug.WriteLine("ContextMenue " + this.GetType().Name + " unregisterd");
                }
            }
        }

        /// <summary>
        /// Event Listener to get called.
        /// </summary>
        /// <param name="aEvent">A event.</param>
        /// <returns></returns>
        public ContextMenuInterceptorAction notifyContextMenuExecute(ContextMenuExecuteEvent aEvent)
        {
            XIndexContainer xContextMenu = aEvent.ActionTriggerContainer;
            XMultiServiceFactory xMenuElementFactory = (XMultiServiceFactory)xContextMenu;

            if (AddMenuEnties(aEvent, xMenuElementFactory, xContextMenu))
            {
                return ContextMenuInterceptorAction.CONTINUE_MODIFIED;
            }
            return ContextMenuInterceptorAction.IGNORED;
        }

        /// <summary>
        /// Creates a line seperator.
        /// </summary>
        /// <param name="xMenuElementFactory">The x menu element factory.</param>
        /// <returns></returns>
        protected XPropertySet CreateLineSeperator(XMultiServiceFactory xMenuElementFactory)
        {
            // create a line separator for our new help sub menu
            XPropertySet xSeparator =
                (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator");
            short aSeparatorType = ActionTriggerSeparatorType.LINE;
            xSeparator.setPropertyValue("SeparatorType", Any.Get(aSeparatorType));
            return xSeparator;
        }


        /// <summary>
        /// Creates the entry.
        /// </summary>
        /// <param name="entry">The entry.</param>
        /// <param name="properties">The properties.</param>
        /// <returns></returns>
        protected XPropertySet CreateEntry(XPropertySet entry, IDictionary<String, Object> properties )
        {
            if (entry != null)
            {
                foreach (var property in properties)
                {
                    entry.setPropertyValue(property.Key, Any.Get(property.Value));
                }
            }
            return entry;
        }

        /// <summary>
        /// Creates the menue entry.
        /// </summary>
        /// <param name="xMenuElementFactory">The x menu element factory.</param>
        /// <param name="text">The text.</param>
        /// <param name="commandUrl">The command URL.</param>
        /// <param name="helpUrl">The help URL.</param>
        /// <returns></returns>
        protected XPropertySet CreateMenueEntry(XMultiServiceFactory xMenuElementFactory, string text, string commandUrl, string helpUrl)
        {
            var entry =
                        (XPropertySet)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger");
            if (entry != null)
            {
                var properties = new Dictionary<String, Object>
                                     {{"Text", text}, {"CommandURL", commandUrl}, {"HelpURL", helpUrl}};
                entry = CreateEntry(entry, properties);
            }
            return entry;
        }

        /// <summary>
        /// Creates a menue entry.
        /// </summary>
        /// <param name="xMenuElementFactory">The x menu element factory.</param>
        /// <param name="text">The text.</param>
        /// <param name="commandUrl">The command URL.</param>
        /// <param name="helpUrl">The help URL.</param>
        /// <param name="subContainer">The sub container.</param>
        /// <returns></returns>
        protected XPropertySet CreateMenueEntry(XMultiServiceFactory xMenuElementFactory, string text, string commandUrl, string helpUrl, XIndexContainer subContainer)
        {
            var entry = CreateMenueEntry(xMenuElementFactory, text, commandUrl, helpUrl);
            if (entry != null)
            {
                entry.setPropertyValue("SubContainer", Any.Get(subContainer));
            }
            return entry;
        }

        /// <summary>
        /// Creates a sub menu container.
        /// </summary>
        /// <param name="xMenuElementFactory">The x menu element factory.</param>
        /// <param name="entries">The entries.</param>
        /// <returns></returns>
        protected XIndexContainer CreateSubMenuContainer(XMultiServiceFactory xMenuElementFactory, XPropertySet[] entries = null)
        {
            var entry = (XIndexContainer)xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer");
            if (entry != null && entries != null && entries.Length > 0)
            {
                for (int i = 0; i < entries.Length; i++)
                {
                    entry.insertByIndex(i,Any.Get(entries[i]));
                }
            }
            return entry;
        }

        /// <summary>
        /// Adds the menu enties.
        /// </summary>
        /// <param name="aEvent">A event.</param>
        /// <param name="xMenuElementFactory">The x menu element factory.</param>
        /// <param name="xContextMenu">The x context menu.</param>
        /// <returns></returns>
        protected abstract bool AddMenuEnties(ContextMenuExecuteEvent aEvent, XMultiServiceFactory xMenuElementFactory, XIndexContainer xContextMenu);

        /// <summary>
        /// Führt anwendungsspezifische Aufgaben durch, die mit der Freigabe, der Zurückgabe oder dem Zurücksetzen von nicht verwalteten Ressourcen zusammenhängen.
        /// </summary>
        public void Dispose()
        {
            try
            {
                UnregisterContextMenuControler(_xFrame);
            }
            catch{}
        }
    }
}
