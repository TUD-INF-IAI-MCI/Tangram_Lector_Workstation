using System;
using System.Collections.Generic;
using tud.mci.tangram.util;
using unoidl.com.sun.star.awt;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.models.eventForwarder;

namespace tud.mci.tangram.controller.observer
{
    class OoTopWindowObserver : IResetable
    {
        private static readonly OoTopWindowObserver instance = new OoTopWindowObserver();
        private readonly XTopWindowListener_Forwarder xTopWindowListener = new XTopWindowListener_Forwarder();
        static XExtendedToolkit extTollkit;

        private OoTopWindowObserver()
        {
            OO.BaseObjectDisposed += new EventHandler<EventArgs>(OO_BaseobjectDisposed);
            registerToTopWindowListenerForwarder();
            initialize();
        }

        private void initialize()
        {
            extTollkit = OO.GetExtTooklkit();
            if (extTollkit != null)
            {
                try {
                    extTollkit.removeTopWindowListener(xTopWindowListener); 
                }
                catch
                {
                    try {
                        extTollkit.removeTopWindowListener(xTopWindowListener); 
                    }
                    catch
                    {
                        try { System.Threading.Thread.Sleep(20); extTollkit.removeTopWindowListener(xTopWindowListener); }
                        catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't remove top window listener form extToolkit", ex); }
                    }
                }
                try {
                    extTollkit.addTopWindowListener(xTopWindowListener); 
                }
                catch
                {
                    try
                    {
                        System.Threading.Thread.Sleep(20);
                        extTollkit.addTopWindowListener(xTopWindowListener);
                    }
                    catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't add top window listener to extToolkit", ex); }
                }
            }
            else
            {
                Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] Can't get EXTENDEDTOOLKIT!!");
            }
        }

        private void registerToTopWindowListenerForwarder()
        {
            if (xTopWindowListener != null)
            {
                xTopWindowListener.Disposing += new EventHandler<EventObjectForwarder>(xTopWindowListener_Disposing);
                xTopWindowListener.WindowActivated += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowActivated);
                xTopWindowListener.WindowClosed += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowClosed);
                xTopWindowListener.WindowClosing += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowClosing);
                xTopWindowListener.WindowDeactivated += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowDeactivated);
                xTopWindowListener.WindowMinimized += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowMinimized);
                xTopWindowListener.WindowNormalized += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowNormalized);
                xTopWindowListener.WindowOpened += new EventHandler<EventObjectForwarder>(xTopWindowListener_WindowOpened);
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] XTopWindow_Forwarder is null");
            }
        }

        /// <summary>
        /// Returns the Singleton instance of this observer.
        /// </summary>
        /// <value>The Singleton instance.</value>
        public static OoTopWindowObserver Instance { get { return instance; } }

        /// <summary>
        /// Gets all OpenOffice top windows.
        /// </summary>
        /// <returns>List of top windows.</returns>
        public static List<Object> GetAllTopWindows()
        {
            return OoAccessibility.GetAllTopWindows();
        }

        /// <summary>
        /// Gets the active OpenOffice top window.
        /// </summary>
        /// <returns>The current active top window object.</returns>
        public static Object GetActiveTopWindow()
        {
            return OoAccessibility.GetActiveTopWindow();
        }

        #region XTopWindowListener

        void xTopWindowListener_WindowOpened(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowOpened " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowOpenedEvent(e.E);
        }

        void xTopWindowListener_WindowNormalized(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowNormalized " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowNormalizedEvent(e.E);
        }

        void xTopWindowListener_WindowMinimized(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowMinimized " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowMinimizedEvent(e.E);
        }

        void xTopWindowListener_WindowDeactivated(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowDeactivated " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowDeactivatedEvent(e.E);

        }

        void xTopWindowListener_WindowClosing(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowClosing " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowClosingEvent(e.E);
        }

        void xTopWindowListener_WindowClosed(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowClosed " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowClosedEvent(e.E);
        }

        void xTopWindowListener_WindowActivated(object sender, EventObjectForwarder e)
        {
            try
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "windowActivated " + e.E.Source.GetHashCode());
            }
            catch { }
            fireWindowActivatedEvent(e.E);

        }

        void xTopWindowListener_Disposing(object sender, EventObjectForwarder e)
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "disposing");
            fireDisposingEvent(e.E);
            initialize();
        }

        void OO_BaseobjectDisposed(object sender, EventArgs e)
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "Basic object disposed");
            initialize();
        }

        #endregion

        #region Events

        #region Event Definitions

        /// <summary>
        /// Occurs when a top window was activated.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowActivated;
        /// <summary>
        /// Occurs when a top window was closed.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowClosed;
        /// <summary>
        /// Occurs when a top window is currently closing.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowClosing;
        /// <summary>
        /// Occurs when  a top window was deactivated.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowDeactivated;
        /// <summary>
        /// Occurs when a top window was minimized.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowMinimized;
        /// <summary>
        /// Occurs when a top window was normalized again.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowNormalized;
        /// <summary>
        /// Occurs when a new top was window opened.
        /// </summary>
        public event EventHandler<OoEventArgs> WindowOpened;
        /// <summary>
        /// Occurs when the event provider is disposing.
        /// </summary>
        public event EventHandler<OoEventArgs> Disposing;

        #endregion

        #region Fire Events

        private void fireWindowActivatedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowActivated != null)
            {
                try
                {
                    WindowActivated.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window activated event", ex); }
            }
        }
        private void fireWindowClosedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowClosed != null)
            {
                try
                {
                    WindowClosed.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window closed event", ex); }
            }
        }
        private void fireWindowClosingEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowClosing != null)
            {
                try
                {
                    WindowClosing.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window closing event", ex); }
            }
        }
        private void fireWindowDeactivatedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowDeactivated != null)
            {
                try
                {
                    WindowDeactivated.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window deactivated event", ex); }
            }
        }
        private void fireWindowMinimizedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowMinimized != null)
            {
                try
                {
                    WindowMinimized.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window minimized event", ex); }
            }
        }
        private void fireWindowNormalizedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowNormalized != null)
            {
                try
                {
                    WindowNormalized.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window normalized event", ex); }
            }
        }
        private void fireWindowOpenedEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (WindowOpened != null)
            {
                try
                {
                    WindowOpened.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire window opend event", ex); }
            }
        }
        private void fireDisposingEvent(unoidl.com.sun.star.lang.EventObject source)
        {
            if (Disposing != null)
            {
                try
                {
                    Disposing.DynamicInvoke(this, new OoEventArgs(source));
                }
                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, this, "can't fire disposing event", ex); }
            }
        }
        #endregion

        #endregion

        public void Reset()
        {
            Logger.Instance.Log(LogPriority.DEBUG, this, "Reset the top window listener");
            initialize();
        }
    }
}