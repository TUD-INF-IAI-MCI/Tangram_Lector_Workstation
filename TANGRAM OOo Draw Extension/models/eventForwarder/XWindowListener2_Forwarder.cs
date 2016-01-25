using System;
using unoidl.com.sun.star.awt;
using System.Threading;

namespace tud.mci.tangram.models.Interfaces
{
    class XWindowListener2_Forwarder : AbstractXEventListener_Forwarder, XWindowListener2
    {
        /// <summary>
        /// is called when the window has been disabled. 
        /// </summary>
        public event EventHandler<EventObjectForwarder> WindowDisabled;
        /// <summary>
        /// is called when the window has been enabled. 
        /// </summary>
        public event EventHandler<EventObjectForwarder> WindowEnabled;
        /// <summary>
        /// is invoked when the window has been hidden. 
        /// </summary>
        public event EventHandler<EventObjectForwarder> WindowHidden;
        /// <summary>
        /// is invoked when the window has been moved. 
        /// </summary>
        public event EventHandler<WindowEventForwarder> WindowMoved;
        /// <summary>
        /// is invoked when the window has been resized. 
        /// </summary>
        public event EventHandler<WindowEventForwarder> WindowResized;
        /// <summary>
        /// is invoked when the window has been shown. 
        /// </summary>
        public event EventHandler<EventObjectForwarder> WindowShown;

        public void windowDisabled(unoidl.com.sun.star.lang.EventObject e)
        {
            Thread thread = new Thread(delegate() { fireWindowDisabledEvent(e); });
            thread.Start();
        }
        public void windowEnabled(unoidl.com.sun.star.lang.EventObject e) {
            Thread thread = new Thread(delegate() { fireWindowEnabledEvent(e); });
            thread.Start(); 
        }
        public void windowHidden(unoidl.com.sun.star.lang.EventObject e) {
            Thread thread = new Thread(delegate() { fireWindowHiddenEvent(e); });
            thread.Start(); 
        }
        public void windowMoved(WindowEvent e) {
            Thread thread = new Thread(delegate() { fireWindowMovedEvent(e); });
            thread.Start();
        }
        public void windowResized(WindowEvent e) {
            Thread thread = new Thread(delegate() { fireWindowResizedEvent(e); });
            thread.Start(); 
        }
        public void windowShown(unoidl.com.sun.star.lang.EventObject e) {
            Thread thread = new Thread(delegate() { fireWindowShownEvent(e); });
            thread.Start(); 
        }

        void fireWindowDisabledEvent(unoidl.com.sun.star.lang.EventObject e)
        {
            if (WindowDisabled != null)
            {
                try
                {
                    WindowDisabled.Invoke(this, new EventObjectForwarder(e));
                }
                catch { }
            }
        }

        void fireWindowEnabledEvent(unoidl.com.sun.star.lang.EventObject e)
        {
            if (WindowEnabled != null)
            {
                try
                {
                    WindowEnabled.Invoke(this, new EventObjectForwarder(e));
                }
                catch { }
            }
        }

        void fireWindowHiddenEvent(unoidl.com.sun.star.lang.EventObject e)
        {
            if (WindowHidden != null)
            {
                try
                {
                    WindowHidden.Invoke(this, new EventObjectForwarder(e));
                }
                catch { }
            }
        }

        void fireWindowMovedEvent(WindowEvent e)
        {
            if (WindowMoved != null)
            {
                try
                {
                    WindowMoved.Invoke(this, new WindowEventForwarder(e));
                }
                catch { }
            }
        }

        void fireWindowResizedEvent(WindowEvent e)
        {
            if (WindowResized != null)
            {
                try
                {
                    WindowResized.Invoke(this, new WindowEventForwarder(e));
                }
                catch { }
            }
        }

        void fireWindowShownEvent(unoidl.com.sun.star.lang.EventObject e)
        {
            if (WindowShown != null)
            {
                try
                {
                    WindowShown.Invoke(this, new EventObjectForwarder(e));
                }
                catch { }
            }
        }

    }


    public class WindowEventForwarder : EventArgs
    {
        public readonly WindowEvent E;
        public WindowEventForwarder(WindowEvent e)
        {
            E = e;
        }
    }
}
