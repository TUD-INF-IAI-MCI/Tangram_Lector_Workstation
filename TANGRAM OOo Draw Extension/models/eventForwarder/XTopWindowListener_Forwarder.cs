using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.awt;

namespace tud.mci.tangram.models.eventForwarder
{
    class XTopWindowListener_Forwarder : AbstractXEventListener_Forwarder, XTopWindowListener
    {
        public event EventHandler<EventObjectForwarder> WindowActivated;
        public event EventHandler<EventObjectForwarder> WindowClosed;
        public event EventHandler<EventObjectForwarder> WindowClosing;
        public event EventHandler<EventObjectForwarder> WindowDeactivated;
        public event EventHandler<EventObjectForwarder> WindowMinimized;
        public event EventHandler<EventObjectForwarder> WindowNormalized;
        public event EventHandler<EventObjectForwarder> WindowOpened;

        void XTopWindowListener.windowActivated(unoidl.com.sun.star.lang.EventObject e) { fireWindowActivatedEvent(e); }
        void XTopWindowListener.windowClosed(unoidl.com.sun.star.lang.EventObject e) { fireWindowClosedEvent(e); }
        void XTopWindowListener.windowClosing(unoidl.com.sun.star.lang.EventObject e) { fireWindowClosingEvent(e); }
        void XTopWindowListener.windowDeactivated(unoidl.com.sun.star.lang.EventObject e) { fireWindowDeactivatedEvent(e); }
        void XTopWindowListener.windowMinimized(unoidl.com.sun.star.lang.EventObject e) { fireWindowMinimizedEvent(e); }
        void XTopWindowListener.windowNormalized(unoidl.com.sun.star.lang.EventObject e) { fireWindowNormalizedEvent(e); }
        void XTopWindowListener.windowOpened(unoidl.com.sun.star.lang.EventObject e) { fireWindowOpenedEvent(e); }

        void fireWindowActivatedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowActivated != null)
            {
                try
                {
                    WindowActivated.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowClosedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowClosed != null)
            {
                try
                {
                    WindowClosed.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowClosingEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowClosing != null)
            {
                try
                {
                    WindowClosing.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowDeactivatedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowDeactivated != null)
            {
                try
                {
                    WindowDeactivated.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowMinimizedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowMinimized != null)
            {
                try
                {
                    WindowMinimized.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowNormalizedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowNormalized != null)
            {
                try
                {
                    WindowNormalized.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

        void fireWindowOpenedEvent(unoidl.com.sun.star.lang.EventObject Source)
        {
            if (WindowOpened != null)
            {
                try
                {
                    WindowOpened.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }

    }
}