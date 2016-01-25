using System;
using unoidl.com.sun.star.lang;
using System.Threading;
using System.Diagnostics;

namespace tud.mci.tangram.models.Interfaces
{
    abstract class AbstractXEventListener_Forwarder : XEventListener
    {
        /// <summary>
        /// gets called when the broadcaster is about to be disposed. 
        /// All listeners and all other objects, which reference the broadcaster should 
        /// release the reference to the source. 
        /// No method should be invoked anymore on this object ( including XComponent::removeEventListener ).
        /// This method is called for every listener registration of derived listener interfaced,
        /// not only for registrations at XComponent.
        /// </summary>
        public event EventHandler<EventObjectForwarder> Disposing;

        void XEventListener.disposing(EventObject Source)
        {
            fireDisposingEvent(Source);
        }

        void fireDisposingEvent(EventObject Source)
        {
            if (Disposing != null)
            {
                try
                {
                    Disposing.Invoke(this, new EventObjectForwarder(Source));
                }
                catch { }
            }
        }
    }

    public class EventObjectForwarder : EventArgs
    {
        public readonly EventObject E;
        public EventObjectForwarder(EventObject e) { E = e; }
    }
}
