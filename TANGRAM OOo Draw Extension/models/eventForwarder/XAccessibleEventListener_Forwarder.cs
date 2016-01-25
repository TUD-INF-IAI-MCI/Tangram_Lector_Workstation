using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.accessibility;
using System.Threading;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.util;

namespace tud.mci.tangram.models.Interfaces
{
    class XAccessibleEventListener_Forwarder : AbstractXEventListener_Forwarder, XAccessibleEventListener
    {
        /// <summary>
        /// is called whenever a accessible event (see AccessibleEventObject) occurs. 
        /// </summary>
        public event EventHandler<AccessibleEventObjectForwarder> NotifyEvent;

        void XAccessibleEventListener.notifyEvent(AccessibleEventObject aEvent)
        {
            TimeLimitExecutor.ExecuteWithTimeLimit(15000, delegate() { fireSelectionEvent(aEvent); }, "AccEvntFrwd_NotifyEvent");
        }

        void fireSelectionEvent(AccessibleEventObject aEvents)
        {
            if (NotifyEvent != null)
            {
                try
                {
                    //var id = OoAccessibility.GetAccessibleEventIdFromShort(aEvents.EventId);
                    //Logger.Instance.Log(LogPriority.DEBUG, "AccFwrd", ">>----> forward Notifiy Event: " + id + " - " + this.GetHashCode());
                    NotifyEvent.Invoke(this, new AccessibleEventObjectForwarder(aEvents));
                }
                catch { }
            }
        }
    }

    public class AccessibleEventObjectForwarder : EventArgs
    {
        public readonly AccessibleEventObject E;

        public AccessibleEventObjectForwarder(AccessibleEventObject e)
        {
            E = e;
        }
    }

}
