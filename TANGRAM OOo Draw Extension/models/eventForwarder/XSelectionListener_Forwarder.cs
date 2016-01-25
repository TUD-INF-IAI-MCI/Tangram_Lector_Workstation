using System;
using unoidl.com.sun.star.view;
using System.Threading;
using tud.mci.tangram.util;

namespace tud.mci.tangram.models.Interfaces
{
    class XSelectionListener_Forwarder : AbstractXEventListener_Forwarder, XSelectionChangeListener
    {
        /// <summary>
        /// is called when the selection changes. 
        /// You can get the new selection via XSelectionSupplier from 
        /// ::com::sun::star::lang::EventObject::Source. 
        /// </summary>
        public event EventHandler<EventObjectForwarder> SelectionChanged;

        void XSelectionChangeListener.selectionChanged(unoidl.com.sun.star.lang.EventObject aEvent)
        {
            TimeLimitExecutor.ExecuteWithTimeLimit( 500, () => { fireSelectionEvent(aEvent); }, "SelectionChanged");
        }

        void fireSelectionEvent(unoidl.com.sun.star.lang.EventObject aEvents)
        {
            if (SelectionChanged != null)
            {
                try
                {
                    SelectionChanged.Invoke(this, new EventObjectForwarder(aEvents));
                }
                catch { }
            }
        }

    }
}
