using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.view;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.accessibility;

namespace tud.mci.tangram.controller
{
    /// <summary>
    /// Base EventArg class for converting Oo-events
    /// </summary>
    public class OoEventArgs : EventArgs
    {
        public readonly Object Source;

        public OoEventArgs(unoidl.com.sun.star.lang.EventObject e)
        {
            if (e.Source != null)
            {
                if (e.Source is uno.Any)
                {
                    Source = ((uno.Any)e.Source).Value;
                    return;
                }
            }
            Source = e.Source;
        }

        public OoEventArgs(Object e)
        {
            Source = e;
        }
    }

    public class EventForwarder : XEventListener
    {
        #region XEventListener
        public event EventHandler<ForwardedEventArgs> Disposing;
        void XEventListener.disposing(EventObject Source)
        {
            if (Disposing != null)
            {
                try
                {
                    Disposing.Invoke(this, new ForwardedEventArgs(Source));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward disposing event: " + e);
                }
            }
        }
        #endregion
    }

    class SelectionEventForwarder : EventForwarder, XSelectionChangeListener
    {
        #region XSelectionChangeListener
        public event EventHandler<ForwardedEventArgs> SelectionChanged;
        void XSelectionChangeListener.selectionChanged(EventObject aEvent)
        {

           // Logger.Instance.Log(LogPriority.DEBUG, this, "[SELECTION INFO] Selection changed in forwarder");

            if (SelectionChanged != null)
            {
                try
                {
                    SelectionChanged.Invoke(this, new ForwardedEventArgs(aEvent));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward selection event: " + e);
                }
            }
        }
        #endregion
    }

    /// <summary>
    /// Forwarder Class for Event listener implementation. 
    /// This should prevent the need of including the OO dlls in projects who use this observer.
    /// </summary>
    public class PropertiesEventForwarder : EventForwarder, XPropertyChangeListener, XPropertiesChangeListener, XVetoableChangeListener, XModifyListener, XAccessibleEventListener
    {
        #region XPropertyChangeListener
        public event EventHandler<ForwardedEventArgs> PropertyChange;
        void XPropertyChangeListener.propertyChange(PropertyChangeEvent evt)
        {
            if (PropertyChange != null)
            {
                try
                {
                    PropertyChange.Invoke(this, new ForwardedEventArgs(evt));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward property change event: " + e);
                }
            }
        }
        #endregion

        #region XPropertiesChangeListener
        public event EventHandler<ForwardedEventArgs> PropertiesChange;
        void XPropertiesChangeListener.propertiesChange(PropertyChangeEvent[] aEvent)
        {
            if (PropertiesChange != null)
            {
                try
                {
                    PropertiesChange.Invoke(this, new ForwardedEventArgs(aEvent));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward properties changed event: " + e);
                }
            }
        }
        #endregion

        #region XVetoableChangeListener
        public event EventHandler<ForwardedEventArgs> VetoableChange;
        void XVetoableChangeListener.vetoableChange(PropertyChangeEvent aEvent)
        {
            if (VetoableChange != null)
            {
                try
                {
                    VetoableChange.Invoke(this, new ForwardedEventArgs(aEvent));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward vetoable property event: " + e);
                }
            }
        }
        #endregion

        #region XModifyListener
        public event EventHandler<ForwardedEventArgs> Modified;
        void XModifyListener.modified(EventObject aEvent)
        {
            if (Modified != null)
            {
                try
                {
                    Modified.Invoke(this, new ForwardedEventArgs(aEvent));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward modify event: " + e);
                }
            }
        }
        #endregion 
    
        #region XAccessibleEventListener
        public event EventHandler<ForwardedEventArgs> NotifyEvent;
        void XAccessibleEventListener.notifyEvent(AccessibleEventObject aEvent)
        {
            // Logger.Instance.Log(LogPriority.DEBUG, this, "[ACCESSIBLE INFO] Accessible event in forwarder");

            if (NotifyEvent != null)
            {
                try
                {
                    NotifyEvent.Invoke(this, new ForwardedEventArgs(aEvent));
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] Could not forward accessible event: " + e);
                }
            }
        }
        #endregion
    }

    public class ForwardedEventArgs : EventArgs
    {
        /// <summary>
        /// Forwarded EventArgs. Because they can't be sent as EventArgs they have to be wrapped.
        /// </summary>
        /// <value>The original event args.</value>
        public Object Args { get; private set; }
        public ForwardedEventArgs(Object args) { Args = args; }
    }


    public abstract class PropertiesEventForwarderBase
    {
        #region Members

        protected readonly PropertiesEventForwarder eventForwarder;

        public PropertiesEventForwarderBase()
        {
            eventForwarder = new PropertiesEventForwarder();
            initEventForwarding();

        }

        #endregion

        #region Event Forwarding

        protected virtual void initEventForwarding()
        {
            if (eventForwarder != null)
            {
                eventForwarder.Disposing += new EventHandler<ForwardedEventArgs>(eventForwarder_disposing);
                eventForwarder.PropertiesChange += new EventHandler<ForwardedEventArgs>(eventForwarder_propertiesChange);
                eventForwarder.PropertyChange += new EventHandler<ForwardedEventArgs>(eventForwarder_propertyChange);
                eventForwarder.VetoableChange += new EventHandler<ForwardedEventArgs>(eventForwarder_vetoableChange);
                eventForwarder.Modified += new EventHandler<ForwardedEventArgs>(eventForwarder_modified);
                eventForwarder.NotifyEvent += new EventHandler<ForwardedEventArgs>(eventForwarder_notifyEvent);
            }
        }

        void eventForwarder_notifyEvent(object sender, ForwardedEventArgs e)
        {
            notifyEvent(e != null ? e.Args as AccessibleEventObject : null);
        }

        void eventForwarder_modified(object sender, ForwardedEventArgs e)
        {
            modified(e != null ? e.Args as EventObject : null);
        }

        protected virtual void eventForwarder_vetoableChange(object sender, ForwardedEventArgs e)
        {
            vetoableChange(e != null ? e.Args as PropertyChangeEvent : null);
        }

        protected virtual void eventForwarder_propertyChange(object sender, ForwardedEventArgs e)
        {
            propertyChange(e != null ? e.Args as PropertyChangeEvent : null);
        }

        protected virtual void eventForwarder_propertiesChange(object sender, ForwardedEventArgs e)
        {
            propertiesChange(e != null ? e.Args as PropertyChangeEvent[] : null);
        }

        protected virtual void eventForwarder_disposing(object sender, ForwardedEventArgs e)
        {
            disposing(e != null ? e.Args as EventObject : null);
        }

        #endregion

        #region Event Handling

        protected virtual void disposing(EventObject Source) { }
        protected virtual void propertyChange(PropertyChangeEvent evt) { }
        protected virtual void vetoableChange(PropertyChangeEvent aEvent) { }
        protected virtual void propertiesChange(PropertyChangeEvent[] aEvent) { }
        protected virtual void modified(EventObject aEvent) { }
        protected virtual void notifyEvent(AccessibleEventObject aEvent) { }
        #endregion

    }

}
