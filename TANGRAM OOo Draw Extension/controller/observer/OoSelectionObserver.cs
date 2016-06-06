using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.view;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.container;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.frame;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using unoidl.com.sun.star.lang;
using System.Drawing;
using tud.mci.tangram.util;

namespace tud.mci.tangram.controller
{
    public sealed class OoSelectionObserver : IResetable
    {
        private static readonly OoSelectionObserver instance = new OoSelectionObserver();
        private readonly SelectionEventForwarder eventForwarder = new SelectionEventForwarder();

        private OoSelectionObserver() {
            initEventForwarding();
        }

        public static OoSelectionObserver Instance
        {
            get
            {
                return instance;
            }
        }

        #region Event Forwarding

        private void initEventForwarding()
        {
            if (eventForwarder != null)
            {
                eventForwarder.Disposing += new EventHandler<ForwardedEventArgs>(eventForwarder_disposing);
                eventForwarder.SelectionChanged += new EventHandler<ForwardedEventArgs>(eventForwarder_selectionChanged);
            }
        }

        void eventForwarder_selectionChanged(object sender, ForwardedEventArgs e)
        {
            selectionChanged(e != null ? e.Args as EventObject : null);
        }

        void eventForwarder_disposing(object sender, ForwardedEventArgs e)
        {
            disposing(e != null ? e.Args as EventObject : null);
        }

        #endregion

        #region XSelectionChangeListener

        /// <summary>
        /// XSelectionChangeListener call back function. DON'T USE THIS.
        /// </summary>
        /// <param name="aEvent">A event.</param>
        void selectionChanged(unoidl.com.sun.star.lang.EventObject aEvent)
        {

            //FIXME: make this dead
            return;
            //if (aEvent.Source != null)
            //{

            //    System.Diagnostics.Debug.WriteLine("################ SELECTION CHANGE");

            //    try
            //    {
            //        if (aEvent.Source is XSelectionSupplier)
            //        {
            //            TimeLimitExecutor.ExecuteWithTimeLimit(2000, () =>
            //            {
            //                XSelectionSupplier sSupl = aEvent.Source as XSelectionSupplier;
            //                fireSelectionChangedEvent(sSupl, aEvent.Source);
            //            }, "SelectionChendeEventHanlder");
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        Logger.Instance.Log(LogPriority.DEBUG, this, "Exception in XSelectionChangeListener.selectionChanged\r\n" + e);
            //    }
            //}
        }

        /// <summary>
        /// XSelectionChangeListener call back function. DON'T USE THIS.
        /// </summary>
        /// <param name="aEvent">A event.</param>
        void disposing(unoidl.com.sun.star.lang.EventObject Source) { }

        #endregion

        static Object _selectionLock = new Object();
        public static Object GetSelection(XSelectionSupplier selectionSupplier)
        {
            Object selection = null;
            if (selectionSupplier != null)
            {
                lock (_selectionLock)
                {
                    //FIXME: you cannot do this in a time limited execution this will lead to an hang on -- why??????
                    try
                    {
                       // TimeLimitExecutor.WaitForExecuteWithTimeLimit(20000, () =>
                        //{
                            Thread.BeginCriticalRegion();
                            var anySelection = selectionSupplier.getSelection();
                            Thread.EndCriticalRegion();
                        
                            if (anySelection.hasValue())
                            {
                                selection = anySelection.Value;
                            }
                        // }, "GetSelection");
                    }
                    catch (DisposedException)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, "OoSelectionObserver", "Selection supplier disposed");
                        OO.CheckConnection();
                    }
                    catch (ThreadAbortException) { }
                }
            }
            return selection;
        }

        /// <summary>
        /// Occurs when as selection changed.
        /// </summary>
        public event EventHandler<OoSelectionChandedEventArgs> SelectionChanged;

        private void fireSelectionChangedEvent(XSelectionSupplier selectionSupplier, Object sender)
        {
            if (SelectionChanged != null && selectionSupplier != null)
            {
                Object selection = new Object();
                selection = GetSelection(selectionSupplier);
                try
                {
                    SelectionChanged.DynamicInvoke(sender, new OoSelectionChandedEventArgs(selection));
                }
                catch { }
            }
        }

        #region IResetable

        public void Reset()
        {
        }

        #endregion

        /// <summary>
        /// Registers the listener to the given object.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public bool RegisterListenerToElement(Object element)
        {
            XSelectionSupplier supl = element as XSelectionSupplier;
            //TODO: maybe extend this
            if (supl == null && element is XModel)
            {
                XModel model = element as XModel;
                var controller = model.getCurrentController();
                if (controller != null && controller is XSelectionSupplier)
                {
                    supl = controller as XSelectionSupplier;
                }
            }

            if (supl != null)
            {
                try { supl.removeSelectionChangeListener(Instance.eventForwarder); }
                catch { }
                try
                {
                    supl.addSelectionChangeListener(Instance.eventForwarder);
                    return true;
                }
                catch
                {
                    System.Threading.Thread.Sleep(5);
                    try
                    {
                        supl.addSelectionChangeListener(Instance.eventForwarder);
                        return true;
                    }
                    catch { }
                }
            }

            return false;
        }
    }

    /// <summary>
    /// Event Object for selection events
    /// </summary>
    public class OoSelectionChandedEventArgs : EventArgs
    {
        /// <summary>
        /// The selection object
        /// </summary>
        public readonly Object Selection;

        public OoSelectionChandedEventArgs(Object selection)
        {
            this.Selection = selection;
        }

    }



}
