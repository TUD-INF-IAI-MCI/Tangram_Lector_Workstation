using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.TangramLector.Extension;
using BrailleIO;
using tud.mci.tangram;
using BrailleIO.Interface;
using tud.mci.tangram.TangramLector;
using tud.mci.tangram.audio;

namespace ShowOffAdapterMediator
{
    class ShowOffAdapterSupplier : IBrailleIOAdapterSupplier, IInitialObjectReceiver
    {

        #region Members

        IBrailleIOShowOffMonitor monitor;

        AbstractBrailleIOAdapterBase adapter;

        AudioRenderer audioRenderer = null;

        InteractionManager interactionManager = null;

        BrailleIOMediator io = null;

        List<IDisposable> shutDowner = new List<IDisposable>();

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowOffAdapterSupplier"/> class.
        /// </summary>
        public ShowOffAdapterSupplier()
        {
            try
            {
                monitor = new ShowOff();
                ((ShowOff)monitor).SetTitle("Tangram Lektor - Monitor");
                monitor.Disposed += new EventHandler(monitor_Disposed);
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] Cant create instance of ShowOff:\n" + ex);
            }
        }

        #endregion

        #region IBrailleIOAdapterSupplier

        BrailleIO.Interface.IBrailleIOAdapter IBrailleIOAdapterSupplier.GetAdapter(BrailleIO.Interface.IBrailleIOAdapterManager manager)
        {
            if (manager != null && monitor != null)
            {
                adapter = monitor.GetAdapter(manager);
                return adapter;
            }
            return null;
        }

        bool IBrailleIOAdapterSupplier.InitializeAdapter()
        {
            unregisterToEvents();
            if (adapter != null)
            {
                adapter.Synch = true;
            }
            registerToEvents();

            return true;
        }

        bool IBrailleIOAdapterSupplier.IsMainAdapter()
        {
            return false;
        }

        bool IBrailleIOAdapterSupplier.IsMonitor(out List<String> monitoringAdapter)
        {
            monitoringAdapter = new List<String>() { "BrailleIOBrailleDisAdapter.BrailleIOAdapter_BrailleDisNet", "BrailleIOHbsAdapter.BrailleIOAdapter_HyperBrailleS" };
            //monitoringAdapter = new List<Type>();
            return true;
        }

        bool IBrailleIOAdapterSupplier.StartMonitoringAdapter(BrailleIO.Interface.IBrailleIOAdapter adapter)
        {
            if (adapter != null)
            {
                #region BrailleDis events
                ((IBrailleIOAdapterSupplier)this).StopMonitoringAdapter(adapter);
                adapter.touchValuesChanged += new EventHandler<BrailleIO_TouchValuesChanged_EventArgs>(_bda_touchValuesChanged);
                #endregion
            }

            return true;
        }

        /// <summary>
        /// Stops the monitoring of the adapter.
        /// </summary>
        /// <param name="adapter">The adapter to monitor.</param>
        /// <returns>
        /// 	<c>true</c> if the specified adapter monitoring was stopped; otherwise, <c>false</c>.
        /// </returns>
        bool IBrailleIOAdapterSupplier.StopMonitoringAdapter(BrailleIO.Interface.IBrailleIOAdapter adapter)
        {
            if (adapter != null)
            {
                #region BrailleDis events
                adapter.touchValuesChanged -= new EventHandler<BrailleIO_TouchValuesChanged_EventArgs>(_bda_touchValuesChanged);
                #endregion
            }

            return true;
        }

        #endregion
        
        #region Event Hanlding

        void monitor_Disposed(object sender, EventArgs e)
        {
            if (shutDowner != null && shutDowner.Count > 0)
            {
                foreach (var item in shutDowner)
                {
                    try
                    {
                        if (item != null) { item.Dispose(); }
                    }
                    catch { }
                }
            }
        }

        #endregion
        
        #region  Monitoring in ShowOff Adapter

        void _bda_touchValuesChanged(object sender, BrailleIO.Interface.BrailleIO_TouchValuesChanged_EventArgs e)
        {
            if (monitor != null) monitor.PaintTouchMatrix(e.touches);
        }

        void interactionManager_ButtonReleased(object sender, ButtonReleasedEventArgs e)
        {
            if (monitor != null) monitor.MarkButtonAsPressed(e.PressedGenericKeys);
            if (monitor != null) monitor.UnmarkButtons(e.ReleasedGenericKeys);
        }

        void interactionManager_ButtonPressed(object sender, ButtonPressedEventArgs e)
        {
            if (monitor != null) monitor.MarkButtonAsPressed(e.PressedGenericKeys);
        }

        void audioRenderer_Stoped(object sender, EventArgs e)
        {
            if (monitor != null)
            {
                monitor.SetStatusText("[AUDIO OUTPUT STOPED]");
            }
        }

        void audioRenderer_TextSpoken(object sender, AudioRendererTextEventArgs e)
        {
            if (monitor != null && e != null)
            {
                monitor.SetStatusText("[AUDIO OUTPUT] " + e.Text);
            }
        }

        void audioRenderer_FilePlayed(object sender, AudioRendererSoundEventArgs e)
        {
            if (monitor != null && e != null)
            {
                monitor.SetStatusText("[AUDIO OUTPUT SOUND] " + e.SoundName);
            }
        }

        void audioRenderer_Finished(object sender, EventArgs e)
        {
            if (monitor != null)
            {
                monitor.SetStatusText(String.Empty);
            }
        }

        #endregion


        private void registerToEvents()
        {
            if (audioRenderer != null)
            {
                //register to monitoring events
                audioRenderer.FilePlayed += new EventHandler<AudioRendererSoundEventArgs>(audioRenderer_FilePlayed);
                audioRenderer.TextSpoken += new EventHandler<AudioRendererTextEventArgs>(audioRenderer_TextSpoken);
                audioRenderer.Stoped += new EventHandler<EventArgs>(audioRenderer_Stoped);
                audioRenderer.Finished += new EventHandler<EventArgs>(audioRenderer_Finished);
            }

            if (interactionManager != null)
            {
                //for monitoring the pressed buttons in the ShowOffAdapter
                interactionManager.ButtonPressed += new EventHandler<ButtonPressedEventArgs>(interactionManager_ButtonPressed);
                interactionManager.ButtonReleased += new EventHandler<ButtonReleasedEventArgs>(interactionManager_ButtonReleased);
            }
        }

        private void unregisterToEvents()
        {
            if (audioRenderer != null)
            {
                audioRenderer.FilePlayed -= new EventHandler<AudioRendererSoundEventArgs>(audioRenderer_FilePlayed);
                audioRenderer.TextSpoken -= new EventHandler<AudioRendererTextEventArgs>(audioRenderer_TextSpoken);
                audioRenderer.Stoped -= new EventHandler<EventArgs>(audioRenderer_Stoped);
                audioRenderer.Finished -= new EventHandler<EventArgs>(audioRenderer_Finished);
            }

            if (interactionManager != null)
            {
                //for monitoring the pressed buttons in the ShowOffAdapter
                interactionManager.ButtonPressed -= new EventHandler<ButtonPressedEventArgs>(interactionManager_ButtonPressed);
                interactionManager.ButtonReleased -= new EventHandler<ButtonReleasedEventArgs>(interactionManager_ButtonReleased);
            }
        }

        #region IInitialObjectReceiver

        bool IInitialObjectReceiver.InitializeObjects(params object[] objs)
        {
            bool success = true;
            if (objs != null && objs.Length > 0)
            {
                unregisterToEvents();
                foreach (var item in objs)
                {
                    try
                    {
                        if (item != null)
                        {
                            if (item is IDisposable)
                            {
                                shutDowner.Add(item as IDisposable);
                            }

                            if (item is AudioRenderer)
                            {
                                audioRenderer = item as AudioRenderer;
                            }
                            else if (item is InteractionManager)
                            {
                                interactionManager = item as InteractionManager;
                            }
                            else if (item.GetType().FullName.Equals("BrailleIO.BrailleIOMediator")) // is BrailleIO.BrailleIOMediator)
                            {
                                io = item as BrailleIO.BrailleIOMediator;
                                if (io == null)
                                {
                                    throw new NullReferenceException(
                                        @"The referenced Type 'BrailleIO.BrailleIOMediator' seems to be the same but a type conversion wasn't possible. 
                                        This can be caused by adding the type defining reference (dll) twice to the project. 
                                        Build the extension without local copies of overhanded types.");
                                }
                                DebugMonitorTextRenderer dbmtr = new DebugMonitorTextRenderer(monitor, io);
                            }
                        }
                    }
                    catch
                    {
                        success = false;
                    }
                }
                registerToEvents();
            }
            return success;
        }

        #endregion
    }



}

