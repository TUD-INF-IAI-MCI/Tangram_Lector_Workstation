using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.TangramLector.Extension;
using BrailleIO;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;
using tud.mci.tangram;

namespace BrailleDisAdapterMediator
{
    class BrailleDisAdapterSupplier : IBrailleIOAdapterSupplier
    {

        #region Members

        AbstractBrailleIOAdapterBase _bda = null;

        LL LL = new LL(Properties.Resources.Language);

        readonly AudioRenderer audioRenderer = AudioRenderer.Instance;

        #endregion

        #region Constructor

        public BrailleDisAdapterSupplier() { }

        #endregion

        #region IBrailleIOAdapterSupplier

        /// <summary>
        /// Gets the Adapter.
        /// </summary>
        /// <param name="manager"></param>
        /// <returns></returns>
        BrailleIO.Interface.IBrailleIOAdapter IBrailleIOAdapterSupplier.GetAdapter(BrailleIO.Interface.IBrailleIOAdapterManager manager)
        {
            if (manager != null && _bda == null)
            {
                _bda = new BrailleIOBrailleDisAdapter.BrailleIOAdapter_BrailleDisNet(manager);
            }

            return _bda;
        }

        /// <summary>
        /// Initializes the adapter.
        /// </summary>
        /// <returns>
        /// 	<c>true</c>if the initialization was sucessfull, otherwise <c>false</c>
        /// </returns>
        bool IBrailleIOAdapterSupplier.InitializeAdapter()
        {
            if (_bda != null)
            {
                try
                {
                    _bda.initialized -= new EventHandler<BrailleIO.Interface.BrailleIO_Initialized_EventArgs>(_bda_initialized);
                    _bda.closed -= new EventHandler<BrailleIO.Interface.BrailleIO_Initialized_EventArgs>(_bda_closed);

                    _bda.initialized += new EventHandler<BrailleIO.Interface.BrailleIO_Initialized_EventArgs>(_bda_initialized);
                    _bda.closed += new EventHandler<BrailleIO.Interface.BrailleIO_Initialized_EventArgs>(_bda_closed);

                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Instance.Log(LogPriority.IMPORTANT, this, "[ERROR] CAnt register or unregister for/from events:\n" + ex);
                }
            }
            return false;
        }

        /// <summary>
        /// Determines whether this adapter should be used as main adapter.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if [is main adapter]; otherwise, <c>false</c>.
        /// </returns>
        bool IBrailleIOAdapterSupplier.IsMainAdapter()
        {
            return true;
        }

        /// <summary>
        /// Determines whether the specified adapter can be used as a monitor for another adapter.
        /// </summary>
        /// <param name="monitoringAdapter">The adapter type that can be monitored.</param>
        /// <returns>
        /// 	<c>true</c> if the specified adapter is a monitor; otherwise, <c>false</c>.
        /// </returns>
        bool IBrailleIOAdapterSupplier.IsMonitor(out List<Type> monitoringAdapter)
        {
            monitoringAdapter = null;
            return false;
        }

        /// <summary>
        /// Monitors the adapter.
        /// </summary>
        /// <param name="adapter">The adapter to monitor.</param>
        /// <returns>
        /// 	<c>true</c> if the specified adapter can be monitored; otherwise, <c>false</c>.
        /// </returns>
        bool IBrailleIOAdapterSupplier.StartMonitoringAdapter(BrailleIO.Interface.IBrailleIOAdapter adapter)
        {
            return false;
        }

        /// <summary>
        /// Stops the monitoring of the adapter.
        /// </summary>
        /// <param name="adapter">The adapter to monitor.</param>
        /// <returns>
        /// 	<c>true</c> if the specified adapter monitoring was stoped; otherwise, <c>false</c>.
        /// </returns>
        bool IBrailleIOAdapterSupplier.StopMonitoringAdapter(BrailleIO.Interface.IBrailleIOAdapter adapter)
        {
            return true;
        }

        #endregion

        #region Event Handling

        void _bda_closed(object sender, BrailleIO.Interface.BrailleIO_Initialized_EventArgs e)
        {
            audioRenderer.PlaySound(LL.GetTrans("bio.bdis_closed"));
        }

        void _bda_initialized(object sender, BrailleIO.Interface.BrailleIO_Initialized_EventArgs e)
        {
            audioRenderer.PlaySound(LL.GetTrans("bio.bdis_started"));
        }

        #endregion

    }
}
