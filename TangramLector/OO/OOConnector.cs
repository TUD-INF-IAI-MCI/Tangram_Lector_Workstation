using System;
using System.Threading;
using System.Threading.Tasks;

namespace tud.mci.tangram.TangramLector.OO
{
    /// <summary>
    /// Try to connect to an installed OpenOffice Application and registers some listeners.
    /// </summary>
    public class OoConnector : IDisposable
    {
        /// <summary>
        /// Gets the Open office observer.
        /// This is the main gate to the Open Office API and the observation of DRAW documents. 
        /// </summary>
        /// <value>The observer.</value>
        public OoObserver Observer { get; private set; }
        /// <summary>
        /// Gets the singleton instance for this connector.
        /// </summary>
        /// <value>The instance.</value>
        public static OoConnector Instance { get { return _instance; } }
        private static readonly OoConnector _instance = new OoConnector();
        /// <summary>
        /// flag if the connection thread should been terminated or not
        /// </summary>
        private static bool run = true;

        OoConnector()
        {
            initialize();            
        }

        ~OoConnector(){
            run = false;
        }       

        private void initialize()
        {
            if (tud.mci.tangram.util.OO.ConnectToOO())
            {
                Logger.Instance.Log(LogPriority.MIDDLE, this, "[NOTICE] Connection to OpenOffice established");
                Observer = new OoObserver();
            }
            else
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] Connection to OpenOffice could not been established");
                Logger.Instance.Log(LogPriority.IMPORTANT, this, "[NOTIFICATION] Start connection task");

                Task ts = new Task(
                    () =>
                    {
                        while (!tud.mci.tangram.util.OO.ConnectToOO() && run)
                        { 
                            Thread.Sleep(5000); 
                        }
                        if(run){ Observer = new OoObserver(); }
                    }
                    );
                ts.Start();

            }
        }

        /// <summary>
        /// Resets the observer.
        /// </summary>
        public void ResetObserver()
        {
            if(Observer != null) Observer.Reset();
            else Observer = new OoObserver();
        }

        public void Dispose()
        {
            run = false;
        }
    }
}
