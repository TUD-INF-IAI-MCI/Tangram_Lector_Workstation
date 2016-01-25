using System;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TangramLectorTrayTaskManager
{
    public partial class TangramLectorTaskManager : Form
    {
        // the tangram lector process id
        public int TangramLectorProcessID { get; private set; }

        // singleton
        public static TangramLectorTaskManager Instance;

        // main method, process is supposed to be started by tangram lector at its start
        [STAThread]
        public static void Main(string[] args)
        {
            // new instance
            Instance = new TangramLectorTaskManager();

            // default for no given tangram lector PID
            Instance.TangramLectorProcessID = -1;

            // try to parse arguments to get PID
            if (args != null)
            {
                string arg = String.Empty;
                for (int i = 0; i < args.Length; i++)
                {
                    arg = args[i];
                    if (arg.ToLower().Equals("-pid") && i < args.Length - 1)
                    {
                        try { Instance.TangramLectorProcessID = int.Parse(args[i + 1]); }
                        catch (FormatException) { }
                        catch (OverflowException) { }
                    }
                }
            }
            if (Instance.TangramLectorProcessID>=0)
            {
                // run user interface loop 
                Application.Run(Instance);
            }
            else
            {
                Instance.checkRunningTangramLectorTimer.Stop();
                Instance.exitThis(ErrorCodes.NO_PID);
            }
        }

        // sets up the user interface (invisible form, but a tray icon with menu, see forms designer)
        public TangramLectorTaskManager()
        {
            InitializeComponent();
            Visible = false;
            // run the timer that periodically checks for the tangram process 
            checkRunningTangramLectorTimer.Start();
        }

        // this keeps form invisible on start
        protected override void SetVisibleCore(bool value)
        {
            if (value && !IsHandleCreated) CreateHandle();  // Ensure Load event runs
            value = false;
            base.SetVisibleCore(value);
        }

        // handler for menu entry "Exit"
        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            if (TangramLectorProcessID>=0)
            {
                Process p = Process.GetProcessById(TangramLectorProcessID);
                if (p != null)
                try
                {
                    p.Kill();
                    exitThis(ErrorCodes.OK);
                }
                catch(Exception ex)
                {
                    checkRunningTangramLectorTimer.Stop();
                    exitThis(ErrorCodes.UNABLE_TO_KILL, ex);
                }
            }
        }

        private void restartMenuItem_Click(object sender, EventArgs e)
        {
            Process p = Process.GetProcessById(TangramLectorProcessID);
            try
            {
                p.Kill();
                System.Diagnostics.ProcessStartInfo newStartInfo = new System.Diagnostics.ProcessStartInfo("TangramLector.exe");
                try
                {
                    Process.Start(newStartInfo);
                    exitThis(ErrorCodes.OK);
                }
                catch (Exception ex)
                {
                    exitThis(ErrorCodes.UNABLE_TO_RESTART,ex);
                }
            }
            catch (Exception ex)
            {
                checkRunningTangramLectorTimer.Stop();
                exitThis(ErrorCodes.UNABLE_TO_KILL,ex);
            }
        }

        // handler for timer, checks if tangram is still running
        private void checkRunningTangramLector_Tick(object sender, EventArgs e)
        {
            if (TangramLectorProcessID < 0)
            {
                checkRunningTangramLectorTimer.Stop();
                exitThis(ErrorCodes.NO_PID);
                return;
            }
            
            try
            {
                Process p = Process.GetProcessById(TangramLectorProcessID);
                if (p.HasExited)
                {
                    p.Kill();
                    exitThis(ErrorCodes.ALREADY_KILLED);
                }
            }
            // if an exception is raised the process could not be found, so exit
            catch (Exception ex)
            {
                exitThis(ErrorCodes.OK, ex);
            }
        }

        enum ErrorCodes
        {
            OK,
            UNKNOWN_ERROR,
            NO_PID,
            ALREADY_KILLED,
            UNABLE_TO_KILL,
            UNABLE_TO_RESTART
        };

        private void exitThis(ErrorCodes errorCode, Exception ex = null)
        {
            checkRunningTangramLectorTimer.Stop();
            switch (errorCode)
            {
                case ErrorCodes.OK:
                    break;
                case ErrorCodes.NO_PID:
                    MessageBox.Show("Tangram Lector nicht gefunden.\r\nDie Prozess-Nummer muss beim Aufruf im Format \"-pid 9999\" übergeben werden (ohne Anführungszeichen)."
                        + (ex != null ? "\r\n" + ex.ToString() : String.Empty),
                        Instance.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case ErrorCodes.ALREADY_KILLED:
                    MessageBox.Show("Tangram-Prozess (PID " + TangramLectorProcessID + " ) ist bereits beendet."
                        + (ex != null ? "\r\n" + ex.ToString() : String.Empty),
                        Instance.Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case ErrorCodes.UNABLE_TO_KILL:
                    MessageBox.Show("Konnte Tangram-Prozess (PID " + TangramLectorProcessID + " ) nicht beenden."
                        + (ex != null ? "\r\n" + ex.ToString() : String.Empty),
                        Instance.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case ErrorCodes.UNABLE_TO_RESTART:
                    MessageBox.Show("Konnte Tangram-Prozess nicht neu starten."
                        + (ex != null ? "\r\n" + ex.ToString() : String.Empty),
                        Instance.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case ErrorCodes.UNKNOWN_ERROR:
                default:
                     MessageBox.Show("Unbekannter Fehler. Tangram Lector Task Manager wird beendet."
                        + (ex != null ? "\r\n" + ex.ToString() : String.Empty),
                        Instance.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
            trayIcon.Visible = false;
            trayIcon.Dispose();
            Application.Exit();
        }


    }
}
