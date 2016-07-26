using System;
using System.Windows.Forms;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.TangramLector
{
    class Program
    {
        public static readonly LectorApplicationContext LectorApplicationContext = LectorApplicationContext.Instance;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                run();
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, "Program", "[FATAL ERROR] Unhandled exception occurred", ex);
                Console.WriteLine("FATAL ERROR:\r\n" + ex);
            }
        }

        private static void run()
        {
            // try to get right uno path from registry and set as environment variable
            string unoPath = null;
#if LIBRE
            // access 32bit registry entry for latest LibreOffice for Current User
            Microsoft.Win32.RegistryKey hkcuView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32);
            Microsoft.Win32.RegistryKey hkcuUnoInstallPathKey = hkcuView32.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false);
            if (hkcuUnoInstallPathKey != null && hkcuUnoInstallPathKey.ValueCount > 0)
            {
                unoPath = (string)hkcuUnoInstallPathKey.GetValue(hkcuUnoInstallPathKey.GetValueNames()[hkcuUnoInstallPathKey.ValueCount - 1]);
            }
            else
            {
                // access 32bit registry entry for latest LibreOffice for Local Machine (All Users)
                Microsoft.Win32.RegistryKey hklmView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32);
                Microsoft.Win32.RegistryKey hklmUnoInstallPathKey = hklmView32.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false);
                if (hklmUnoInstallPathKey != null && hklmUnoInstallPathKey.ValueCount > 0)
                {
                    unoPath = (string)hklmUnoInstallPathKey.GetValue(hklmUnoInstallPathKey.GetValueNames()[hklmUnoInstallPathKey.ValueCount - 1]);
                }
            }
#else
            // access 32bit registry entry for latest OpenOffice for Current User
            Microsoft.Win32.RegistryKey hkcuView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32);
            Microsoft.Win32.RegistryKey hkcuUnoInstallPathKey = hkcuView32.OpenSubKey(@"SOFTWARE\OpenOffice\UNO\InstallPath", false);
            if (hkcuUnoInstallPathKey != null && hkcuUnoInstallPathKey.ValueCount > 0)
            {
                unoPath = (string)hkcuUnoInstallPathKey.GetValue(hkcuUnoInstallPathKey.GetValueNames()[hkcuUnoInstallPathKey.ValueCount - 1]);
            }
            else
            {
                // access 32bit registry entry for latest OpenOffice for Local Machine (All Users)
                Microsoft.Win32.RegistryKey hklmView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32);
                Microsoft.Win32.RegistryKey hklmUnoInstallPathKey = hklmView32.OpenSubKey(@"SOFTWARE\OpenOffice\UNO\InstallPath", false);
                if (hklmUnoInstallPathKey != null && hklmUnoInstallPathKey.ValueCount > 0)
                {
                    unoPath = (string)hklmUnoInstallPathKey.GetValue(hklmUnoInstallPathKey.GetValueNames()[hklmUnoInstallPathKey.ValueCount - 1]);
                }
            }
#endif
            if (unoPath != null)
            {
                System.Diagnostics.Debug.WriteLine("Setting UNO_Path Environment Variable: SET UNO_PATH=\"" + unoPath + "\"");
                System.Environment.SetEnvironmentVariable("UNO_PATH", unoPath, EnvironmentVariableTarget.Process);
                System.Diagnostics.Debug.WriteLine("Setting URE_BOOTSTRAP Environment Variable: SET URE_BOOTSTRAP=\"" + "vnd.sun.star.pathname:" + unoPath + "\\fundamental.ini" + "\"");
                System.Environment.SetEnvironmentVariable("URE_BOOTSTRAP", "vnd.sun.star.pathname:" + unoPath + "\\fundamental.ini");

                string pathvar = System.Environment.GetEnvironmentVariable("PATH");
                System.Environment.SetEnvironmentVariable("PATH", pathvar + ";" + unoPath + "\\..\\URE\\bin");

                // LibreOffice5 does not have a URE directory anymore!!
                if (unoPath.Contains("LibreOffice 5"))
                {
                    Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + @";" + unoPath, EnvironmentVariableTarget.Process);
                }
            }

            try
            {
                // start tangram lector tray task manager
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo("TangramLectorTrayTaskManager.exe");
                startInfo.Arguments = "-pid " + System.Diagnostics.Process.GetCurrentProcess().Id.ToString();
                System.Diagnostics.Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Unable to start Tangram Lector Tray Task Manager. Exception:\r\n" + ex.ToString());
            }

            var lgui = new LectorGUI();
            if (lgui != null) lgui.Disposed += new EventHandler(lgui_Disposed);
            try
            {
                 Application.Run(LectorApplicationContext);
            }
            catch (OutOfMemoryException)
            {
                GC.Collect();
            }
            catch (Exception e)
            {
                Logger.Instance.Log(LogPriority.IMPORTANT, "APPLICATION", "[FATAL ERROR] Application exception happens: \n" + e);
                lgui_Disposed(null, null);  
            }
            finally
            {
                              
            }
        }

        /// <summary>
        /// Handles the Disposed event of the lgui control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        static void lgui_Disposed(object sender, EventArgs e)
        {
            try
            {
                DesktopOverlayWindow.Instance.Dispose();
                InteractionManager.Instance.Dispose();
                AudioRenderer.Instance.Dispose();
                Logger.Instance.Dispose();
                WindowManager.Instance.ScreenObserver.Stop();
                Application.Exit();
            }
            catch(System.Exception ex) { Logger.Instance.Log(LogPriority.IMPORTANT, "Program", "[FATAL ERROR] Unhandled exception occurred in lgui_Disposed", ex); }
        }
    }
}
