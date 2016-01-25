using System;
using System.Windows.Forms;

namespace tud.mci.tangram.TangramLector 
{
    /// <summary>
    /// The running application context, holding a desktop overlay window which can be invoked by other threads.
    /// </summary>
    class LectorApplicationContext : ApplicationContext
    {
        private static readonly LectorApplicationContext _instance = new LectorApplicationContext();
        public static LectorApplicationContext Instance { get { return _instance; } }

        public static DesktopOverlayWindow DesktopOverlayWindow;

        static void desktopOverlayWindowDisposed(object sender, EventArgs e)
        {
            System.Console.Error.WriteLine("DesktopOverlayWindow disposed \r\n" + sender + "\r\n" + e);
        }

        public LectorApplicationContext()
        {
            DesktopOverlayWindow = DesktopOverlayWindow.Instance;
        }
    }
}
