using System;
using System.Collections.Generic;
using System.Drawing;
using BrailleIO;
using BrailleIO.Interface;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.audio;
using tud.mci.tangram.TangramLector.Extension;
using System.Configuration;
using System.Threading.Tasks;

namespace tud.mci.tangram.TangramLector
{
    public partial class LectorGUI : IDisposable, ILocalizable
    {
        #region Members

        #region private Member

        readonly Logger logger = Logger.Instance;
        BrailleIOMediator io; // Mediator and entry point for the BrailleIO framework 
        //IBrailleIOShowOffMonitor monitor; // debug monitor and simulation of a real pin device
        bool _run = true;

        readonly InteractionManager interactionManager = InteractionManager.Instance;
        readonly AudioRenderer audioRenderer = AudioRenderer.Instance;
        readonly ScriptFunctionProxy functionProxy = ScriptFunctionProxy.Instance;
        readonly OoConnector OpenOffice = OoConnector.Instance;

        static readonly LL LL = new LL(Properties.Resources.Language);

        #endregion

        #region public Member

        public readonly WindowManager windowManager;

        public readonly List<string> patternList;
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="LectorGUI"/> is running.
        /// Setting this to <c>false</c> kills this object!
        /// </summary>
        /// <value><c>true</c> if active; otherwise, <c>false</c> - Setting this to <c>false</c> disposes this object.</value>
        public bool Active { get { return _run; } set { _run = value; if (!value) this.Dispose(); } }

        internal List<IBrailleIOAdapterSupplier> adapterSupplierExtensions = new List<IBrailleIOAdapterSupplier>();

        #endregion

        #endregion

        #region Constructor / Destructor

        public LectorGUI()
        {
            try
            {
                logger.Priority = LogPriority.IMPORTANT;
                try
                {
                    // set log priority by config
                    String appConfigLogPriority = ConfigurationManager.AppSettings["DefaultLogPriority"];
                    if (!String.IsNullOrWhiteSpace(appConfigLogPriority))
                    {
                        LogPriority p = (LogPriority)Enum.Parse(typeof(LogPriority), appConfigLogPriority, true);
                        logger.Priority = p;
                    }
                }
                catch { }


#if DEBUG
                logger.Priority = LogPriority.DEBUG;
#endif

                windowManager = WindowManager.Instance;

                initAudioRenderer();
                initBrailleIOAdapters();

                if (windowManager != null)
                {
                    windowManager.Disposing += new EventHandler(windowManager_Disposing);
                }

                if (functionProxy != null)
                {
                    functionProxy.Initialize(interactionManager);
                    // load specialized script function proxy extensions
                    initializeSSFPExtensions();
                }

                logger.Log(LogPriority.OFTEN, this, "[NOTICE] LGui loaded successfully");
            }
            catch (System.Exception ex)
            {
                logger.Log(LogPriority.IMPORTANT, this, "[FATAL ERROR] Unhandled exception occurred", ex);
            }
        }

        ~LectorGUI()
        {
            logger.Log(LogPriority.OFTEN, this, "[NOTICE] Lector GUI closed");
            this.Dispose();
        }

        #endregion

        #region Initialization Methods

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

        /// <summary>
        /// Initializes the BrailleIO framework.
        /// Sets up a connection Thread for a BrailleDis device.
        /// Sets up a ShowOff-Adapter as debug output and input simulation.
        /// Add the devices to the InteractionManager so it can register to the input events.
        /// Register for Button events from the interaction manager to show them as debug output 
        /// on the ShowOff adapter.
        /// </summary>
        private void initBrailleIOAdapters()
        {
            io = BrailleIOMediator.Instance;
            //Task t = new Task(new Action(() =>
            //{
                if (io != null)
                {
                    List<IBrailleIOAdapter> adapters = LoadAvailableAdapters(io.AdapterManager,
                        this.windowManager,
                        this.interactionManager,
                        this.audioRenderer,
                        this.OpenOffice,
                        this.io
                        );
                }
            //}));
            //t.Start();
        }

        /// <summary>
        /// Initializes the audio renderer. 
        /// Set the Voice to 'Steffi' if available. Play a sound and speak welcome message.
        /// </summary>
        private void initAudioRenderer()
        {
            logger.Log(LogPriority.DEBUG, this, "Installed voices: [" + String.Join("] [", audioRenderer.GetVoices()) + "]");

            // get voice from user-settings
            String stdVoice = Properties.Settings.Default.sound_voice;
            if (!String.IsNullOrWhiteSpace(stdVoice))
            {
                logger.Log(LogPriority.DEBUG, this, "User setting 'sound_voice': " + stdVoice);
                audioRenderer.SetStandardVoiceName(stdVoice);
            }

            // get speed from user-settings
            int vSpeed = Properties.Settings.Default.sound_speed;
            logger.Log(LogPriority.DEBUG, this, "User setting 'sound_speed': " + vSpeed);
            AudioRenderer.Speed = vSpeed;

            // get volume from user-settings
            int vVolume = Properties.Settings.Default.sound_volume;
            logger.Log(LogPriority.DEBUG, this, "User setting 'sound_volume': " + vVolume);
            AudioRenderer.Volume = vVolume;

            audioRenderer.PlayWave(StandardSounds.Start);
            audioRenderer.PlaySound(LL.GetTrans("tangram.lector.bio.welcome"));
        }

        #endregion

        #region IDisposable

        void windowManager_Disposing(object sender, EventArgs e)
        {
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[NOTICE] WindowManager disposed.");
            this.Dispose();
        }

        public event EventHandler Disposed;

        public void Dispose()
        {
            logger.Log("LGui disposed");
            setTangramScreen();
            System.Threading.Thread.Sleep(100);
            try
            {
                if (Disposed != null) Disposed.DynamicInvoke(this, EventArgs.Empty);
            }
            catch { }

            try
            {
                if (io != null && this.io.AdapterManager != null && this.io.AdapterManager.GetAdapters() != null)
                    foreach (var adapter in this.io.AdapterManager.GetAdapters())
                    {
                        if (adapter != null) adapter.Disconnect();
                    }
            }
            catch { }

            try { OoConnector.Instance.Dispose(); }
            catch { }

            try // backstore the audiosettings
            {
                // Properties.Settings.Default.sound_voice = AudioRenderer.Instance.;
                Properties.Settings.Default.sound_speed = AudioRenderer.Speed;
                Properties.Settings.Default.sound_volume = AudioRenderer.Volume;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex) { }


        }

        private void setTangramScreen()
        {
            try
            {
                if (io != null)
                {
                    string name = "TangramSceen";
                    BrailleIOScreen ts = new BrailleIOScreen(name);

                    BrailleIOViewRange tv = new BrailleIOViewRange(0, 0, io.GetDeviceSizeX(), io.GetDeviceSizeY());
                    tv.SetBitmap(Bitmap.FromFile(@"config/pics/tactile_logo.bmp"));
                    ts.AddViewRange(name + "View", tv);
                    io.AddView(name, ts);
                    io.ShowView(name);
                    io.RefreshDisplay();
                    io.RefreshDisplay();
                    System.Threading.Thread.Sleep(10);
                }
            }
            catch { }
        }

        #endregion

        #region Extensibility

        /// <summary>
        /// sub directory in which all extensions are stored
        /// </summary>
        public const string EXTENSION_DIR_NAME = "Extensions";

        /// <summary>
        /// Gets the absolute path to the directory where all extensions are stored.
        /// </summary>
        /// <returns>Absolute path to the extension directory</returns>
        public static string GetExtensionDirectoryPath()
        {
            String path = String.Empty;
            path = extensibility.ExtensionLoader.GetCurrentDllDirectory();
            path += @"\" + EXTENSION_DIR_NAME;
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            return path;
        }

        private static void initializeExtensionInterfaces(Object extensionObj, params object[] iniObjects)
        {
            if (extensionObj is IInitialObjectReceiver)
            {
                try
                {
                    ((IInitialObjectReceiver)extensionObj).InitializeObjects(iniObjects);
                }
                catch (Exception e) { Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[EEROR] can't initialize extension Object as IInitialObjectReceiver:\n" + e); }
            }
            if (extensionObj is IInitializable)
            {
                try { ((IInitializable)extensionObj).Initialize(); }
                catch (Exception e) { Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[EEROR] can't initialize extension Object as IInitializable:\n" + e); }
            }
            if (extensionObj is IRegistrable)
            {
                try { ((IRegistrable)extensionObj).RegisterTo(ScriptFunctionProxy.Instance); }
                catch (Exception e) { Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[EEROR] can't initialize extension Object as IRegistrable:\n" + e); }
            }
        }

        #region Adapter Load

        /// <summary>
        /// sub directory in which all adapters are stored
        /// </summary>
        public const string ADAPTER_DIR_NAME = EXTENSION_DIR_NAME + "\\Adapter";


        /// <summary>
        /// Loads the available <see cref="IBrailleIOAdapter"/> from the adapter extension folder.
        /// </summary>
        /// <param name="manager">The manager.</param>
        /// <param name="iniObjects">The objects sent for initialization.</param>
        /// <returns>List of <see cref="IBrailleIOAdapter"/> from the adapter extension folder.</returns>
        public List<IBrailleIOAdapter> LoadAvailableAdapters(IBrailleIOAdapterManager manager, params object[] iniObjects)
        {
            List<IBrailleIOAdapter> adapters = new List<IBrailleIOAdapter>();

            adapterSupplierExtensions = loadAllAdpaterSuppliers(GetAdapterDirectoryPath());

            if (adapterSupplierExtensions != null && adapterSupplierExtensions.Count > 0)
            {
                foreach (IBrailleIOAdapterSupplier adapterSupplier in adapterSupplierExtensions)
                {
                    if (adapterSupplier != null)
                    {
                        initializeExtensionInterfaces(adapterSupplier, iniObjects);

                        IBrailleIOAdapter adapter = adapterSupplier.GetAdapter(manager);

                        if (adapter != null && !adapters.Contains(adapter))
                        {
                            // add to the list of all available adapters
                            adapters.Add(adapter);

                            initializeExtensionInterfaces(adapter, iniObjects);
                            try
                            {
                                bool initialized = adapterSupplier.InitializeAdapter();
                                if (initialized && manager != null)
                                {
                                    if (adapterSupplier.IsMainAdapter())
                                    {
                                        manager.ActiveAdapter = adapter;
                                    }
                                    else
                                    {
                                        manager.AddAdapter(adapter);
                                    }
                                    InteractionManager.Instance.AddNewDevice(adapter);
                                }
                            }
                            catch (Exception e) { 
                                Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[EEROR] can't initialize and register extension adapter:\n" + e); 
                            }
                        }
                    }
                }
            }

            mapMonitors(adapters, adapterSupplierExtensions);

            return adapters;
        }

        /// <summary>
        /// Gets the directory in which the adapters are stored as extensions.
        /// </summary>
        /// <returns>the path to the directory with all available adapter extensions</returns>
        public static string GetAdapterDirectoryPath()
        {
            String path = String.Empty;
            path = extensibility.ExtensionLoader.GetCurrentDllDirectory();
            path += @"\" + ADAPTER_DIR_NAME;

            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            return path;
        }


        private static List<String> _blocked = new List<string>();
        /// <summary>
        /// Gets the blocked extension list.
        /// </summary>
        /// <returns>list of all configured extension (folder) names, that should be blocked while loading</returns>
        public static List<String> GetBlockedExtensionList()
        {

            if (_blocked.Count < 1)
            {
                try
                {
                    // set log priority by config
                    String blockedString = ConfigurationManager.AppSettings["BlockedEXT"];
                    if (!String.IsNullOrWhiteSpace(blockedString))
                    {
                        string[] b = blockedString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        if (b.Length > 0)
                        {
                            foreach (string item in b)
                            {
                                _blocked.Add("EXT_" + item.Trim());
                            }
                        }
                    }
                }
                catch { }
            }

            return _blocked;
        }


        static List<IBrailleIOAdapterSupplier> loadAllAdpaterSuppliers(string path)
        {
            List<IBrailleIOAdapterSupplier> suppliers = new List<IBrailleIOAdapterSupplier>();

            if (System.IO.Directory.Exists(path))
            {
                // load all available Types implementing the requested interface
                var adptrs = extensibility.ExtensionLoader
                       .LoadAllExtensions
                       (typeof(IBrailleIOAdapterSupplier),
                       path
                       );

                // get list of extensions that should not been loaded
                var blocked = GetBlockedExtensionList();

                // build the adapter suppliers
                if (adptrs != null && adptrs.Count > 0)
                {
                    foreach (var suppl in adptrs)
                    {
                        if (blocked.Contains(suppl.Key)) continue;

                        try
                        {
                            var types = suppl.Value;
                            if (types != null && types.Count > 0)
                            {
                                foreach (Type type in types)
                                {
                                    if (typeof(IBrailleIOAdapterSupplier).IsAssignableFrom(type))
                                    {
                                        IBrailleIOAdapterSupplier o = extensibility.ExtensionLoader.CreateObjectFromType(type) as IBrailleIOAdapterSupplier;
                                        if (o != null)
                                        {
                                            suppliers.Add(o);
                                        }
                                    }
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[ERROR] Can't build instance of adapter supplier:\n" + ex);
                        }
                    }
                }
            }

            return suppliers;
        }

        private static void mapMonitors(List<IBrailleIOAdapter> adapters, List<IBrailleIOAdapterSupplier> suppliers)
        {
            if (adapters != null && adapters.Count > 0 && suppliers != null && suppliers.Count > 0)
            {
                foreach (var supplier in suppliers)
                {
                    if (supplier != null)
                    {
                        List<String> monitorables;
                        if (supplier.IsMonitor(out monitorables))
                        {
                            if (monitorables != null && monitorables.Count > 0)
                            {
                                foreach (String monitorable in monitorables)
                                {
                                    if (!String.IsNullOrWhiteSpace(monitorable))
                                    {
                                        foreach (var adapter in adapters)
                                        {
                                            String name = adapter.GetType().FullName;
                                            if (adapter.GetType().FullName.Equals(monitorable))
                                            {
                                                supplier.StartMonitoringAdapter(adapter);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region Specialized Script Function Proxies Load

        /// <summary>
        /// sub directory in which all adapters are stored
        /// </summary>
        public const string SSFP_DIR_NAME = EXTENSION_DIR_NAME + "\\SpecializedScriptFunctionProxies";

        private void initializeSSFPExtensions()
        {
            //Task t = new Task(new Action(() =>
            //{
            var ssfp = loadAllSpecializedScriptFunctionProxyExtension(GetSSFPDirectoryPath());

            if (functionProxy != null && ssfp != null && ssfp.Count > 0)
            {
                foreach (var item in ssfp)
                {
                    if (item != null)
                    {
                        initializeExtensionInterfaces(item,
                                    this.windowManager,
                                    this.interactionManager,
                                    this.audioRenderer,
                                    this.OpenOffice,
                                    this.io
                            );
                        // register as function proxy
                        functionProxy.AddProxy(item);
                    }
                }
            }
            //}));
            //t.Start();
        }

        /// <summary>
        /// Gets the directory in which the specialized script function proxies are stored as extensions.
        /// </summary>
        /// <returns>the path to the directory with all available specialized script function proxies extensions</returns>
        public static string GetSSFPDirectoryPath()
        {
            String path = String.Empty;
            path = extensibility.ExtensionLoader.GetCurrentDllDirectory();
            path += @"\" + SSFP_DIR_NAME;

            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            return path;
        }

        static List<IInteractionContextProxy> loadAllSpecializedScriptFunctionProxyExtension(string path)
        {
            List<IInteractionContextProxy> ssfp = new List<IInteractionContextProxy>();

            if (System.IO.Directory.Exists(path))
            {
                // load all available Types implementing the requested interface
                var proxies = extensibility.ExtensionLoader
                       .LoadAllExtensions
                       (typeof(IInteractionContextProxy),
                       path
                       );

                // get list of extensions that should not been loaded
                var blocked = GetBlockedExtensionList();

                // build the specialized function proxies
                if (proxies != null && proxies.Count > 0)
                {
                    foreach (var suppl in proxies)
                    {
                        if (blocked.Contains(suppl.Key)) continue;
                        try
                        {
                            var types = suppl.Value;
                            if (types != null && types.Count > 0)
                            {
                                foreach (Type type in types)
                                {
                                    if (typeof(IInteractionContextProxy).IsAssignableFrom(type))
                                    {
                                        IInteractionContextProxy o = extensibility.ExtensionLoader.CreateObjectFromType(type) as IInteractionContextProxy;
                                        if (o != null) { ssfp.Add(o); }
                                    }
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, "LectorGUI", "[ERROR] Can't build instance of specialized function proxy:\n" + ex);
                        }
                    }
                }
            }
            return ssfp;
        }

        #endregion

        #endregion
    }
}