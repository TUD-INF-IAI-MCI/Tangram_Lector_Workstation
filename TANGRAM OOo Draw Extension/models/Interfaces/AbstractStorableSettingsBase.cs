using unoidl.com.sun.star.container;
using unoidl.com.sun.star.ui;

namespace tud.mci.tangram.models.Interfaces
{
    abstract class AbstractStorableSettingsBase
    {
        /// <summary>
        /// Gets the Resource URL for the ToolBar ("private:resource/toolbar/[...]")
        /// </summary>
        /// <value>The toolbar URL.</value>
        public string RecourceUrl { get; protected set; }
        
        private XIndexContainer _settings;
        protected XUIConfigurationManager XCfgMgr { get; set; }
        
        public XIndexContainer Settings
        {
            get { return _settings; }
            set
            {
                _settings = value;
                StoreSettings();
            }
        }
        
        /// <summary>
        /// Stores the settings.
        /// </summary>
        protected void StoreSettings()
        {
            if (XCfgMgr == null)
                return;
            if (XCfgMgr.hasSettings(RecourceUrl))
                XCfgMgr.replaceSettings(RecourceUrl, Settings);
            else
                XCfgMgr.insertSettings(RecourceUrl, Settings);
        }
    }
}
