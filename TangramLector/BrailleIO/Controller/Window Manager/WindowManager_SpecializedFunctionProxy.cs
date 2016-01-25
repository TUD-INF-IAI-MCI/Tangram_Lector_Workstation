using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager : AbstractSpecializedFunctionProxyBase
    {
        override public bool Active
        {
            get
            {
                return true;
            }
            set { }
        }

        override public int ZIndex
        {
            get
            {
                return 0;
            }
            set { }
        }

        void registerAsSpecializedFunctionProxy()
        {
            if (ScriptFunctionProxy.Instance != null)
            {
                ScriptFunctionProxy.Instance.AddProxy(this);
            }
        }

        #region events

        #endregion
    }
}
