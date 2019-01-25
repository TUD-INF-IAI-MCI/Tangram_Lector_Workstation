using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager : AbstractSpecializedFunctionProxyBase
    {
        /// <summary>Gets or sets a value indicating whether this <see cref="T:tud.mci.tangram.TangramLector.IInteractionContextProxy"/> is active.</summary>
        /// <value>
        ///   <c>true</c> if active; otherwise, <c>false</c>.</value>
        override public bool Active
        {
            get
            {
                return true;
            }
            set { }
        }

        /// <summary>sorting index for calling.
        /// The higher the value the earlier it is called
        /// in the function proxy chain.</summary>
        /// <value>The index of the z.</value>
        override public int ZIndex
        {
            get
            {
                return 0;
            }
            set { }
        }

        /// <summary>Registers as specialized function proxy.</summary>
        void registerAsSpecializedFunctionProxy()
        {
            if (ScriptFunctionProxy.Instance != null)
            {
                ScriptFunctionProxy.Instance.AddProxy(this);
            }
        }
    }
}
