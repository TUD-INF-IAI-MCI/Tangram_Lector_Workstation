using System;
using System.Collections.Generic;
using System.Linq;
using BrailleIO;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;
using tud.mci.tangram.audio;


namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager
    {
        /// <summary>
        /// Register function proxy that will be called as first specialized function proxy.
        /// With this, window manager has the opportunity to handle button commands before they were sent to other function proxies.
        /// </summary>
        private void registerTopMostSpFProxy()
        {
            TopMostSpecializedFunctionProxy tmsp = new TopMostSpecializedFunctionProxy();
            
            if (ScriptFunctionProxy.Instance != null && tmsp != null)
            {
                tmsp.ButtonCombinationReleased += new EventHandler<ButtonCombinationEventArgs>(tmsp_ButtonCombinationReleased);
                ScriptFunctionProxy.Instance.AddProxy(tmsp);
            }
        }

        void tmsp_ButtonCombinationReleased(object sender, ButtonCombinationEventArgs e)
        {
            // handle top most button combos //

        }
    }

    class TopMostSpecializedFunctionProxy : AbstractSpecializedFunctionProxyBase
    {
        public TopMostSpecializedFunctionProxy() : base(99999)
        {
            Active = true;
        }

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            if (sender != null && e != null)
            {
                fire_buttonCombinationReleased(e.ReleasedGenericKeys);

                // cancel button combos //

            }
        }

        /// <summary>
        /// Occurs when a [button combination was released].
        /// </summary>
        new public event EventHandler<ButtonCombinationEventArgs> ButtonCombinationReleased;

        private void fire_buttonCombinationReleased(List<string> buttons)
        {
            if (ButtonCombinationReleased != null)
            {
                ButtonCombinationReleased.DynamicInvoke(this, new ButtonCombinationEventArgs(buttons));
            }
        }

    }

    class ButtonCombinationEventArgs : EventArgs 
    {
        public List<string> buttons { get; private set; }

        public ButtonCombinationEventArgs(List<string> buttons)
        {
            this.buttons = buttons;
        }
    
    }

}
