using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager : IFeedbackReceiver
    {
        void IFeedbackReceiver.ReceiveTextualFeedback(string text)
        {
            this.SetDetailRegionContent(text);
        }

        void IFeedbackReceiver.ReceiveTextualNotification(String text)
        {
            this.ShowTemporaryMessageInDetailRegion(text);
        }

        void IFeedbackReceiver.ReceiveAuditoryFeedback(string audio)
        {
            this.audioRenderer.PlaySoundImmediately(audio);
        }
    }
}
