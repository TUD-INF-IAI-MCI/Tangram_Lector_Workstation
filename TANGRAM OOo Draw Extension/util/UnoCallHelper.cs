using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace tud.mci.tangram.util
{
    //static class UnoCallHelper
    //{
    //    private static IUnoCallReceiver _ucr { get; set; }
    //    public static IUnoCallReceiver getGlobalUnoCallReceiver(NamedPipeServer NPS)
    //    {
    //        if (_ucr == null)
    //        {
    //            _ucr = new PipeUnoCallReceiver(NPS);
    //        }
    //        return _ucr;
    //    }
    //    public static IUnoCallReceiver getGlobalUnoCallReceiver()
    //    {
    //        return _ucr;
    //    }
    //}

    //public class PipeUnoCallReceiver : IUnoCallReceiver
    //{
    //    public PipeUnoCallReceiver(NamedPipeServer NPS)
    //    {
    //        if (NPS != null)
    //        {
    //            NPS.FunctionAvailableEvent += new global::Tangram.ComFramework.Interfaces.FunctionAvailableEventHandler(NPS_FunctionAvailableEvent);
    //            NPS.FunctionCallEvent += new global::Tangram.ComFramework.Interfaces.FunctionCallEventHandler(NPS_FunctionCallEvent);
    //            NPS.InitalizationEvent += new global::Tangram.ComFramework.Interfaces.InitalizationCallEventHandler(NPS_InitalizationEvent);
    //        }
    //    }
    //    #region events
    //    public event InitalizationCallEventHandler InitalizationEvent;
    //    public event FunctionAvailableEventHandler FunctionAvailableEvent;
    //    public event FunctionCallEventHandler FunctionCallEvent;
        
    //    #region event fowarding
    //    void NPS_FunctionCallEvent(string url)
    //    {
    //        if (this.FunctionCallEvent != null)
    //        {
    //            string function = CommandUrlHelper.getPathFromCommandUrl(url);
    //            var parameter = CommandUrlHelper.getParameterFromCommandUrl(url);
    //            this.FunctionCallEvent(function, parameter);
    //        }
    //    }
    //    void NPS_FunctionAvailableEvent(string functionName)
    //    {
    //        if (this.FunctionAvailableEvent != null)
    //        {
    //            this.FunctionAvailableEvent(functionName);
    //        }
    //    }
    //    void NPS_InitalizationEvent()
    //    {
    //        if (this.InitalizationEvent != null)
    //        {
    //            this.InitalizationEvent();
    //        }
    //    }
    //    #endregion

    //    #endregion
    //}

    //public delegate void InitalizationCallEventHandler();
    //public delegate bool FunctionAvailableEventHandler(String functionName);
    //public delegate void FunctionCallEventHandler(String function, Dictionary<String, String> parameter);

    //public interface IUnoCallReceiver
    //{
    //    #region Events
    //    event InitalizationCallEventHandler InitalizationEvent;
    //    event FunctionAvailableEventHandler FunctionAvailableEvent;
    //    event FunctionCallEventHandler FunctionCallEvent;
    //    #endregion
    //}
}
