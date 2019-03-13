using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using tud.mci.tangram.models;
using tud.mci.tangram.models.dialogs;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.view;

namespace tud.mci.tangram.util
{
    /// <summary>
    /// Provides functions to handle GUI calls via API.
    /// ATTENTION: this affect the user of the GUI.
    /// </summary>
    public static class OoDispatchHelper
    {
        #region Dispatcher

        /// <summary>
        /// Gets the dispatch helper as anonymous object.
        /// </summary>
        /// <returns>The XDispatchHelper as Object.</returns>
        public static Object GetDispatchHelper_anonymous()
        {
            return GetDispatcher(OO.GetMultiServiceFactory());
        }

        /// <summary>
        /// Provides an easy way to dispatch an URL using one call instead of multiple ones.
        /// Normally a complete dispatch is splitted into different parts: 
        /// - converting and parsing the URL 
        /// - searching for a valid dispatch object available on a dispatch provider 
        /// - dispatching of the URL and it's parameters
        /// </summary>
        /// <param name="_msf">The MSF.</param>
        /// <returns>The dispatch helper or <c>null</c></returns>
        internal static XDispatchHelper GetDispatcher(XMultiServiceFactory _msf = null)
        {
            if (_msf == null) _msf = OO.GetMultiServiceFactory();
            if (_msf != null)
            {
                //Create the Dispatcher
                return _msf.createInstance("com.sun.star.frame.DispatchHelper") as XDispatchHelper;
            }
            return null;
        }

        #endregion

        #region Dispatch Calls

        /// <summary>
        /// Makes a dispatch call to the OpenOffice / LibreOffice GUI.
        /// ATTENTION: this affects the common GUI usage!
        /// </summary>
        /// <param name="commandUrl">The command URL to be executed. Could be e.g. one of those:
        /// <see cref="tud.mci.tangram.util.DispatchURLs"/>, 
        /// <see cref="tud.mci.tangram.util.DispatchURLs_WriterCommands"/>, 
        /// <see cref="tud.mci.tangram.util.DispatchURLs_CalcCommands"/>, 
        /// <see cref="tud.mci.tangram.util.DispatchURLs_DrawImpressCommands"/>, 
        /// <see cref="tud.mci.tangram.util.DispatchURLs_ChartCommands"/>, 
        /// <see cref="tud.mci.tangram.util.DispatchURLs_MathCommands"/>
        /// </param>
        /// <param name="docViewContrl">The document view contrl.</param>
        /// <param name="args">The arguments.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentException">The given control does not supply a dispatch provider (need Interface unoidl.com.sun.star.frame.XDispatchProvider) - docViewContrl</exception>
        public static bool CallDispatch(string commandUrl, Object docViewContrl
            , Object[] args = null)
        {
            if (!(docViewContrl is XDispatchProvider))
            {
                throw new ArgumentException("The given control does not supply a dispatch provider (need Interface unoidl.com.sun.star.frame.XDispatchProvider)", "docViewContrl");
            }

            PropertyValue[] propArgs = ConvertAnonymousArray(args);
            return CallDispatch(commandUrl, docViewContrl as XDispatchProvider, "", 0, propArgs);
        }

        /// <summary>
        /// Executes a dispatch.
        /// The dispatch call is executed in a time limited way <see cref="TimeLimitExecutor"/> (200 ms).
        /// </summary>
        /// <param name="commandUrl">The command URL to dispatch. Describes the 
        /// feature which should be supported by internally used dispatch object. 
        /// <see cref="https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands#Draw.2FImpress_commands"/></param>
        /// <param name="docViewContrl">The document view controller. Points to 
        /// the provider, which should be asked for valid dispatch objects.</param>
        /// <param name="_frame">Specifies the frame which should be the target 
        /// for this request.</param>
        /// <param name="_sFlag">Optional search parameter for finding the frame 
        /// if no special TargetFrameName was used.</param>
        /// <param name="args">Optional arguments for this request They depend on 
        /// the real implementation of the dispatch object.</param>
        /// <remarks>This function is time limited to 200 ms.</remarks>
        internal static bool CallDispatch(string commandUrl, XDispatchProvider docViewContrl, String _frame = "", int _sFlag = 0, PropertyValue[] args = null)
        {
            bool successs = false;

            if (!String.IsNullOrWhiteSpace(commandUrl))
            {
                bool abort = TimeLimitExecutor.WaitForExecuteWithTimeLimit(
                    200,
                    new Action(() =>
                    {
                        var disp = GetDispatcher(OO.GetMultiServiceFactory());
                        if (disp != null)
                        {
                            // A possible result of the executed internal dispatch. 
                            // The information behind this any depends on the dispatch!
                            var result = disp.executeDispatch(docViewContrl, commandUrl, _frame, _sFlag, args);

                            if (result.hasValue() && result.Value is unoidl.com.sun.star.frame.DispatchResultEvent)
                            {
                                var val = result.Value as unoidl.com.sun.star.frame.DispatchResultEvent;
                                if (val != null && val.State == (short)DispatchResultState.SUCCESS)
                                {
                                    successs = true;
                                }
                            }
                        }
                    }),
                    "Dispatch Call");
                successs &= abort;
            }

            return successs;
        }

        #endregion


        /// <summary>
        /// Executes a dispatch on a previously (GUI) selected object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="docSelSuppl">The document sel suppl.</param>
        /// <param name="commandUrl">The command URL.</param>
        /// <param name="docViewContrl">The document view contrl.</param>
        /// <param name="_frame">The frame.</param>
        /// <param name="_sFlag">The s flag.</param>
        /// <param name="args">The arguments.</param>
        /// <exception cref="System.ArgumentException">Objects could not be selected. - obj</exception>
        internal static bool CallDispatchOnObject(
            Object obj, XSelectionSupplier docSelSuppl,
            string commandUrl, XDispatchProvider docViewContrl,
            String _frame = "", int _sFlag = 0, PropertyValue[] args = null)
        {

            if (obj != null && docSelSuppl != null
                && docViewContrl != null && !String.IsNullOrWhiteSpace(commandUrl))
            {
                // you have to select objects to call commands
                try
                {
                    docSelSuppl.select(Any.Get(obj));
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Objects could not be selected.", "obj", ex);
                }

                return CallDispatch(commandUrl, docViewContrl, _frame, _sFlag, args);
            }
            return false;
        }

        #region Selection

        /// <summary>
        /// Gets the selection.
        /// </summary>
        /// <param name="selectProv">The selection provider - commonly this is the document.</param>
        /// <returns>the selection or <c>null</c></returns>
        internal static Object GetSelection(XSelectionSupplier selectProv)
        {
            if (selectProv != null)
            {
                var selection = selectProv.getSelection().Value;
                return selection;
            }
            return null;
        }

        /// <summary>
        /// Sets the selection.
        /// </summary>
        /// <param name="selectProv">The selection provider - commonly this is the document.</param>
        /// <param name="selection">The selection. Use <c>null</c> to reset the selection.</param>
        internal static bool SetSelection(XSelectionSupplier selectProv, Object selection = null)
        {
            if (selectProv != null)
            {
                return selectProv.select(Any.Get(selection));
            }
            return false;
        }

        /// <summary>
        /// Calls an actions with a previously selection change and a selection reset afterwards.
        /// </summary>
        /// <param name="act">The action to call.</param>
        /// <param name="selectProv">The selection provider - commonly this is the document.</param>
        /// <param name="selection">The selection. Use <c>null</c> to reset the selection.</param>
        internal static void ActionWithChangeAndResetSelection(Action act, XSelectionSupplier selectProv, Object selection = null)
        {
            try
            {
                TimeLimitExecutor.ExecuteWithTimeLimit(1000, () =>
                {
                    try
                    {
                        var oldSel = GetSelection(selectProv);
                        Thread.Sleep(100);
                        var succ = SetSelection(selectProv, selection);
                        Thread.Sleep(100);
                        if (succ && act != null)
                        {
                            act.Invoke();
                            if (Thread.CurrentThread.IsAlive)
                                Thread.Sleep(250);
                        }
                        Thread.Sleep(100);
                        SetSelection(selectProv, oldSel);
                    }
                    catch (Exception ex){
                        Logger.Instance.Log(LogPriority.ALWAYS, "OoDispatchHelper", "[FATAL ERROR] Can't call dispatch command with selection - Thread interrupted:", ex);
                    }
                }, "Delete Object");
            }
            catch (ThreadInterruptedException ex)
            {
                Logger.Instance.Log(LogPriority.ALWAYS, "OoDispatchHelper", "[FATAL ERROR] Can't call dispatch command with selection - Thread interrupted:", ex);
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.ALWAYS, "OoDispatchHelper", "[FATAL ERROR] Can't call dispatch command with selection:", ex);
            }
        }

        #endregion

        #region Helper Functions

        /// <summary>
        /// Converts an anonymous object array to PropertyValues.
        /// </summary>
        /// <param name="args">The arguments.</param>
        /// <returns>A PropertyValue array or <c>null</c></returns>
        internal static PropertyValue[] ConvertAnonymousArray(Object[] args)
        {
            PropertyValue[] propArgs = null;
            if (args != null && args.Length > 0)
                propArgs = Array.ConvertAll<object, PropertyValue>(args, (x) => { return x as PropertyValue; });
            return propArgs;
        }

        #endregion

    }

    #region enums

    /// <summary>
    /// possible values for DispatchResultEvent State
    /// </summary>
    enum DispatchResultState : short
    {
        /// <summary>
        /// indicates: dispatch failed
        /// </summary>
        FAILURE = 0,
        /// <summary>
        /// indicates: dispatch was successfully
        /// </summary>
        SUCCESS = 1,
        /// <summary>
        /// indicates: result isn't defined
        /// </summary>
        DONTKNOW = 2
    }

    #endregion 

}
