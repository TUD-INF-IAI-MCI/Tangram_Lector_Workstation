using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using BrailleIO;
using BrailleIO.Interface;

namespace tud.mci.tangram.TangramLector
{
    public class ShowOffBrailleIOButtonMediator : AbstractBrailleIOButtonMediatorBase
    {
        // BrailleIO_KeyStateChanged_EventArgs
        // raw contains
        // Key = "pressedKeys"
        // Key = "releasedKeys"
        // Key = "timeStampTickCount"
        public ShowOffBrailleIOButtonMediator() : base() { }
        public ShowOffBrailleIOButtonMediator(BrailleIODevice device) : base(device) { }

        /// <summary>Gets all pressed generic buttons.</summary>
        /// <param name="args">The <see cref="T:System.EventArgs"/> instance containing the event data and all current keys states.</param>
        /// <returns>a list of pressed and interpreted generic buttons</returns>
        override public List<String> GetAllPressedGenericButtons(System.EventArgs args)
        {
            List<String> pressedKeys = new List<String>();
            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    return GetAllPressedGenericButtons(kscea.raw);
                }
            }
            return pressedKeys;
        }

        /// <summary>Gets all pressed generic buttons.</summary>
        /// <param name="raw">The raw event data.</param>
        /// <returns>a list of pressed and interpreted generic buttons</returns>
        override public List<String> GetAllPressedGenericButtons(OrderedDictionary raw)
        {
            List<String> pressedKeys = new List<String>();
            if (raw != null)
            {
                var keys = raw.Keys;
                if (raw.Contains("pressedKeys"))
                {
                    List<String> pkl = raw["pressedKeys"] as List<String>;
                    if (pkl != null && pkl.Count > 0)
                    {
                        foreach (var item in pkl)
                        {
                            if (!String.IsNullOrWhiteSpace(item)) { pressedKeys.Add(item.Trim()); }
                        }
                    }
                }
            }
            return pressedKeys;
        }

        /// <summary>
        /// Gets all released generic buttons.
        /// </summary>
        /// <param name="args">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        /// <returns></returns>
        override public List<String> GetAllReleasedGenericButtons(System.EventArgs args)
        {
            List<String> releasedKeys = new List<String>();

            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    return GetAllReleasedGenericButtons(kscea.raw);
                }
            }
            return releasedKeys;
        }

        /// <summary>Gets all released generic buttons.</summary>
        /// <param name="raw">The raw event data.</param>
        /// <returns>a list of released and interpreted generic buttons</returns>
        override public List<String> GetAllReleasedGenericButtons(OrderedDictionary raw)
        {
            List<String> releasedKeys = new List<String>();

            if (raw != null)
            {
                var keys = raw.Keys;
                if (raw.Contains("releasedKeys"))
                {
                    List<String> rkl = raw["releasedKeys"] as List<String>;
                    if (rkl != null && rkl.Count > 0)
                    {
                        foreach (var item in rkl)
                        {
                            if (!String.IsNullOrWhiteSpace(item)) { releasedKeys.Add(item.Trim()); }
                        }
                    }
                }
            }
            return releasedKeys;
        }

        #region IBrailleIOButtonMediator

        /// <summary>
        /// Gets all adapter types this mediator is related to.
        /// </summary>
        /// <returns>
        /// a list of adapter class types this mediator is related to
        /// </returns>
        public override List<Type> GetRelatedAdapterTypes()
        {
            return new List<Type>(1) { typeof(BrailleIO.BrailleIOAdapter_ShowOff) };
        }

        #endregion

    }
}