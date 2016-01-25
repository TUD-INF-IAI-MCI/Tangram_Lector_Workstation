using System;
using System.Collections.Generic;
using BrailleIO;
using BrailleIO.Interface;

namespace tud.mci.tangram.TangramLector
{
    public class ShowOffBrailleIOButtonMediator : AbstractBrailleIOButtonMediatorBase, IBrailleIOButtonMediator
    {
        // BrailleIO_KeyStateChanged_EventArgs
        // raw contains
        // Key = "pressedKeys"
        // Key = "releasedKeys"
        // Key = "timeStampTickCount"
        public ShowOffBrailleIOButtonMediator() : base() { }
        public ShowOffBrailleIOButtonMediator(BrailleIODevice device) : base(device) { }

        public List<String> GetAllPressedGenericButtons(System.EventArgs args)
        {
            List<String> pressedKeys = new List<String>();
            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.raw != null)
                    {
                        var keys = kscea.raw.Keys;
                        if (kscea.raw.Contains("pressedKeys"))
                        {
                            List<String> pkl = kscea.raw["pressedKeys"] as List<String>;
                            if (pkl != null && pkl.Count > 0)
                            {
                                foreach (var item in pkl)
                                {
                                    if (!String.IsNullOrWhiteSpace(item)) { pressedKeys.Add(item.Trim()); }
                                }
                            }
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
        public List<String> GetAllReleasedGenericButtons(System.EventArgs args)
        {
            List<String> releasedKeys = new List<String>();

            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    //if (kscea.keyCode != BrailleIO_DeviceButtonStates.None)
                    //{
                    //    List<String> ks = BraillIOButtonMediatorHelper.GetAllGeneralReleasedButtons(kscea.keyCode);
                    //    if (ks != null && ks.Count > 0) releasedKeys.AddRange(ks);
                    //}
                    if (kscea.raw != null)
                    {
                        var keys = kscea.raw.Keys;
                        if (kscea.raw.Contains("releasedKeys"))
                        {
                            List<String> rkl = kscea.raw["releasedKeys"] as List<String>;
                            if (rkl != null && rkl.Count > 0)
                            {
                                foreach (var item in rkl)
                                {
                                    if (!String.IsNullOrWhiteSpace(item)) { releasedKeys.Add(item.Trim()); }
                                }
                            }
                        }
                    }
                }
            }
            return releasedKeys;
        }

        public object GetGesture(System.EventArgs args)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public List<BrailleIO_DeviceButton> GetAllPressedGeneralButtons(EventArgs args)
        {
            List<BrailleIO_DeviceButton> pressedKeys = new List<BrailleIO_DeviceButton>();

            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.keyCode != BrailleIO_DeviceButtonStates.None)
                    {
                        pressedKeys = GetAllPressedGeneralButtons(kscea.keyCode);
                    }
                }
            }
            return pressedKeys;
        }

        public List<BrailleIO_DeviceButton> GetAllReleasedGeneralButtons(EventArgs args)
        {
            List<BrailleIO_DeviceButton> releasedKeys = new List<BrailleIO_DeviceButton>();

            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.keyCode != BrailleIO_DeviceButtonStates.None)
                    {
                        releasedKeys = GetAllReleasedGeneralButtons(kscea.keyCode);
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
        public List<Type> GetRelatedAdapterTypes()
        {
            return new List<Type>(1) { typeof(BrailleIO.BrailleIOAdapter_ShowOff) };
        }

        #endregion

    }
}