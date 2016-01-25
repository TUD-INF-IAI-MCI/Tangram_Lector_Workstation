using System;
using System.Collections.Generic;
using System.Linq;
using BrailleIO;
using BrailleIO.Interface;

namespace tud.mci.tangram.TangramLector
{
    public class BrailleDisBrailleIOButtonMediator : AbstractBrailleIOButtonMediatorBase, IBrailleIOButtonMediator
    {
        //BrailleIO_KeyStateChanged_EventArgs
        // raw contains
        // Key = "pressedKeys"
        // Key = "releasedKeys"
        // Key = "keyboardState"
        // Key = "timeStampTickCount"
        // Key = "allPressedKeys"
        // Key = "allReleasedKeys"
        public BrailleDisBrailleIOButtonMediator() : base() { }
        public BrailleDisBrailleIOButtonMediator(BrailleIODevice device) : base(device) { }

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
                        if (kscea.raw.Contains("keyboardState"))
                        {
                            object a = kscea.raw["keyboardState"];
                            if (a is HyperBraille.HBBrailleDis.BrailleDisKeyboard)
                            {
                                List<String> allKeys = BrailleIOBrailleDisAdapter.BraillDisButtonInterpreter.toSingleBrailleKeyEventList(((HyperBraille.HBBrailleDis.BrailleDisKeyboard)a).AllKeys);
                                if (allKeys != null && allKeys.Count > 0) pressedKeys.AddRange(allKeys);
                            }
                        }

                        if (kscea.raw.Contains("allPressedKeys"))
                        {
                            List<String> pkl = kscea.raw["allPressedKeys"] as List<String>;
                            if (pkl != null && pkl.Count > 0)
                            {
                                foreach (var item in pkl)
                                {
                                    if (!String.IsNullOrWhiteSpace(item)) { if (pressedKeys != null && !pressedKeys.Contains(item.Trim()))pressedKeys.Add(item.Trim()); }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                //todo: handle key released events
            }

            releasedGenericPressedkeys = lastGenericPressedkeys.Except(pressedKeys).ToList();
            lastGenericPressedkeys = pressedKeys;
            return pressedKeys;
        }

        public List<BrailleIO_DeviceButton> GetAllPressedGeneralButtons(System.EventArgs args)
        {
            List<BrailleIO_DeviceButton> ks = new List<BrailleIO_DeviceButton>();
            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.keyCode != BrailleIO_DeviceButtonStates.None)
                    {
                        ks = GetAllPressedGeneralButtons(kscea.keyCode);
                        if (ks != null && ks.Count > 0) lastGeneralPressedkeys.AddRange(ks);
                    }
                }
            }
            return ks;
        }

        public List<BrailleIO_DeviceButton> GetAllReleasedGeneralButtons(System.EventArgs args)
        {
            List<BrailleIO_DeviceButton> ks = new List<BrailleIO_DeviceButton>();
            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.keyCode != BrailleIO_DeviceButtonStates.None)
                    {
                        ks = GetAllReleasedGeneralButtons(kscea.keyCode);
                        //if (ks != null && ks.Count > 0) lastGeneralPressedkeys.AddRange(ks);
                    }
                }
            }
            return ks;
        }

        public List<String> GetAllReleasedGenericButtons(System.EventArgs args)
        {
            List<String> releasedKeys = new List<String>();
            if (args != null && args is BrailleIO_KeyStateChanged_EventArgs)
            {
                BrailleIO_KeyStateChanged_EventArgs kscea = args as BrailleIO_KeyStateChanged_EventArgs;
                if (kscea != null)
                {
                    if (kscea.raw != null)
                    {
                        var keys = kscea.raw.Keys;
                        if (kscea.raw.Contains("releasedKeys"))
                        {
                            List<String> pkl = kscea.raw["releasedKeys"] as List<String>;
                            if (pkl != null && pkl.Count > 0)
                            {
                                foreach (var item in pkl)
                                {
                                    if (!String.IsNullOrWhiteSpace(item)) { if (releasedKeys != null && !releasedKeys.Contains(item.Trim()))releasedKeys.Add(item.Trim()); }
                                }
                            }
                        }

                        if (kscea.raw.Contains("allReleasedKeys"))
                        {
                            List<String> pkl = kscea.raw["allReleasedKeys"] as List<String>;
                            if (pkl != null && pkl.Count > 0)
                            {
                                foreach (var item in pkl)
                                {
                                    if (!String.IsNullOrWhiteSpace(item)) { if (releasedKeys != null && !releasedKeys.Contains(item.Trim()))releasedKeys.Add(item.Trim()); }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                //todo: handle key pressed events
            }
            return releasedKeys;
        }

        public object GetGesture(System.EventArgs args)
        {
            //TDOD:
            return null;
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
            return new List<Type>(1) { typeof(BrailleIOBrailleDisAdapter.BrailleIOAdapter_BrailleDisNet) };
        }

        #endregion

    }
}