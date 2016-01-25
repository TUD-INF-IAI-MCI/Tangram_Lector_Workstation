using BrailleIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tud.mci.tangram.audio;
using tud.mci.tangram.TangramLector;
using tud.mci.tangram.TangramLector.Extension;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;

namespace BlindPictionary
{
    public class BlindPictionary : AbstractSpecializedFunctionProxyBase, IInitializable, IInitialObjectReceiver
    {

        #region Member

        WindowManager wm = null;
        AudioRenderer audioRenderer = null;
        OoConnector ooc = null;
        BrailleIOMediator io = null;

        #endregion

        #region Constructor

        public BlindPictionary()
            : base(100000)
        {

        }

        #endregion

        #region IInitializable

        bool IInitializable.Initialize()
        {
            if (wm != null)
            {
                wm.StartFullscreen();
                wm.ZoomTo(WindowManager.BS_FULLSCREEN_NAME, WindowManager.VR_CENTER_NAME, 0.0704111109177273);
                return true;
            }
            return false;
        }

        #endregion

        #region IInitialObjectReceiver

        bool IInitialObjectReceiver.InitializeObjects(params object[] objs)
        {
            bool success = false;
            if (objs != null && objs.Length > 0)
            {
                foreach (var item in objs)
                {
                    if (item != null)
                    {
                        if (item is WindowManager)
                        {
                            wm = item as WindowManager;
                            Active = true;
                        }
                        else if (item is AudioRenderer)
                        {
                            audioRenderer = item as AudioRenderer;
                        }
                        else if (item is OoConnector)
                        {
                            ooc = item as OoConnector;
                        }
                        else if (item is BrailleIOMediator)
                        {
                            io = item as BrailleIOMediator;
                        }
                    }
                }
                success = true;
            }

            return success;
        }

        #endregion

        #region Button combination handling

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            if (sender != null && e != null)
            {
                if (e.ReleasedGenericKeys != null && e.ReleasedGenericKeys.Count > 0)
                {

                    int count = e.ReleasedGenericKeys.Count;

                    switch (count)
                    {
                        #region single button commands
                        case 1:
                            switch (e.ReleasedGenericKeys[0])
                            {
                                case "nsll":// links blättern
                                    e.Cancel = false;
                                    break;
                                case "nsl":// links verschieben
                                    e.Cancel = false;
                                    break;
                                case "nsr":// rechts verschieben
                                    e.Cancel = false;
                                    break;
                                case "nsrr":// rechts blättern
                                    e.Cancel = false;
                                    break;
                                case "nsuu":// hoch blättern
                                    e.Cancel = false;
                                    break;
                                case "nsu":// hoch verschieben
                                    e.Cancel = false;
                                    break;
                                case "nsd":// runter verschieben
                                    e.Cancel = false;
                                    break;
                                case "nsdd":// runter blättern
                                    e.Cancel = false;
                                    break;
                                case "rslu": // kleiner Zoom in
                                    e.Cancel = false;
                                    break;
                                case "rsld": // kleiner Zoom out
                                    e.Cancel = false;
                                    break;
                                case "rsru": // kleiner Zoom in
                                    e.ReleasedGenericKeys[0] = "rslu";
                                    e.Cancel = false;
                                    break;
                                case "rsrd": // kleiner Zoom out
                                    e.ReleasedGenericKeys[0] = "rsld";
                                    e.Cancel = false;
                                    break;
                                case "lr": // Invertieren
                                    e.Cancel = false;
                                    break;
                                case "rl": // Schwellwert minus
                                    e.Cancel = false;
                                    break;
                                case "r": // Schwellwert plus
                                    e.Cancel = false;
                                    break;

                                case "clu": // Sprung nach ganz oben
                                    e.Cancel = false;
                                    break;
                                case "cll": // Sprung nach ganz links
                                    e.Cancel = false;
                                    break;
                                case "cld": // Sprung nach ganz unten
                                    e.Cancel = false;
                                    break;
                                case "clr": // Sprung nach ganz rechts
                                    e.Cancel = false;
                                    break;


                                case "cru": // Sprung nach ganz oben
                                    e.ReleasedGenericKeys[0] = "clu";
                                    e.Cancel = false;
                                    break;
                                case "crl": // Sprung nach ganz links
                                    e.ReleasedGenericKeys[0] = "cll";
                                    e.Cancel = false;
                                    break;
                                case "crd": // Sprung nach ganz unten
                                    e.ReleasedGenericKeys[0] = "cld";
                                    e.Cancel = false;
                                    break;
                                case "crr": // Sprung nach ganz rechts
                                    e.ReleasedGenericKeys[0] = "clr";
                                    e.Cancel = false;
                                    break;

                                default:
                                    e.Cancel = true;
                                    break;
                            }
                            break;
                        #endregion

                        #region 2 button commands
                        case 2:

                            if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsdd", "nsd" }).ToList().Count == 2
                                || e.ReleasedGenericKeys.Intersect(new List<String> { "nsuu", "nsu" }).ToList().Count == 2
                                || e.ReleasedGenericKeys.Intersect(new List<String> { "nsll", "nsl" }).ToList().Count == 2
                                || e.ReleasedGenericKeys.Intersect(new List<String> { "nsrr", "nsr" }).ToList().Count == 2
                                || e.ReleasedGenericKeys.Intersect(new List<String> { "rl", "r" }).ToList().Count == 2
                                )
                            {
                                e.Cancel = false;
                            }
                            else
                            {
                                e.Cancel = true;
                            }
                            break;
                        #endregion

                        default:
                            e.Cancel = true;
                            break;
                    }
                }
            }
        }

        #endregion

    }
}
