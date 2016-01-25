using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;
using tud.mci.tangram.TangramLector.Extension;
using tud.mci.tangram.TangramLector;
using BrailleIO;
using tud.mci.tangram.audio;
using tud.mci.tangram;
using tud.mci.tangram.TangramLector.OO;
using System.Drawing;
using tud.mci.tangram.controller.observer;
using BrailleIO.Interface;

namespace ZoomingComparisonTest
{
    class ZoomingComparisonTest : AbstractSpecializedFunctionProxyBase, IInitializable, IInitialObjectReceiver
    {

        #region Member

        WindowManager wm = null;
        AudioRenderer audioRenderer = null;
        OoConnector ooc = null;
        BrailleIOMediator io = null;

        #endregion

        public ZoomingComparisonTest() : base(90000)
        {

        }


        #region IInitializable

        bool IInitializable.Initialize()
        {
            if (wm != null)
            {
                wm.StartFullscreen();
                wm.ZoomTo(WindowManager.BS_FULLSCREEN_NAME, WindowManager.VR_CENTER_NAME, 0.0704111109177273);
                addDebugViewRange(wm.GetVisibleScreen());
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
                        if (item is WindowManager) { 
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

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            if (sender != null && e != null)
            {

                #region cancel button combos
                
                if (e.ReleasedGenericKeys != null && e.ReleasedGenericKeys.Count == 1)
                {
                    switch (e.ReleasedGenericKeys[0])
                    {
                        // cancel big zoom and other functions which are overwritten //
                        case "rsru":
                            e.Cancel = true;
                            break;
                        case "rsrd":
                            e.Cancel = true;
                            break;
                        case "k2":
                            e.Cancel = true;
                            break;
                        case "k3":
                            e.Cancel = true;
                            break;
                        case "k4":
                            e.Cancel = true;
                            break;
                        case "k7":
                            e.Cancel = true;
                            break;

                        // no cancel of zooming and panning functions //
                        case "rslu":
                            if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                            {
                                e.Cancel = false;
                            }
                            else e.Cancel = true;
                            break;
                        case "rsld":
                            if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                            {
                                e.Cancel = false;
                            }
                            else e.Cancel = true;
                            break;
                        case "nsll":
                            e.Cancel = false;
                            break;
                        case "nsl":
                            e.Cancel = false;
                            break;
                        case "nsr":
                            e.Cancel = false;
                            break;
                        case "nsrr":
                            e.Cancel = false;
                            break;
                        case "nsuu":
                            e.Cancel = false;
                            break;
                        case "nsu":
                            e.Cancel = false;
                            break;
                        case "nsd":
                            e.Cancel = false;
                            break;
                        case "nsdd":
                            e.Cancel = false;
                            break;
                        case "k1":
                            if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                            {
                                e.Cancel = false;
                            }
                            else e.Cancel = true;
                            break;
                        case "clu":
                            e.Cancel = false;
                            break;
                        case "cll":
                            e.Cancel = false;
                            break;
                        case "clr":
                            e.Cancel = false;
                            break;
                        case "cld":
                            e.Cancel = false;
                            break;
                        case "clc":
                            e.Cancel = false;
                            break;

                        // cancel all other functions //
                        default:
                            e.Cancel = true;
                            break;
                    }
                }

                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsdd", "nsd" }).ToList().Count == 2)
                {
                    e.Cancel = false;
                }
                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsuu", "nsu" }).ToList().Count == 2)
                {
                    e.Cancel = false;
                }
                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsll", "nsl" }).ToList().Count == 2)
                {
                    e.Cancel = false;
                }
                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "nsrr", "nsr" }).ToList().Count == 2)
                {
                    e.Cancel = false;
                }

                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "k4", "k5" }).ToList().Count == 2)
                {
                    if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                    {
                        e.Cancel = false;
                    }
                    else e.Cancel = true;
                }

                else if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k2" }).ToList().Count == 2)
                {
                    if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                    {
                        e.Cancel = false;
                    }
                    else e.Cancel = true;
                }

                else if (e.ReleasedGenericKeys != null && e.ReleasedGenericKeys.Count == 5)
                {
                    if (e.ReleasedGenericKeys.Intersect(new List<String> { "k1", "k3", "k4", "k6", "k8" }).ToList().Count == 5)
                    {
                        if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                        {
                            e.Cancel = false;
                        }
                        else e.Cancel = true;
                    }
                    else e.Cancel = true;
                }

                else e.Cancel = true;

                #endregion

                #region handle top most button combos

                if (wm != null && audioRenderer != null)
                {
                    BrailleIOScreen vs = wm.GetVisibleScreen();
                    BrailleIOViewRange center = wm.GetActiveCenterView(vs);

                    if (e.ReleasedGenericKeys != null && e.ReleasedGenericKeys.Count == 1)
                    {
                        switch (e.ReleasedGenericKeys[0])
                        {
                            case "rsru": // kleiner Zoom in zum Mittelpunkt
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    if (vs != null && center != null)
                                    {
                                        if (zoomInMiddleWithFactor(vs.Name, center.Name, 1.5))
                                        {
                                            audioRenderer.PlaySoundImmediately("vergrößert auf " + getZoomPercentageBasedOnPrintZoom(vs.Name, center.Name) + " Prozent");
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] midpoint zoom in");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                }
                                return;
                            case "rsrd": // kleiner Zoom out zum Mittelpunkt
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    if (vs != null && center != null)
                                    {
                                        if (zoomInMiddleWithFactor(vs.Name, center.Name, (double)1 / 1.5))
                                        {
                                            audioRenderer.PlaySoundImmediately("verkleinert auf " + getZoomPercentageBasedOnPrintZoom(vs.Name, center.Name) + " Prozent");
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] midpoint zoom out");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                }
                                return;
                            case "k2": // Druckzoom zum Mittelpunkt
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    if (vs != null && center != null)
                                    {
                                        if (zoomInMiddleTo(vs.Name, WindowManager.VR_CENTER_NAME, WindowManager.GetPrintZoomLevel()))
                                        {
                                            if (ooc != null && ooc.Observer != null)
                                            {
                                                if (ooc.Observer.TextRendererHook != null)
                                                {
                                                    ooc.Observer.TextRendererHook.Active = false;
                                                }
                                            }
                                            audioRenderer.PlaySoundImmediately("Ausschnitt auf Druckzoom optimiert");
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] midpoint print zoom");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                }
                                return;
                            case "k1": // Druckzoom ohne Anzeige von Braille
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    if (ooc != null && ooc.Observer != null)
                                    {
                                        if (ooc.Observer.TextRendererHook != null)
                                        {
                                            ooc.Observer.TextRendererHook.Active = false;
                                        }
                                    }
                                }
                                return;
                            case "k3": // Breite/Höhe einpassen
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    if (vs != null && !vs.Name.Equals(WindowManager.BS_MINIMAP_NAME))
                                    {
                                        if (wm.ZoomTo(vs.Name, WindowManager.VR_CENTER_NAME, 0.0704111109177273)) // entspricht genau Druckzoom (100%) geteilt durch 1,5 (verwendeter Zoomfaktor)
                                        {
                                            audioRenderer.PlaySoundImmediately("Höhe eingepasst");
                                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] zoom fit to height");
                                        }
                                        else audioRenderer.PlayWaveImmediately(StandardSounds.End);
                                    }
                                }
                                return;
                            case "k4": // randomisierte Platzierung des fokussierten Elements im sichtbaren Bereich
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    moveToRandomOffset();
                                }
                                return;
                            case "k6":
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "START TIME");
                                }
                                return;
                            case "k7":
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "STOPP TIME");
                                    audioRenderer.PlaySoundImmediately("Stopp time");
                                    logData();
                                }
                                return;
                            case "k8":
                                if (e.Device.AdapterType == "BrailleIO.BrailleIOAdapter_ShowOff")
                                {
                                    Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "STOPP TIME");
                                    audioRenderer.PlaySoundImmediately("Stopp time");
                                    logData();
                                }
                                return;
                        }
                    }
                }

                #endregion
            }
        }


        #region functions

        /// <summary>
        /// Get zoom level in percentage. Print zoom is defined as 100%.
        /// </summary>
        /// <param name="viewName">name of the current BrailleIOScreen</param>
        /// <param name="viewRangeName">name of the current view range</param>
        /// <returns>current zoom level in % or Double.NaN if there is no zoom level</returns>
        private double getZoomPercentageBasedOnPrintZoom(string viewName, string viewRangeName)
        {
            if (io != null && io.GetView(viewName) as BrailleIOScreen != null)
            {
                BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
                if (vr != null)
                {
                    double zoom = vr.GetZoom();
                    double zoomPercentage = zoom / WindowManager.GetPrintZoomLevel();
                    zoomPercentage = (int)(zoomPercentage * 100);
                    return zoomPercentage;
                }
            }
            return Double.NaN;
        }

        /// <summary>
        /// Zooms the specified view with the given factor.
        /// </summary>
        /// <param name="viewName">Name of the view.</param>
        /// <param name="viewRangeName">Name of the view range.</param>
        /// <param name="factor">The zoom factor.</param>
        protected bool zoomInMiddleWithFactor(string viewName, string viewRangeName, double factor)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return false;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);

            bool success = false;
            if (vr != null)
            {
                double oldZoom = vr.GetZoom();
                if (oldZoom > 0)
                {
                    double newZoom = oldZoom * factor;
                    success = midpointZoom(vr, oldZoom, newZoom);
                }
                else return false;
            }
            else return false;

            this.io.RefreshDisplay();
            return success;
        }

        /// <summary>
        /// Zooms the specified view to the given zoom level.
        /// </summary>
        /// <param name="viewName">Name of the view.</param>
        /// <param name="viewRangeName">Name of the view range.</param>
        /// <param name="zoomLevel">The zoom level.</param>
        protected bool zoomInMiddleTo(string viewName, string viewRangeName, double zoomLevel)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return false;

            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);

            bool success = false;
            if (vr != null)
            {
                var oldZoom = vr.GetZoom();
                var newZoom = zoomLevel;
                success = midpointZoom(vr, oldZoom, newZoom);
            }
            else return false;

            io.RefreshDisplay();
            return success;
        }

        /// <summary>
        /// Zooming that depends on the midpoint of the ViewRange
        /// </summary>
        /// <param name="vr"></param>
        /// <param name="oldZoom"></param>
        /// <param name="newZoom"></param>
        /// <returns></returns>
        private bool midpointZoom(BrailleIOViewRange vr, double oldZoom, double newZoom)
        {
            if (vr != null)
            {
                Point oldCenter = new Point();
                var oldvrdin = vr.ContentBox;

                // Prüfung auf größte Zoomstufe
                if (newZoom > BrailleIOViewRange.MAX_ZOOM_LEVEL)
                {
                    if (oldZoom == BrailleIOViewRange.MAX_ZOOM_LEVEL) return false;
                    newZoom = BrailleIOViewRange.MAX_ZOOM_LEVEL;
                }
                // Prüfung auf kleinste Zoomstufe
                if (vr.ContentBox.Height >= vr.ContentHeight && vr.ContentBox.Width >= vr.ContentWidth)
                {
                    if (oldZoom >= newZoom) return false;
                    else
                    {
                        oldCenter = new Point(
                            (int)Math.Round(((double)vr.ContentWidth / 2) + (vr.GetXOffset() * -1)),
                            (int)Math.Round(((double)vr.ContentHeight / 2) + (vr.GetYOffset() * -1))
                        );
                    }
                }
                else
                {
                    // central point of center region as center for zooming
                    oldCenter = new Point(
                        (int)Math.Round(((double)oldvrdin.Width / 2) + (vr.GetXOffset() * -1)),
                        (int)Math.Round(((double)oldvrdin.Height / 2) + (vr.GetYOffset() * -1))
                        );
                }

                if (newZoom > 0)
                {
                    double zoomRatio = newZoom / oldZoom;
                    Point newCenter = new Point(
                        (int)Math.Round(oldCenter.X * zoomRatio),
                        (int)Math.Round(oldCenter.Y * zoomRatio)
                        );

                    Point newOffset = new Point(
                        (int)Math.Round((newCenter.X - ((double)oldvrdin.Width / 2)) * -1),
                        (int)Math.Round((newCenter.Y - ((double)oldvrdin.Height / 2)) * -1)
                        );

                    vr.SetZoom(newZoom);
                    vr.MoveTo(new Point(Math.Min(newOffset.X, 0), Math.Min(newOffset.Y, 0)));
                }
                else // set to smallest zoom level
                {
                    vr.SetZoom(-1);
                    vr.SetXOffset(0);
                    vr.SetYOffset(0);
                }

                // check for correct panning
                if (vr.GetXOffset() > 0) { vr.SetXOffset(0); }
                if (vr.GetYOffset() > 0) { vr.SetYOffset(0); }

                if ((vr.ContentWidth + vr.GetXOffset()) < vr.ContentBox.Width)
                {
                    int maxOffset = Math.Min(0, vr.ContentBox.Width - vr.ContentWidth);
                    vr.SetXOffset(maxOffset);
                }

                if ((vr.ContentHeight + vr.GetYOffset()) < vr.ContentBox.Height)
                {
                    int maxOffset = Math.Min(0, vr.ContentBox.Height - vr.ContentHeight);
                    vr.SetYOffset(maxOffset);
                }

                return true;
            }
            return false;

        }

        /// <summary>
        /// Shows a randomized section in which the focused element is visible.
        /// </summary>
        private void moveToRandomOffset()
        {
            OoConnector ooc = OoConnector.Instance;
            if (ooc != null && ooc.Observer != null && wm != null)
            {
                OoShapeObserver shape = ooc.Observer.GetLastSelectedShape();
                if (shape != null)
                {
                    BrailleIOScreen vs = wm.GetVisibleScreen();
                    if (vs != null)
                    {
                        BrailleIOViewRange center = wm.GetActiveCenterView(vs);
                        if (center != null)
                        {
                            Rectangle boundings = shape.GetRelativeScreenBoundsByDom();
                            double zoom = center.GetZoom();
                            Point offset = calculateRandomOffset(boundings, zoom);

                            if (!(center.ContentBox.Height >= center.ContentHeight && center.ContentBox.Width >= center.ContentWidth))
                            {
                                wm.MoveToPosition(vs.Name, center.Name, offset);
                            }

                            offset = center.OffsetPosition;

                            Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "MOVE TO RANDOM OFFSET");
                            zoomingTestLogging(vs, center, shape, boundings, zoom, offset);
                            audioRenderer.PlaySoundImmediately("Zufallsoffset");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Calculates a random offset. 
        /// With this offset the focused element will stay in visible section of the view range. 
        /// </summary>
        /// <returns></returns>
        private Point calculateRandomOffset(Rectangle boundingBox, double zoomFactor)
        {
            Point offset = new Point();
            Random random = new Random();

            int x_min = boundingBox.Left - (int)((double)105 / zoomFactor);
            int x_max = boundingBox.Right - (int)((double)15 / zoomFactor);
            int y_min = boundingBox.Top - (int)((double)45 / zoomFactor);
            int y_max = boundingBox.Bottom - (int)((double)15 / zoomFactor);

            double randX = (double)random.Next(x_min, x_max);
            double randY = (double)random.Next(y_min, y_max);
            offset.X = (int)(-randX * zoomFactor);
            offset.Y = (int)(-randY * zoomFactor);

            return offset;
        }

        /// <summary>
        /// Writes important data for test in log file.
        /// </summary>
        /// <param name="vs"></param>
        /// <param name="vr"></param>
        /// <param name="shape"></param>
        /// <param name="boundings"></param>
        /// <param name="zoom"></param>
        /// <param name="offset"></param>
        private void zoomingTestLogging(BrailleIOScreen vs, BrailleIOViewRange vr, OoShapeObserver shape, Rectangle boundings, double zoom, Point offset)
        {
            Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "test id = " + DateTime.Now.Ticks.ToString());
            if (vs != null && vr != null && wm != null) Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "zoom percent level = " + getZoomPercentageBasedOnPrintZoom(vs.Name, vr.Name).ToString() + "%");
            if (zoom > 0 && offset != null) Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "zoom level = " + zoom.ToString() + "; offset = " + offset.ToString());
            if (shape != null) Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "shape name = " + shape.Name);
            if (boundings != null)
            {
                Rectangle pinBoundings = new Rectangle((int)(boundings.X * zoom), (int)(boundings.Y * zoom), (int)(boundings.Width * zoom), (int)(boundings.Height * zoom)); // bounding box in pin coordinates (relative to the content origin)
                Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "bounding box in pins = " + pinBoundings.ToString());

                Point shapeCenter = new Point(pinBoundings.X + (pinBoundings.Width / 2), pinBoundings.Y + (pinBoundings.Height / 2));
                shapeCenter = new Point(shapeCenter.X + offset.X, shapeCenter.Y + offset.Y); // shape center in absolute pin device coordinates
                if (shapeCenter != null) Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "shape center position = " + shapeCenter.ToString());
            }
        }


        /// <summary>
        /// Gets all data for logging and calls the test logging method.
        /// </summary>
        private void logData()
        {
            OoConnector ooc = OoConnector.Instance;
            if (ooc != null && ooc.Observer != null && wm != null)
            {
                OoShapeObserver shape = ooc.Observer.GetLastSelectedShape();
                if (shape != null)
                {
                    BrailleIOScreen vs = wm.GetVisibleScreen();
                    if (vs != null)
                    {
                        BrailleIOViewRange center = wm.GetActiveCenterView(vs);
                        if (center != null)
                        {
                            Rectangle boundings = shape.GetRelativeScreenBoundsByDom();
                            double zoom = center.GetZoom();
                            Point offset = center.OffsetPosition;

                            zoomingTestLogging(vs, center, shape, boundings, zoom, offset);
                        }
                    }
                }
            }
        }


        #region Debugmessage

        private const string VR_NAME = "Debug";
        private BrailleIOViewRange debugVr = null;
        private BrailleIOScreen activeScreen = null;

        private void addDebugViewRange(BrailleIOScreen screen)
        {
            if (screen != null && wm != null)
            {
                activeScreen = screen;
                debugVr = new BrailleIOViewRange(0, 0, 120, 4);
                debugVr.SetZIndex(-10);
                debugVr.Name = VR_NAME;
                debugVr.SetText(getZoomPercentageBasedOnPrintZoom(screen.Name, WindowManager.VR_CENTER_NAME).ToString() + "%");
                screen.AddViewRange(VR_NAME, debugVr);
                var center = wm.GetActiveCenterView(screen);
                center.PropertyChanged += new EventHandler<BrailleIOPropertyChangedEventArgs>(center_PropertyChanged);
            }
        }

        void center_PropertyChanged(object sender, BrailleIOPropertyChangedEventArgs e)
        {
            if (debugVr != null && sender != null && sender is BrailleIOViewRange && e != null && activeScreen != null)
            {
                if (e.PropertyName.Equals("Zoom"))
                {
                    debugVr.SetText(getZoomPercentageBasedOnPrintZoom(activeScreen.Name, WindowManager.VR_CENTER_NAME).ToString() + "%");
                }
            }
        }
        #endregion

        #endregion

    }
}
