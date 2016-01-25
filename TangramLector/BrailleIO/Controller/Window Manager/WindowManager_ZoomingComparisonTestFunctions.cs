using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BrailleIO;
using System.Drawing;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.controller.observer;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager
    {
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
            if (ooc != null && ooc.Observer != null)
            {
                OoShapeObserver shape = ooc.Observer.GetLastSelectedShape();
                if (shape != null)
                {
                    BrailleIOScreen vs = GetVisibleScreen();
                    if (vs != null)
                    {
                        BrailleIOViewRange center = GetActiveCenterView(vs);
                        if (center != null)
                        {
                            Rectangle boundings = shape.GetRelativeScreenBoundsByDom();
                            double zoom = center.GetZoom();
                            Point offset = calculateRandomOffset(boundings, zoom);

                            if (!(center.ContentBox.Height >= center.ContentHeight && center.ContentBox.Width >= center.ContentWidth))
                            {
                                MoveToPosition(vs.Name, center.Name, offset);
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
            if (vs != null && vr != null) Logger.Instance.Log(LogPriority.MIDDLE, "[ZOOMING TEST]", "zoom percent level = " + GetZoomPercentageBasedOnPrintZoom(vs.Name, vr.Name).ToString() + "%");
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
            if (ooc != null && ooc.Observer != null)
            {
                OoShapeObserver shape = ooc.Observer.GetLastSelectedShape();
                if (shape != null)
                {
                    BrailleIOScreen vs = GetVisibleScreen();
                    if (vs != null)
                    {
                        BrailleIOViewRange center = GetActiveCenterView(vs);
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
    }
}
