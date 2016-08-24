using System;
using System.Drawing;
using BrailleIO;
using BrailleIO.Interface;
using TangramLector.OO;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector.OO;
using BrailleIO.Renderer;
using BrailleIO.Renderer.Structs;
using System.Collections.Generic;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.TangramLector
{
    public partial class WindowManager : tud.mci.tangram.TangramLector.IZoomProvider
    {

        IBrailleIOAdapter getActiveAdapter()
        {
            if (io != null && io.AdapterManager != null)
            {
                return io.AdapterManager.ActiveAdapter;
            }
            return null;
        }

        #region Zooming

        public static float GetPrintZoomLevel(tud.mci.tangram.Accessibility.OoAccessibleDocWnd wnd = null)
        {
            if (wnd == null && OoConnector.Instance != null && OoConnector.Instance.Observer != null)
            {
                wnd = OoConnector.Instance.Observer.GetActiveDocument();
            }
            if (wnd != null)
            {
                var pagesObs = wnd.DrawPagesObs;
                if (pagesObs != null)
                {
                    var zoomLevel = (float)pagesObs.ZoomValue / 100f;
                    return _PRINT_ZOOM_FACTOR / zoomLevel;
                }
            }
            return _PRINT_ZOOM_FACTOR;
        }


        /// <summary>
        /// Get zoom level in percentage. Print zoom is defined as 100%.
        /// </summary>
        /// <param name="viewName">name of the current BrailleIOScreen</param>
        /// <param name="viewRangeName">name of the current view range</param>
        /// <returns>current zoom level in % or Double.NaN if there is no zoom level</returns>
        public double GetZoomPercentageBasedOnPrintZoom(string viewName, string viewRangeName)
        {
            if (io != null && io.GetView(viewName) as BrailleIOScreen != null)
            {
                BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
                if (vr != null)
                {
                    double zoom = vr.GetZoom();
                    double zoomPercentage = zoom / GetPrintZoomLevel();
                    zoomPercentage = (int)(zoomPercentage * 100);
                    zoomPercentage = Math.Round(zoomPercentage / 10, MidpointRounding.AwayFromZero) * 10; // round to 10er values
                    return zoomPercentage;
                }
            }
            return Double.NaN;
        }

        /// <summary>
        /// Zooms the given view range to the new zoom level and calculates the new offset out of the old center position.
        /// </summary>
        /// <param name="vr">view range that should be zoomed</param>
        /// <param name="oldZoom">old zoom of the view range</param>
        /// <param name="newZoom">new zoom</param>
        /// <returns>true, if zoom was successful</returns>
        private bool zoom(BrailleIOViewRange vr, double oldZoom, double newZoom)
        {
            if (vr != null)
            {
                Point oldCenter = new Point();
                var oldvrdin = vr.ContentBox;
                bool zoomToShape = false;
                Point oldOffset = new Point(vr.GetXOffset(), vr.GetYOffset());

                // Prüfung auf größte Zoomstufe
                if (newZoom > vr.MAX_ZOOM_LEVEL)
                {
                    if (oldZoom == vr.MAX_ZOOM_LEVEL) return false;
                    newZoom = vr.MAX_ZOOM_LEVEL;
                }
                // Prüfung auf kleinste Zoomstufe
                if (vr.ContentBox.Height >= vr.ContentHeight && vr.ContentBox.Width >= vr.ContentWidth)
                {
                    if (oldZoom >= newZoom) return false;
                    
                    oldCenter = Point.Empty;
                    // central point of focused element as center for zooming
                    if (OoConnector.Instance != null && OoConnector.Instance.Observer != null)
                    {
                        OoShapeObserver shape = OoConnector.Instance.Observer.GetLastSelectedShape();
                        if (shape != null)
                        {
                            Rectangle shapeBoundingbox = shape.GetRelativeScreenBoundsByDom();
                            if (shapeBoundingbox != null)
                            {
                                // calculate shape position and size in pins (relative to document boundings on pin device)
                                Point shapePosition = new Point((int)Math.Round(shapeBoundingbox.X * vr.GetZoom()), (int)Math.Round(shapeBoundingbox.Y * vr.GetZoom()));
                                Size shapeSize = new Size((int)Math.Round(shapeBoundingbox.Width * vr.GetZoom()), (int)Math.Round(shapeBoundingbox.Height * vr.GetZoom()));
                                Point shapeCenter = new Point(shapePosition.X + shapeSize.Width / 2, shapePosition.Y + shapeSize.Height / 2);
                                oldCenter = new Point(shapeCenter.X, shapeCenter.Y);
                            }
                        }
                    }

                    if (oldCenter != Point.Empty) zoomToShape = true;
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
                    oldCenter = Point.Empty;
                    // central point of focused element as center for zooming
                    if (OoConnector.Instance != null && OoConnector.Instance.Observer != null)
                    {
                        OoShapeObserver shape = OoConnector.Instance.Observer.GetLastSelectedShape();
                        if (shape != null)
                        {
                            Rectangle shapeBoundingbox = shape.GetRelativeScreenBoundsByDom();
                            if (shapeBoundingbox != null)
                            {
                                // calculate shape position and size in pins (relative to document boundings on pin device)
                                Point shapePosition = new Point((int)Math.Round(shapeBoundingbox.X * vr.GetZoom()), (int)Math.Round(shapeBoundingbox.Y * vr.GetZoom()));
                                Size shapeSize = new Size((int)Math.Round(shapeBoundingbox.Width * vr.GetZoom()), (int)Math.Round(shapeBoundingbox.Height * vr.GetZoom()));
                                Point shapeCenter = new Point(shapePosition.X + shapeSize.Width / 2, shapePosition.Y + shapeSize.Height / 2);
                                oldCenter = new Point(shapeCenter.X, shapeCenter.Y);
                            }
                        }
                    }

                    if (oldCenter != Point.Empty) zoomToShape = true;
                    else
                    {
                        // central point of center region as center for zooming
                        oldCenter = new Point(
                            (int)Math.Round(((double)oldvrdin.Width / 2) + (vr.GetXOffset() * -1)),
                            (int)Math.Round(((double)oldvrdin.Height / 2) + (vr.GetYOffset() * -1))
                            );
                    }
                }

                if (newZoom > 0)
                {
                    double zoomRatio = newZoom / oldZoom;
                    Point newCenter = new Point(
                        (int)Math.Round(oldCenter.X * zoomRatio),
                        (int)Math.Round(oldCenter.Y * zoomRatio)
                        );

                    Point newOffset = new Point();
                    if (zoomToShape)
                    {
                        newOffset = new Point(oldOffset.X + (oldCenter.X - newCenter.X), oldOffset.Y + (oldCenter.Y - newCenter.Y));
                    }
                    else
                    {
                        newOffset = new Point(
                            (int)Math.Round((newCenter.X - ((double)oldvrdin.Width / 2)) * -1),
                            (int)Math.Round((newCenter.Y - ((double)oldvrdin.Height / 2)) * -1)
                            );
                    }

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
        /// Zooms the specified view with the given factor.
        /// </summary>
        /// <param name="viewName">Name of the view.</param>
        /// <param name="viewRangeName">Name of the view range.</param>
        /// <param name="factor">The zoom factor.</param>
        protected bool zoomWithFactor(string viewName, string viewRangeName, double factor)
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
                    success = zoom(vr, oldZoom, newZoom);
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
        public bool ZoomTo(string viewName, string viewRangeName, double zoomLevel)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return false;

            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);

            bool success = false;
            if (vr != null)
            {
                var oldZoom = vr.GetZoom();
                var newZoom = zoomLevel;
                success = zoom(vr, oldZoom, newZoom);
            }
            else return false;

            io.RefreshDisplay();
            return success;
        }

        #endregion

        #region Panning

        /// <summary>
        /// Move the content of the given view range in horizontal direction.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="steps">count of pins to move</param>
        protected bool moveHorizontal(string viewName, string viewRangeName, int steps)
        {
            if (io == null && io.GetView(viewName) as BrailleIOScreen != null) return false;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            if (GetVisibleScreen().Name.Equals(BS_MINIMAP_NAME))
            {
                vr = screenBeforeMinimap.GetViewRange(viewRangeName); // offset has to be the offset of mainscreen view range instead of minimap view range, because minimap offset is always 0/0
            }
            if (vr != null)
            {
                if ((steps > 0 && vr.GetXOffset() >= 0) || (new Point(vr.GetXOffset(), vr.GetYOffset()).Equals(vr.MoveHorizontal(steps))))
                {
                    audioRenderer.PlayWaveImmediately(StandardSounds.End);
                    return false;
                }
            }
            io.RefreshDisplay(true);
            return true;
        }

        /// <summary>
        /// Move the content of the given view range in vertical direction.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="steps">count of pins to move</param>
        protected bool moveVertical(string viewName, string viewRangeName, int steps)
        {
            if (io == null && io.GetView(viewName) as BrailleIOScreen != null) return false;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            if (GetVisibleScreen().Name.Equals(BS_MINIMAP_NAME))
            {
                vr = screenBeforeMinimap.GetViewRange(viewRangeName); // offset has to be the offset of mainscreen view range instead of minimap view range, because minimap offset is always 0/0
            }
            if (vr != null)
            {
                if ((steps > 0 && vr.GetYOffset() >= 0) || (new Point(vr.GetXOffset(), vr.GetYOffset()).Equals(vr.MoveVertical(steps))))
                {
                    audioRenderer.PlayWaveImmediately(StandardSounds.End);
                    return false;
                }
            }
            io.RefreshDisplay(true);
            return true;
        }

        /// <summary>
        /// Move the given view range to the given offset position.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="offset">new offset position</param>
        /// <returns>true, if view range was moved successfully</returns>
        public bool MoveToPosition(string viewName, string viewRangeName, Point offset)
        {
            if (io == null && io.GetView(viewName) as BrailleIOScreen != null) return false;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            if (vr == null) return false;
            vr.MoveTo(offset);
            return true;
        }

        /// <summary>
        /// Move current view range to the given object. 
        /// For example, this should be used for following the focused object on the pin device.
        /// </summary>
        /// <param name="objectBounds">object that should be shown centered on pin device</param>
        public void MoveToObject(Rectangle boundingBox)
        {
            if (boundingBox == null || IsInMinimapMode()) return;

            BrailleIOScreen vs = GetVisibleScreen();
            BrailleIOViewRange center = GetActiveCenterView(vs);

            if (center != null)
            {
                // use not whole screen, but open office document position as screen position, because contentWidth and contentHeight of view range is related to this instead of to the whole screen
                //Rectangle boundingBox = new Rectangle(objectBounds.X, objectBounds.Y, objectBounds.Width, objectBounds.Height);
                Point centeredScreenPosition = new Point((int)Math.Round(boundingBox.X + boundingBox.Width * 0.5), (int)Math.Round(boundingBox.Y + boundingBox.Height * 0.5));
                //Point centeredScreenPosition = new Point(boundingBox.X, boundingBox.Y);

                Point pinDevicePosition = new Point(0, 0);
                double zoom = center.GetZoom();
                if (zoom != 0)
                {
                    pinDevicePosition = new Point((int)(Math.Round(centeredScreenPosition.X * -zoom) + ((int)Math.Round(center.ContentBox.Width * 0.5))), (int)(Math.Round(centeredScreenPosition.Y * -zoom) + ((int)Math.Round(center.ContentBox.Height * 0.5))));
                    //pinDevicePosition = new Point((int)(Math.Round(centeredScreenPosition.X * -zoom)), (int)(Math.Round(centeredScreenPosition.Y * -zoom)));
                }
                MoveToPosition(vs.Name, center.Name, pinDevicePosition);
            }
        }

        private void jumpToBottomBorder(BrailleIOViewRange center)
        {
            if (center != null)
            {
                int offset = center.ContentHeight - center.ContentBox.Height;
                center.SetYOffset(-offset);
                io.RefreshDisplay();
                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.jump.down"));
                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] jump to bottom");
            }
        }

        private void jumpToRightBorder(BrailleIOViewRange center)
        {
            if (center != null)
            {
                int offset = center.ContentWidth - center.ContentBox.Width;
                center.SetXOffset(-offset);
                io.RefreshDisplay();
                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.jump.right"));
                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] jump to right");
            }
        }

        private void jumpToLeftBorder(BrailleIOViewRange center)
        {
            if (center != null)
            {
                center.SetXOffset(0);
                io.RefreshDisplay();
                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.jump.left"));
                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] jump to left");
            }
        }

        private void jumpToTopBorder(BrailleIOViewRange center)
        {
            if (center != null)
            {
                center.SetYOffset(0);
                io.RefreshDisplay();
                audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.panning.jump.up"));
                Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] jump to top");
            }
        }

        /// <summary>
        /// Add the given scroll step to the current offset position of view range to realize a scrolling.
        /// </summary>
        /// <param name="vr">View range to scroll.</param>
        /// <param name="scrollStep">Steps to scroll the view range.</param>
        private bool scrollViewRange(BrailleIOViewRange vr, int scrollStep)
        {
            if (vr != null)
            {
                int newPos = vr.OffsetPosition.Y + scrollStep;
                if (newPos <= 0 && newPos >= (vr.ContentBox.Height - vr.ContentHeight))
                {
                    vr.SetYOffset(newPos);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Set offset position of view range to given position.
        /// </summary>
        /// <param name="vr">View range to set new offset.</param>
        /// <param name="y_pos">new offset position.</param>
        private void scrollViewRangeTo(BrailleIOViewRange vr, int y_pos)
        {
            if (vr != null && y_pos <= 0)
            {
                if (y_pos >= (vr.ContentBox.Height - vr.ContentHeight)) vr.SetYOffset(y_pos);
                else vr.SetYOffset(0);
            }
        }

        #endregion

        #region toggle screen, view ranges and switch mode methods

        // TODO: correct toggle function for active detail area view range
        private void toggleFullDetailScreen()
        {
            BrailleIOScreen vs = GetVisibleScreen();
            BrailleIOViewRange detailVR = vs.GetViewRange(VR_DETAIL_NAME); // TODO: ImageData class creates own detail area --> do not use global detailarea
            BrailleIOViewRange topVR = vs.GetViewRange(VR_TOP_NAME);
            detailVR.SetHeight(deviceSize.Height - topVR.GetHeight());
            detailVR.SetTop(topVR.GetHeight() - 3);
            io.RefreshDisplay();
            audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.wm.views.detail_maximize"));
        }

        /// <summary>
        /// Starts Fullscreen mode.
        /// </summary>
        public bool StartFullscreen()
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (vs != null && !vs.Name.Equals(WindowManager.BS_FULLSCREEN_NAME))
            {
                if (!fullscreenVisible())
                {
                    toggleFullscreen();
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Quit Fullscreen mode.
        /// </summary>
        public bool QuitFullscreen()
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (vs != null && vs.Name.Equals(WindowManager.BS_FULLSCREEN_NAME))
            {
                if (fullscreenVisible())
                {
                    toggleFullscreen();
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Ask if fullscreen is visible.
        /// </summary>
        /// <returns>true if it is visible</returns>
        private bool fullscreenVisible()
        {
            var view = io.GetView(BS_FULLSCREEN_NAME) as BrailleIO.Interface.IViewable;
            if (view != null) return view.IsVisible();
            return false;
        }

        /// <summary>
        /// Switch between fullscreen and main screen.
        /// </summary>
        private void toggleFullscreen()
        {
            if (io.GetView(BS_FULLSCREEN_NAME) != null)
            {
                var view = io.GetView(BS_FULLSCREEN_NAME) as BrailleIO.Interface.IViewable;
                BrailleIOScreen ms = io.GetView(BS_MAIN_NAME) as BrailleIOScreen;
                if (view != null && ms != null)
                {
                    BrailleIOViewRange mainscreenCenterVR = ms.GetViewRange(VR_CENTER_NAME);
                    BrailleIOViewRange fullscreenCenterVR = (view as BrailleIOScreen).GetViewRange(VR_CENTER_NAME);

                    if (view.IsVisible())
                    {
                        if (mainscreenCenterVR != null && fullscreenCenterVR != null)
                        {
                            syncViewRangeWithVisibleCenterViewRange(mainscreenCenterVR);
                            syncContrastSettings(fullscreenCenterVR, mainscreenCenterVR);
                        }
                        io.HideView(BS_FULLSCREEN_NAME);
                        io.ShowView(BS_MAIN_NAME);
                    }
                    else
                    {
                        if (mainscreenCenterVR != null && fullscreenCenterVR != null)
                        {
                            syncViewRangeWithVisibleCenterViewRange(fullscreenCenterVR);
                            syncContrastSettings(mainscreenCenterVR, fullscreenCenterVR);
                        }
                        io.HideView(BS_MAIN_NAME);
                        io.ShowView(BS_FULLSCREEN_NAME);
                    }
                }
            }
        }

        /// <summary>
        /// Switch between minimap and currently visible screen.
        /// </summary>
        private void toggleMinimap()
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (vs != null)
            {
                if (vs.Name.Equals(BS_MINIMAP_NAME)) // exit minimap mode
                {
                    BrailleIOScreen ms = screenBeforeMinimap;
                    if (ms != null)
                    {
                        BrailleIOViewRange ocvr = vs.GetViewRange(VR_CENTER_NAME);
                        BrailleIOViewRange ncvr = ms.GetViewRange(VR_CENTER_NAME);
                        if (ocvr != null && ncvr != null)
                        {
                            syncContrastSettings(ocvr, ncvr);
                        }
                    }
                    io.HideView(vs.Name);
                    io.ShowView(ms.Name);
                    SetDetailRegionContent(GetLastDetailRegionContent());
                }
                else // show minimap
                {
                    screenBeforeMinimap = vs;

                    BrailleIOScreen ms = io.GetView(BS_MINIMAP_NAME) as BrailleIOScreen;
                    if (ms != null)
                    {
                        BrailleIOViewRange ocvr = vs.GetViewRange(VR_CENTER_NAME);
                        BrailleIOViewRange ncvr = ms.GetViewRange(VR_CENTER_NAME);
                        if (ocvr != null && ncvr != null)
                        {
                            syncContrastSettings(ocvr, ncvr);
                        }
                    }
                    io.HideView(vs.Name);
                    io.ShowView(BS_MINIMAP_NAME);
                    BrailleIOViewRange vr = vs.GetViewRange(VR_CENTER_NAME);
                    if (vr != null)
                    {
                        SetDetailRegionContent(LL.GetTrans("tangram.lector.wm.zooming.current.short", ((int)(vr.GetZoom() * 100)).ToString()));
                    }
                }
            }
        }

        /// <summary>
        /// Determines whether the minimap mode is active.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if minimap mode is active; otherwise, <c>false</c>.
        /// </returns>
        public bool IsInMinimapMode()
        {
            BrailleIOScreen vs = GetVisibleScreen();
            if (vs != null && vs.Name.Equals(BS_MINIMAP_NAME))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Change view and interaction mode. 
        /// </summary>
        /// <param name="view">new mode</param>
        public bool ChangeLectorView(LectorView view)
        {
            if (!view.Equals(currentView))
            {
                currentView = view;
                if (view.Equals(LectorView.Braille))
                {
                    InteractionManager.ChangeMode(InteractionMode.Braille);
                }
                else if (view.Equals(LectorView.Drawing))
                {
                    InteractionManager.ChangeMode(InteractionMode.Normal);

                }
                updateStatusRegionContent();
                return true;
            }
            else return false;
        }

        /// <summary>
        /// Get the active center view of the given screen
        /// </summary>
        /// <param name="screen">screen</param>
        /// <returns>active center view</returns>
        public BrailleIOViewRange GetActiveCenterView(BrailleIOScreen screen)
        {
            if (screen != null)
            {
                BrailleIOViewRange center = screen.GetViewRange(VR_CENTER_NAME);
                BrailleIOViewRange center2 = screen.GetViewRange(VR_CENTER_2_NAME);
                if (center != null && center.IsVisible()) return center;
                else if (center2 != null && center2.IsVisible()) return center2;
            }
            return null;
        }

        /// <summary>
        /// Changes the active center view to the given one (center or center_2).
        /// </summary>
        /// <param name="screen">The screen.</param>
        /// <param name="viewName">Name of the view. [VR_CENTER_NAME or VR_CENTER_2_NAME]</param>
        /// <returns></returns>
        BrailleIOViewRange changeActiveCenterView(BrailleIOScreen screen, String viewName)
        {
            //TODO: do this generic?!
            if (screen != null)
            {
                BrailleIOViewRange center = screen.GetViewRange(VR_CENTER_NAME);
                BrailleIOViewRange center2 = screen.GetViewRange(VR_CENTER_2_NAME);

                if (viewName.Equals(VR_CENTER_2_NAME))
                {
                    if (center2 != null)
                    {
                        if (center != null) { center.SetVisibility(false); }
                        center2.SetVisibility(true);
                        return center2;
                    }
                }
                else
                {
                    if (center != null) { center.SetVisibility(true); }
                    if (center2 != null) { center2.SetVisibility(false); }
                    return center;
                }
            }
            return null;
        }

        /// <summary>
        /// Change visibility of the given view range. Center view range cannot be hidden.
        /// </summary>
        /// <param name="viewName">name of the screen</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="visible">true if view range should become visible, false if view range should become invisible</param>
        /// <returns>true if successful, false if not successful</returns>
        private bool changeViewVisibility(string viewName, string viewRangeName, bool visible)
        {
            if (!viewRangeName.Equals(VR_CENTER_NAME))
            {
                BrailleIOScreen vs = io.GetView(viewName) as BrailleIOScreen;
                if (vs != null)
                {
                    BrailleIOViewRange vr = vs.GetViewRange(viewRangeName);
                    BrailleIOViewRange center = vs.GetViewRange(VR_CENTER_NAME);
                    BrailleIOViewRange top = vs.GetViewRange(VR_TOP_NAME);
                    BrailleIOViewRange bottom = vs.GetViewRange(VR_DETAIL_NAME);

                    if (vr != null && center != null && top != null && bottom != null)
                    {
                        vr.SetVisibility(visible);
                        // change margin of center region
                        if (visible)
                        {
                            if (viewRangeName.Equals(VR_TOP_NAME))
                            {
                                if (bottom.IsVisible()) center.SetMargin(7, 0);
                                else center.SetMargin(7, 0, 0);
                            }
                            else if (viewRangeName.Equals(VR_DETAIL_NAME))
                            {
                                if (top.IsVisible()) center.SetMargin(7, 0);
                                else center.SetMargin(0, 0, 7);
                            }
                        }
                        else
                        {
                            if (viewRangeName.Equals(VR_TOP_NAME))
                            {
                                if (bottom.IsVisible()) center.SetMargin(0, 0, 7);
                                else center.SetMargin(0, 0);
                            }
                            else if (viewRangeName.Equals(VR_DETAIL_NAME))
                            {
                                if (top.IsVisible()) center.SetMargin(7, 0, 0);
                                else center.SetMargin(0, 0);
                            }
                        }
                        return true;
                    }
                }
            }
            audioRenderer.PlayWaveImmediately(StandardSounds.Error); // operation not possible
            return false;
        }

        #endregion

        #region graphic methods

        /// <summary>
        /// Invert the image.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        private void invertImage(string viewName, string viewRangeName)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            if (vr != null)
            {
                vr.InvertImage = !vr.InvertImage;
            }
            io.RefreshDisplay();
        }

        /// <summary>
        /// Reduce or increase contrast of the image with the given value.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="factor">factor that should be added to current contrast</param>
        /// <returns>true if threshold could be set to new value, false if it is already max or min value</returns>
        private bool updateContrast(string viewName, string viewRangeName, int factor)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return false;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            int oldThreshold = 0;
            int newThreshold = 0;
            if (vr != null)
            {
                oldThreshold = vr.GetContrastThreshold();
                newThreshold = vr.SetContrastThreshold(oldThreshold + factor);
            }
            io.RefreshDisplay();
            if ((oldThreshold >= 255 || oldThreshold <= 0) && (newThreshold >= 255 || newThreshold <= 0)) return false;
            else return true;
        }

        /// <summary>
        /// Change contrast of the image to the given value.
        /// </summary>
        /// <param name="viewName">name of the view</param>
        /// <param name="viewRangeName">name of the view range</param>
        /// <param name="contrast">new contrast value (should be between 0 and 255)</param>
        private void setContrast(string viewName, string viewRangeName, int contrast)
        {
            if (io == null || io.GetView(viewName) as BrailleIOScreen == null) return;
            BrailleIOViewRange vr = ((BrailleIOScreen)io.GetView(viewName)).GetViewRange(viewRangeName);
            if (vr != null)
            {
                vr.SetContrastThreshold(contrast);
            }
            io.RefreshDisplay(true);
        }

        #endregion

        #region touched elements handling

        /// <summary>
        /// Reads a selected oo acc item aloud.
        /// </summary>
        /// <param name="observed">The observed Open Office window.</param>
        /// <param name="p">The point to check.</param>
        private void handleSelectedOoAccItem(Accessibility.OoAccessibleDocWnd observed, Point p, GestureEventArgs eventData = null)
        {
            if (observed != null && observed.DocumentComponent != null)
            {
                //try to get the touched element
                tud.mci.tangram.Accessibility.OoAccComponent c = observed.DocumentComponent.GetAccessibleFromScreenPos(p);

                if (eventData == null || ( // selection for reading
                    eventData.PressedGenericKeys.Count < 1 && eventData.ReleasedGenericKeys.Count == 1
                    && eventData.ReleasedGeneralKeys.Contains(BrailleIO_DeviceButton.Gesture)
                    ))
                {
                    // try to get a shape observer for better audio output
                    if (c != null && observed.DrawPagesObs != null)
                    {
                        OoShapeObserver sObs = observed.DrawPagesObs.GetRegisteredShapeObserver(c);
                        String text = "";
                        if (sObs != null)
                        {
                            text = OoElementSpeaker.PlayElementTitleAndDescriptionImmediately(sObs);
                        }
                        else
                        {
                            text = OoElementSpeaker.PlayElementImmediately(c);
                        }
                        SetDetailRegionContent(text);
                    }
                }
                else if (eventData != null && eventData.ReleasedGenericKeys.Contains("hbr"))
                {
                    OoShapeObserver shape = OoConnector.Instance.Observer.GetShapeForModification(c, observed);
                    //TODO: what to do if shape is null?
                }
            }
        }

        /// <summary>
        /// If Braille text is at the requested position in the given view range that this function tries to get the corresponding text.
        /// </summary>
        /// <param name="vr">The viewrange to check.</param>
        /// <param name="e">the gestrure evnet</param>
        private String CheckForTouchedBrailleText(BrailleIOViewRange vr, GestureEventArgs e)
        {
            if (vr.ContentRender is ITouchableRenderer)
            {
                var touchRendeer = (ITouchableRenderer)vr.ContentRender;

                // calculate the position in the content
                int x = (int)e.Gesture.NodeParameters[0].X;
                int y = (int)e.Gesture.NodeParameters[0].Y;

                x = x - vr.ViewBox.X - vr.ContentBox.X - vr.GetXOffset();
                y = y - vr.ViewBox.Y - vr.ContentBox.Y - vr.GetYOffset();

                var touchedElement = touchRendeer.GetContentAtPosition(x, y);
                if (touchedElement != null && touchedElement is RenderElement)
                {
                    var touchedValue = ((RenderElement)touchedElement).GetValue();
                    if (((RenderElement)touchedElement).HasSubParts())
                    {
                        List<RenderElement> touchedSubparts = ((RenderElement)touchedElement).GetSubPartsAtPoint(x, y);
                        if (touchedSubparts != null && touchedSubparts.Count > 0)
                        {
                            touchedValue = touchedSubparts[0].GetValue();
                        }
                    }

                    System.Diagnostics.Debug.WriteLine("----- [BRAILLE TEXT TOUCHED] : '" + touchedValue.ToString() + "'");
                    return touchedValue.ToString();
                }
            }
            return String.Empty;
        }

        #endregion

        #region Audio

        /// <summary>
        /// Plays a specified text string.
        /// </summary>
        /// <param name="text">The text to play as text to speach.</param>
        private static void play(String text) { AudioRenderer.Instance.PlaySoundImmediately(text); }

        /// <summary>
        /// Aborts all audio output.
        /// </summary>
        public void AbortAudio()
        {
            AudioRenderer.Instance.Abort();
            ShowTemporaryMessageInDetailRegion(LL.GetTrans("tangram.lector.wm.audio.stop"));
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION] abort speech output");
        }

        private void sayInteractionModeChange(FollowFocusModes oldValue, FollowFocusModes newValue)
        {
            switch (FocusMode)
            {
                case FollowFocusModes.NONE:
                    if (oldValue == FollowFocusModes.FOLLOW_BRAILLE_FOCUS)
                    {
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION MODE] end Braille focus following mode");

                        string msg = LL.GetTrans("tangram.lector.wm.modes.end", LL.GetTrans("tangram.lector.wm.modes.FOLLOW_BRAILLE_FOCUS"));
                        play(msg);
                        ShowTemporaryMessageInDetailRegion(msg);
                    }
                    else if (oldValue == FollowFocusModes.FOLLOW_MOUSE_FOCUS)
                    {
                        Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION MODE] end Mouse focus following mode");
                        string msg = LL.GetTrans("tangram.lector.wm.modes.end", LL.GetTrans("tangram.lector.wm.modes.FOLLOW_MOUSE_FOCUS"));
                        play(msg);
                        ShowTemporaryMessageInDetailRegion(msg);
                    }
                    break;
                case FollowFocusModes.FOLLOW_BRAILLE_FOCUS:
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION MODE] start Braille focus following mode");
                    string msg2 = LL.GetTrans("tangram.lector.wm.modes.start", LL.GetTrans("tangram.lector.wm.modes.FOLLOW_BRAILLE_FOCUS"));
                    play(msg2);
                    ShowTemporaryMessageInDetailRegion(msg2);
                    break;
                case FollowFocusModes.FOLLOW_MOUSE_FOCUS:
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[INTERACTION MODE] start Mouse focus following mode");
                    string msg3 = LL.GetTrans("tangram.lector.wm.modes.start", LL.GetTrans("tangram.lector.wm.modes.FOLLOW_MOUSE_FOCUS"));
                    play(msg3);
                    ShowTemporaryMessageInDetailRegion(msg3);
                    break;
                default:
                    break;
            }
        }

        #endregion

        #region IZoomable

        /// <summary>
        /// get the current zoom level
        /// </summary>
        /// <returns></returns>
        double IZoomProvider.GetZoomLevel()
        {
            double z = 1;
            var screen = WindowManager.Instance.GetVisibleScreen();
            if (screen != null && screen.HasViewRange(WindowManager.VR_CENTER_NAME))
            {
                var vr = screen.GetViewRange(WindowManager.VR_CENTER_NAME);
                if (vr != null)
                {
                    return vr.GetZoom();
                }
            }
            return z;
        }

        /// <summary>
        /// set the current zoom level
        /// </summary>
        /// <param name="zoom"></param>
        /// <returns></returns>
        double IZoomProvider.SetZoomLevel(double zoom)
        {
            double z = 1;

            var screen = WindowManager.Instance.GetVisibleScreen();
            if (screen != null && screen.HasViewRange(WindowManager.VR_CENTER_NAME))
            {
                ZoomTo(screen.Name, WindowManager.VR_CENTER_NAME, zoom);

                var vr = screen.GetViewRange(WindowManager.VR_CENTER_NAME);
                if (vr != null)
                {
                    return vr.GetZoom();
                }
            }
            return z;
        }

        /// <summary>
        /// get the maximum allows zoom level
        /// </summary>
        /// <returns></returns>
        double IZoomProvider.GetMaxZoomLevel()
        {
            //TODO: implement
            return 0;
        }

        #endregion

    }
}