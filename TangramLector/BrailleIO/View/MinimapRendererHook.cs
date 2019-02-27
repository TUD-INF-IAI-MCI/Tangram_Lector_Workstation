using BrailleIO;
using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tud.mci.tangram.TangramLector.BrailleIO.View
{
    /// <summary>
    /// Hook for the minimap Screen capturing renderer to present a blinking frame
    /// corresponding to the view port position and size of the previously presented Screen.
    /// </summary>
    /// <seealso cref="BrailleIO.Interface.IBailleIORendererHook" />
    class MinimapRendererHook : IBailleIORendererHook, IActivatable
    {
        #region Member
        BrailleIOMediator io { get { return BrailleIOMediator.Instance; } }
        WindowManager wm { get { return WindowManager.Instance; } }
        
        /// <summary>
        /// stroes the last visible view to handle before turning on the minimap (normal|full screen)
        /// </summary>
        BrailleIOScreen lastActiveView = null;
        
        /// <summary>
        /// Gets a value indicating whether the hook should be active or not.
        /// </summary>
        public bool Active { get; set; }
        
        /// <summary>
        /// Gets a value indicating whether [show raised frame].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [show raised frame]; otherwise, <c>false</c>.
        /// </value>
        public bool ShowRaisedFrame
        {
            get { if (wm != null) return wm.BlinkPinsUp; else return true; }
        }

        private int _borderWidth = 2;
        /// <summary>
        /// Gets or sets the width of the marking border.
        /// </summary>
        /// <value>
        /// The width of the border.
        /// </value>
        public int BorderWidth
        {
            get { return _borderWidth; }
            set
            {
                _borderWidth = value;
                borderOffset = value / 2;
            }
        }
        /// <summary>
        /// The pixel offset for the border which is the half of the border width.
        /// </summary>
        int borderOffset = 1;

        /// <summary>
        /// Scaling factor for resizing original content to its size in minimap mode.
        /// </summary>
        public double MinimapScalingFactor { get; private set; }
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="MinimapRendererHook"/> class.
        /// </summary>
        public MinimapRendererHook()
        {
            if (io != null)
            {
                io.VisibleViewsChanged += io_VisibleViewsChanged;
                BorderWidth = 2;
            }
        }

        #endregion
        
        #region IBailleIORendererHook

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer before he starts his rendering.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        public void PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams)
        {
            // make it always fit to height if default value was set
            if (view is IZoomable && ((IZoomable)view).GetZoom() == 1) 
                ((IZoomable)view).SetZoom(-1);
            return;
        }

        /// <summary>
        /// Posts the render hook.
        /// </summary>
        /// <param name="vr">The vr.</param>
        /// <param name="content">The content.</param>
        /// <param name="result">The result.</param>
        /// <param name="additionalParams">The additional parameters.</param>
        public void PostRenderHook(IViewBoxModel vr, object content, ref bool[,] result, params object[] additionalParams)
        {
            if (!Active || BorderWidth <= 0 || 
                result.GetLength(0) < 1 || result.GetLength(1) < 1) 
                return; // if no border or inactive, stop handling
            
            if (lastActiveView != null)
            {
                // center view of last active screen
                BrailleIOViewRange vr2mark = lastActiveView.GetViewRange(WindowManager.VR_CENTER_NAME);
                if (vr2mark == null) return;

                // currently available space
                double width = vr.ContentBox.Width;
                double height = vr.ContentBox.Height;

                // content height and view box position of the view port to mark
                double xoffset = vr2mark.GetXOffset();
                double yoffset = vr2mark.GetYOffset();
                double cntWidth = vr2mark.ContentWidth;
                double cntHeight = vr2mark.ContentHeight;
                double f2mWidth = vr2mark.ContentBox.Width;
                double f2mHeight = vr2mark.ContentBox.Height;

                if (cntWidth > 0 && cntHeight > 0)
                {
                    // scale factor between vr to mark and the current views zooming factor
                    var zoomScale = ((IZoomable)vr).GetZoom() / vr2mark.GetZoom();

                    // calculate position and size of the frame to mark
                    int x = (int)(Math.Abs(xoffset) * zoomScale - borderOffset); 
                    int y = (int)(Math.Abs(yoffset) * zoomScale - borderOffset);
                    int frameWidth = (int)(Math.Min(f2mWidth * zoomScale + BorderWidth * 2, width));
                    int frameHeigth = (int)(Math.Min(f2mHeight * zoomScale + BorderWidth * 2, height));

                    bool v = ShowRaisedFrame;
                    int rW = result.GetLength(1);
                    int rH = result.GetLength(0);

                    // paint horizontal lines
                    for (int m = 0; m < BorderWidth; m++)
                    {
                        int m1 = m + y;
                        int m2 = m1 + frameHeigth - BorderWidth;
                        for (int n = x; n < x + frameWidth; n++)
                        {
                            if(n > -1)
                            if (n < rW) {
                                // top line
                                if (m1 > -1 && m1 < rH) result[m1, n] = v;
                                // bottom line 
                                if (m2 >= -1 && m2 < rH) result[m2, n] = v;
                            }
                            else break;
                        }
                    }

                    // paint vertical lines
                    for (int n = 0; n < BorderWidth; n++)
                    {
                        int n1 = n + x;
                        int n2 = n1 + frameWidth - BorderWidth;
                        for (int m = y; m < y + frameHeigth; m++)
                        {
                            if(m > -1)
                            if (m < rH)
                            {
                                // left line
                                if (n1 > -1 && n1 < rW) result[m, n1] = v;
                                // right line 
                                if (n2 > -1 && n2 < rW) result[m, n2] = v;

                            }
                            else break;
                        }
                    }
                }
            }
        }

        #endregion

        #region Events

        void io_VisibleViewsChanged(object sender, VisibilityChangedEventArgs e)
        {
            if (e != null)
            {
                // set active or not
                if (e.VisibleViews != null && e.VisibleViews.Count > 0)
                {
                    foreach (var view in e.VisibleViews)
                    {
                        if (view is BrailleIOScreen)
                        {
                            if (view.Name.Equals(WindowManager.BS_MINIMAP_NAME))
                            {
                                Active = true;
                            }
                            else
                            {
                                Active = false;
                            }
                        }
                    }
                }
                else
                {
                    // deactivate
                    Active = false;
                }

                // store the previous screen to be handled in the minimap (normal|full screen)
                if (e.PreviouslyVisibleViews != null && e.PreviouslyVisibleViews.Count > 0)
                {
                    foreach (var item in e.PreviouslyVisibleViews)
                    {
                        if (item is BrailleIOScreen)
                        {
                            switch (item.Name)
                            {
                                case WindowManager.BS_FULLSCREEN_NAME:
                                case WindowManager.BS_MAIN_NAME:
                                    lastActiveView = item as BrailleIOScreen;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        #endregion
    }
}