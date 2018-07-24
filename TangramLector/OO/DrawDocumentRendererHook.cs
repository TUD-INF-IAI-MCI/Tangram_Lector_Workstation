using System;
using System.Drawing;
using BrailleIO;
using BrailleIO.Interface;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector;
using tud.mci.tangram.util;

namespace tud.mci.tangram.TangramLector.OO
{
    /// <summary>
    /// Renderer hook, drawing a dotted frame indicating the draw page boundaries.
    /// </summary>
    public class DrawDocumentRendererHook : IBailleIORendererHook
    {
        #region Members

        #region public Members

        readonly object lockObj = new Object();
        private OoAccessibleDocWnd __wnd;
        public OoAccessibleDocWnd Wnd
        {
            get
            {
                lock (lockObj)
                {
                    return __wnd;
                }
            }

            set
            {
                lock (lockObj)
                {
                    __wnd = value;
                    ask = true;
                }
            }

        }

        /// <summary>
        /// Sets the hook active or not
        /// </summary>
        public volatile bool Active = false;

        #endregion

        #region private Members

        private volatile bool ask = true;
        private Rectangle pos = new Rectangle();

        DateTime last = DateTime.Now;
        TimeSpan refresh = new TimeSpan(0, 0, 0, 0, 250);

        #endregion

        #endregion

        /// <summary>
        /// Forces the hook to ask for the screen position of the given window.
        /// </summary>
        public void Update() { ask = true; }

        #region IBailleIORendererHook implementation

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer after he has done his rendering before returning the result.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="result">The result matrix, may be manipulated. Addressed in [y, x] notation.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        void IBailleIORendererHook.PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            if (Active && Wnd != null)
            {
                if (Wnd.Disposed)
                {
                    Wnd = null;
                    //TODO: maybe set this as not active
                    return;
                }
                if (!ask) // refresh regularly
                {
                    if ((DateTime.Now - last) > refresh) ask = true;
                }

                // check if the bound have to be updated or not
                if (ask || pos.Width < 1 || pos.Height < 1)
                {
                    var activePage = Wnd.GetActivePage();
                    pos = GetPageBoundsAndPosition(activePage);
                    ask = false;
                    last = DateTime.Now;
                }

                Rectangle zPos = GetPageBoundsInViewRange(view, pos);

                //TODO: decide in inner frame or outer frame
                int y1 = zPos.Y - 1;
                int y2 = zPos.Bottom;
                int x1 = 0;


                int width = result.GetLength(1);
                int height = result.GetLength(0);

                //horizontal lines
                for (int x = 0; x < zPos.Width; x += 2)
                {
                    x1 = zPos.X + x;

                    if (x1 >= 0 && x1 < width)
                    {
                        if (y1 >= 0 && y1 < height)
                        {
                            result[y1, x1] = true;
                        }

                        if (y2 >= 0 && y2 < height)
                        {
                            result[y2, x1] = true;
                        }
                    }
                }

                x1 = zPos.X - 1;
                int x2 = zPos.X + zPos.Width;

                //vertical lines
                for (int y = 0; y < zPos.Height; y += 2)
                {
                    y1 = zPos.Y + y;
                    if (y1 >= 0 && y1 < height)
                    {

                        if (x1 >= 0 && x1 < width)
                        {
                            result[y1, x1] = true;
                        }

                        if (x2 >= 0 && x2 < width)
                        {
                            result[y1, x2] = true;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer before he starts his rendering.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        void IBailleIORendererHook.PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams) { }

        #endregion

        #region static helper functions for page bounds

        /// <summary>
        /// Gets the page bounds and position on the screen in pixels.
        /// </summary>
        /// <param name="activePage">The active page to get the position on the screen of.</param>
        /// <returns>The bounds of the given page on the screen.</returns>
        public static Rectangle GetPageBoundsAndPosition(OoDrawPageObserver activePage)
        {
            Rectangle pos = new Rectangle();

            if (activePage != null)
            {
                System.Drawing.Rectangle pageBoundsInPx = OoDrawUtils.convertToPixel(
                    new System.Drawing.Rectangle(
                        -activePage.PagesObserver.ViewOffset.X + activePage.BorderLeft,
                        -activePage.PagesObserver.ViewOffset.Y + activePage.BorderTop,
                        activePage.Width - activePage.BorderLeft - activePage.BorderRight,
                        activePage.Height - activePage.BorderTop - activePage.BorderBottom),
                    activePage.PagesObserver.ZoomValue,
                    OoDrawPagesObserver.PixelPerMeterY,
                    OoDrawPagesObserver.PixelPerMeterY);

                pos = new Rectangle(pageBoundsInPx.X, pageBoundsInPx.Y,
                    pageBoundsInPx.Width, pageBoundsInPx.Height);
            }
            return pos;
        }

        /// <summary>
        /// Gets the page bounds in view range.
        /// </summary>
        /// <param name="view">The view the page is presented in.</param>
        /// <param name="pagePosOnScreen">The page bounds on the screen in pixels.</param>
        /// <returns>The page bounds in the content of the view range in pins.</returns>
        public static Rectangle GetPageBoundsInViewRange(IViewBoxModel view, Rectangle pagePosOnScreen)
        {

            // make the document bounds relative to the chosen zoom level
            double zoom = view is IZoomable ? zoom = ((IZoomable)view).GetZoom() : 1;
            if (((BrailleIOViewRange)view).Name.Equals(WindowManager.VR_CENTER_NAME) && ((BrailleIOViewRange)view).Parent.Name.Equals(WindowManager.BS_MINIMAP_NAME))
            {
                if (WindowManager.Instance != null)
                {
                    zoom = WindowManager.Instance.MinimapScalingFactor; // handling for minimap mode
                }
            }

            Rectangle zPos = new Rectangle(
                (int)(pagePosOnScreen.X * zoom),
                (int)(pagePosOnScreen.Y * zoom),
                (int)(pagePosOnScreen.Width * zoom),
                (int)(pagePosOnScreen.Height * zoom)
                );

            // add the panning offsets
            if (view is IPannable)
            {
                zPos.X += ((IPannable)view).GetXOffset();
                zPos.Y += ((IPannable)view).GetYOffset();
            }

            return zPos;
        }

        #endregion

    }
}
