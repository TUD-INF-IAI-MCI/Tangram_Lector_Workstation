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
    /// Renderer hook drawing a dotted frame indication the page where to draw.
    /// </summary>
    class DrawDocumentRendererHook : IBailleIORendererHook
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

        #endregion

        #endregion

        /// <summary>
        /// Forces the hook to ask for the screen position of the given window.
        /// </summary>
        public void Update() { ask = true; }

        #region IBailleIORendererHook implementation

        // FIXME: bad hack - needed because the OpenOffice api returns an inconsistent way for 
        // the position and size.
        // while the position is the top-left corner minus the padding to the sheet edges, the 
        // size is the size of the sheet including the padding.
        // This results in a wrong placed page - it is misplaced of the top- and left-padding. 
        // This padding have to be get from the page settings. But this would cost to much 
        // performance so estimate it - this estimation is based on the standard settings of
        // OpenOffice Draw DinA4 document.
        private const double pageXOffsetFactor = 0.035;
        private const double pageYOffsetFactor = 0.035;

        DateTime last = DateTime.Now;
        TimeSpan refresh = new TimeSpan(0, 0, 0, 0, 250);

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
                       var spos = new System.Drawing.Rectangle(pageBoundsInPx.X, pageBoundsInPx.Y, pageBoundsInPx.Width, pageBoundsInPx.Height);

                        if (spos.Width > 0 && spos.Height > 0)
                        {
                            pos = spos;
                        }
                        ask = false;
                        last = DateTime.Now;

                    }
                }

                // make the document bounds relative to the chosen zoom level
                double zoom = view is BrailleIO.Interface.IZoomable ? zoom = ((BrailleIO.Interface.IZoomable)view).GetZoom() : 1;
                if (((BrailleIOViewRange)view).Name.Equals(WindowManager.VR_CENTER_NAME) && ((BrailleIOViewRange)view).Parent.Name.Equals(WindowManager.BS_MINIMAP_NAME))
                {
                    if (WindowManager.Instance != null)
                    {
                        zoom = WindowManager.Instance.MinimapScalingFactor; // handling for minimap mode
                    }
                }

                Rectangle zPos = new Rectangle(
                    (int)(pos.X * zoom),
                    (int)(pos.Y * zoom),
                    (int)(pos.Width * zoom),
                    (int)(pos.Height * zoom)
                    );

                // add the panning offsets
                if (view is IPannable)
                {
                    zPos.X += ((IPannable)view).GetXOffset();
                    zPos.Y += ((IPannable)view).GetYOffset();
                }

                //TODO: decide in inner frame or outer frame
                int y1 = zPos.Y - 1;
                int y2 = zPos.Y + zPos.Height;
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

        void IBailleIORendererHook.PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams) 
        {
        
        }

        #endregion
    }
}
