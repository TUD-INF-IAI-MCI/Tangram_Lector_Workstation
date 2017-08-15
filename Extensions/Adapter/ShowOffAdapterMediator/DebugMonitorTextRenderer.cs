using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using BrailleIO;
using BrailleIO.Renderer;
using BrailleIO.Renderer.Structs;
using tud.mci.tangram.TangramLector;
using tud.mci.tangram.audio;

namespace ShowOffAdapterMediator
{
    class DebugMonitorTextRenderer : IDisposable
    {
        #region Members

        // WindowManager wm;

        Thread renderingThread;
        volatile bool _run = true;
        private readonly IBrailleIOShowOffMonitor monitor;
        BrailleIOMediator io = null;

        const int PIN_2_PIXEL_FACTOR = 6;

        #endregion

        #region Constructor
        /// <summary>
        /// Class for displaying readable text for BrailleTexts in the debug monitor
        /// </summary>
        /// <param name="monitor"></param>
        public DebugMonitorTextRenderer(IBrailleIOShowOffMonitor monitor, BrailleIOMediator io)
        {
            this.monitor = monitor;
            this.io = io;
            preparePen();
            startRenderingThread();
        }

        private static void preparePen()
        {
            if (dashedPen != null)
            {
                dashedPen.Width = 1.5f;
                dashedPen.DashCap = System.Drawing.Drawing2D.DashCap.Round;
                dashedPen.DashPattern = new float[] { 4.0F, 4.0f };
            }
        }

        #endregion

        #region Rendering Thread

        void startRenderingThread()
        {
            if (renderingThread == null || !renderingThread.IsAlive || renderingThread.ThreadState != ThreadState.Running)
            {
                _run = false;
                renderingThread = new Thread(renderText);
                renderingThread.Name = "DebugMonitorTextRendering";
                renderingThread.Priority = ThreadPriority.Normal;
                renderingThread.IsBackground = true;
                _run = true;
                renderingThread.Start();
            }
        }

        void renderText()
        {
            while (_run)
            {
                try
                {
                    Bitmap im = new Bitmap(721, 363);
                    Graphics g = Graphics.FromImage(im);

                    g.Clear(Color.Transparent);
                    List<BrailleIOViewRange> vrs = getAllViews();

                    foreach (var vr in vrs)
                    {
                        int x, y, width, height;

                        if (vr != null && vr.ContentRender is ITouchableRenderer)
                        {
                            //get visible area in view Range;
                            Rectangle r = vr.ViewBox;
                            Rectangle cb = vr.ContentBox;

                            int left = -vr.GetXOffset();
                            int top = -vr.GetYOffset();
                            int right = left + cb.Width;
                            int bottom = top + cb.Height;

                            //get vr position on screen
                            x = r.X + cb.X - left;
                            y = r.Y + cb.Y - top;
                            width = cb.Width;
                            height = cb.Height;

                            var items = getRenderingElementsFromViewRange(vr, left, right, top, bottom);

                            List<RenderElement> elements = items as List<RenderElement>;
                            if (items != null && items.Count > 0)
                            {
                                foreach (var e in elements)
                                {
                                    getImageOfRenderingElement(e, ref g, x, y, left, right, top, bottom);
                                }
                            }
                        }
                    }

                    g.Dispose();

                    //im = ChangeOpacity(im, 0.8f);

                    if (monitor != null)
                    {
                        monitor.SetPictureOverlay(im);
                    }
                }
                catch (System.Exception ex)
                {
                    // return;
                    continue;
                }
                finally
                {
                    Thread.Sleep(200);
                }
            }

            System.Diagnostics.Debug.WriteLine("Thread beendet");

        }

        #endregion

        /// <summary>
        /// get all the visible viewRanges in the currently visible screen containing text.
        /// </summary>
        /// <returns>list of ViewRanges containing text</returns>
        List<BrailleIOViewRange> getAllViews()
        {
            List<BrailleIOViewRange> vrs = new List<BrailleIOViewRange>();
            var screens = io.GetActiveViews();

            foreach (var item in screens)
            {
                if (item is BrailleIOScreen)
                {
                    foreach (var vrPair in ((BrailleIOScreen)item).GetOrderedViewRanges())
                    {
                        if (vrPair.Value != null)
                        {
                            BrailleIOViewRange vr = vrPair.Value;
                            if (vr.IsVisible() && (vr.IsText() || vr.IsOther()))
                            {
                                vrs.Add(vr);
                            }
                        }
                    }
                }
                else if (item is BrailleIOViewRange)
                {
                    BrailleIOViewRange vr = item as BrailleIOViewRange;
                    if (vr.IsVisible() && (vr.IsText() || vr.IsOther()))
                    {
                        vrs.Add(vr);
                    }
                }
            }

            return vrs;
        }

        //BrailleIOScreen GetVisibleScreen()
        //{
        //    if (monitor != null && io != null)
        //    {
        //        var views = io.GetViews();
        //        if (views != null && views.Count > 0)
        //        {
        //            try
        //            {
        //                //var vs = views.First(x => (x is BrailleIO.Interface.IViewable && ((BrailleIO.Interface.IViewable)x).IsVisible()));

        //                object vs = null;

        //                for (int i = views.Count - 1; i >= 0; i--)
        //                {
        //                    object x = views[i];
        //                    if ((x is BrailleIO.Interface.IViewable && ((BrailleIO.Interface.IViewable)x).IsVisible()))
        //                    {
        //                        vs = x;
        //                        break;
        //                    }
        //                }

        //                if (vs != null && vs is BrailleIOScreen)
        //                {
        //                    return vs as BrailleIOScreen;
        //                }
        //            }
        //            catch (InvalidOperationException) { } //Happens if no view could been found in the listing
        //        }
        //    }
        //    return null;
        //}


        static IList getRenderingElementsFromViewRange(BrailleIOViewRange vr, int left, int right, int top, int bottom)
        {
            if (vr != null && vr.ContentRender is ITouchableRenderer)
            {
                return ((ITouchableRenderer)vr.ContentRender).GetAllContentInArea(left, right, top, bottom);
            }
            return null;
        }

        #region Image Rendering

        static Font textFont = new Font(FontFamily.GenericMonospace, 14, FontStyle.Bold);
        static SolidBrush textBrush = new SolidBrush(Color.FromArgb(255, 10, 10, 10));
        static SolidBrush backgroundBrush = new SolidBrush(Color.FromArgb(180, 255, 255, 255));
        static Pen dashedPen = new Pen(Color.FromArgb(200, 139, 0, 139));

        static void getImageOfRenderingElement(RenderElement e, ref Graphics g, int xOffset, int yOffset,
            int left, int right, int top, int bottom,
            Pen p = null, int xShrink = 0, int yShrink = 0)
        {
            if (p == null) p = dashedPen != null ? dashedPen : Pens.DarkMagenta;

            if (e.Height > 0 && e.Width > 0 && g != null)
            {
                g.DrawRectangle(p
                    , (e.X + xOffset) * PIN_2_PIXEL_FACTOR + xShrink
                    , (e.Y + yOffset) * PIN_2_PIXEL_FACTOR + yShrink
                    , e.Width * PIN_2_PIXEL_FACTOR - (2 * xShrink)
                    , e.Height * PIN_2_PIXEL_FACTOR - (2 * yShrink)
                    );

                if (e.HasSubParts())
                {
                    foreach (var item in e.GetSubParts())
                    {
                        if (item.IsCompletelyInArea(left, right, top, bottom))
                        {
                            getImageOfRenderingElement(item, ref g, xOffset, yOffset, left, right, top, bottom, Pens.DarkTurquoise, xShrink + 1, yShrink + 1);
                        }
                    }
                }
                else
                {
                    if (e.GetValue() != null)
                    {
                        String value = String.IsNullOrEmpty(e.DisplayName) ? e.GetValue().ToString() : e.DisplayName;
                        if (!String.IsNullOrWhiteSpace(value))
                        {
                            int x = (e.X + xOffset) * PIN_2_PIXEL_FACTOR + (2 * xShrink);
                            int y = (e.Y + yOffset) * PIN_2_PIXEL_FACTOR + (2 * yShrink);
                            int width = e.Width * PIN_2_PIXEL_FACTOR - (4 * xShrink);
                            int height = e.Height * PIN_2_PIXEL_FACTOR - (4 * yShrink);

                            g.FillRectangle(backgroundBrush, new Rectangle(x, y, width, height));
                            g.DrawString(value, textFont, textBrush, new Rectangle(x, y, width, height));
                        }
                    }
                }
            }
        }

        public static Bitmap ChangeOpacity(Image img, float opacityvalue)
        {
            Bitmap bmp = new Bitmap(img.Width, img.Height); // Determining Width and Height of Source Image
            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                ColorMatrix colormatrix = new ColorMatrix();
                colormatrix.Matrix33 = opacityvalue;
                ImageAttributes imgAttribute = new ImageAttributes();
                imgAttribute.SetColorMatrix(colormatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                graphics.DrawImage(img, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, imgAttribute);
            }
            return bmp;
        }

        #endregion

        void IDisposable.Dispose()
        {
            _run = false;
            try
            {
                if (renderingThread != null)
                {
                    renderingThread.Abort();
                    if (monitor != null) monitor.Dispose();
                }
            }
            catch { }
        }
    }
}
