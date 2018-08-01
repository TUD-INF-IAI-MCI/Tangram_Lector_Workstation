using BrailleIO;
using BrailleIO.Interface;
using BrailleIO.Renderer;
using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.TangramLector.BrailleIO.Model;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.Uia;

namespace tud.mci.tangram.TangramLector.BrailleIO.View
{
    /// <summary>
    /// DRAW application observer renderer
    /// </summary>
    /// <seealso cref="BrailleIO.Interface.BrailleIOHookableRendererBase" />
    /// <seealso cref="BrailleIO.Interface.IBrailleIOContentRenderer" />
    class DrawRenderer : AbstractCachingRendererBase, ITouchableRenderer, IDisposable, IContrastThreshold
    {
        #region Members
        /// <summary>
        /// The image renderer to use for tuning the screenshot into a pin-matrix
        /// </summary>
        public readonly BrailleIOImageToMatrixRenderer ImageRenderer = new BrailleIOImageToMatrixRenderer();
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="DrawRenderer"/> class.
        /// </summary>
        public DrawRenderer()
            : base()
        {
            DoesPanning = true;
        }

        #endregion

        #region Rendering

        /// <summary>
        /// Informs the renderer that the content the or view has changed.
        /// You have to call the PrerenderMatrix function manually if you want to have a cached result.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        public override void ContentOrViewHasChanged(IViewBoxModel view, object content)
        {
            if (lastContent != content)
            {
                unregisterFromDrawModelEvents(lastContent as OoDrawModel);
                lastContent = content;
                registerToDrawModelEvents(lastContent as OoDrawModel);
            }
            base.ContentOrViewHasChanged(view, content);
        }

        /// <summary>
        /// Gets the previously rendered and cached matrix.
        /// </summary>
        /// <returns>
        /// The cached rendering result
        /// </returns>
        public override bool[,] GetCachedMatrix()
        {
            //int trys = 0;
            //while (IsRendering && trys++ < 2) { Thread.Sleep(renderingWaitTimeout); }
            return _cachedMatrix;
        }

        /// <summary>
        /// Renders the current content
        /// </summary>
        /// <param name="view"></param>
        /// <param name="content"></param>
        public override void PrerenderMatrix(IViewBoxModel view, object content)
        {
            int trys = 0;
            Task t = new Task(() =>
            {
                while (IsRendering && trys++ < maxRenderingWaitTrys) { Thread.Sleep(renderingWaitTimeout); }
                this.IsRendering = true;
                ContentChanged = false;
                _cachedMatrix = _renderMatrix(view, content, CallHooksOnCacherendering);
                LastRendered = DateTime.Now;
                IsRendering = false;
            });
            t.Start();
        }

        /// <summary>
        /// Renders the real matrix.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="CallHooksOnCacherendering">if set to <c>true</c> [call hooks on cache rendering].</param>
        /// <returns></returns>
        protected virtual bool[,] _renderMatrix(global::BrailleIO.Interface.IViewBoxModel view, object content, bool CallHooksOnCacherendering)
        {
            if (content is OoDrawModel)
            {
                if (view is BrailleIOViewRange &&
                    ((BrailleIOViewRange)view).IsVisible() &&
                    ((BrailleIOViewRange)view).Parent != null &&
                    ((BrailleIOViewRange)view).Parent.IsVisible())
                {

                    // fix zoom -1 = fit to available space
                    if (((OoDrawModel)content).LastScreenCapturing != null
                        && view is IZoomable && ((IZoomable)view).GetZoom() < 0)
                    {
                        var Bounds = ((OoDrawModel)content).LastScreenCapturing.Size;
                        var factor = Math.Min(
                           (double)view.ContentBox.Height / (double)Bounds.Height + 0.000000001,
                            (double)view.ContentBox.Width / (double)Bounds.Width + 0.000000001
                            );
                        ((IZoomable)view).SetZoom(factor);
                    }

                    // set the contrast threshold
                    if (view is IContrastThreshold)
                        ImageRenderer.SetThreshold(
                            ((IContrastThreshold)view).GetContrastThreshold());

                    var result = ImageRenderer.RenderMatrix(view,
                        ((OoDrawModel)content).LastScreenCapturing);

                    return result;
                }
                else
                {
                    ContentChanged = true;
                }
            }
            return new bool[0, 0];
        }

        #endregion

        #region ITouchableRenderer

        /// <summary>
        /// Gets the Object at position x,y in the content.
        /// </summary>
        /// <param name="x">The x position in the content matrix.</param>
        /// <param name="y">The y position in the content matrix.</param>
        /// <returns>
        /// An object at the requester position in the content or <c>null</c>
        /// </returns>
        public virtual object GetContentAtPosition(int x, int y)
        {
            Point p = WindowManager.Instance.GetTapPositionOnScreen(x, x, lastView as BrailleIOViewRange);

            //check if a OpenOffice Window is presented
            if (lastContent is OoDrawModel &&
                ((OoDrawModel)lastContent).ScreenObserver != null &&
                ((OoDrawModel)lastContent).ScreenObserver.Whnd != null)
            {
                try
                {
                    if (OoConnector.Instance != null && OoConnector.Instance.Observer != null)
                    {
                        var observed = OoConnector.Instance.Observer.ObservesWHndl((int)((OoDrawModel)lastContent).ScreenObserver.Whnd);
                        if (observed != null && observed.DocumentComponent != null)
                        {
                            OoAccComponent shape = observed.DocumentComponent.GetAccessibleFromScreenPos(p);
                            if (shape != null)
                            {
                                return shape;
                                //if (observed.DrawPagesObs != null)
                                //{
                                //    var sObs = observed.DrawPagesObs.GetRegisteredShapeObserver(shape);
                                //}
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR]\t can't get touched DRAW object", e);
                }
            }

            try
            {
                // if no OpenOfficeDraw Window could be found
                AutomationElement element = UiaPicker.GetElementFromScreenPosition(p.X, p.Y);
                return element;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Get all Objects inside (or at least partial) the given area.
        /// NOT IMPLEMENTED!
        /// </summary>
        /// <param name="left">Left border of the region to test (X).</param>
        /// <param name="right">Right border of the region to test (X + width).</param>
        /// <param name="top">Top border of the region to test (Y).</param>
        /// <param name="bottom">Bottom border of the region to test (Y + height).</param>
        /// <returns>
        /// A list of elements inside or at least partial inside the requested area.
        /// </returns>
        public virtual System.Collections.IList GetAllContentInArea(int left, int right, int top, int bottom)
        {
            return null;
        }

        #endregion

        #region IDisposable

        public virtual void Dispose()
        {
            try
            {
            }
            catch (Exception) { }
        }

        #endregion

        #region events

        void registerToDrawModelEvents(OoDrawModel model)
        {
            if (model != null)
            {
                model.ScreenShotChanged += model_ScreenShotChanged;
            }
        }

        void unregisterFromDrawModelEvents(OoDrawModel model)
        {
            try
            {
                if (model != null)
                {
                    model.ScreenShotChanged -= model_ScreenShotChanged;
                }
            }
            catch (Exception) { }
        }

        void model_ScreenShotChanged(object sender, EventArgs e)
        {
            // ((OoDrawModel)lastContent).LastScreenCapturing.Save(@"C:\test\" + DateTime.Now.Ticks + "_capturing.png", System.Drawing.Imaging.ImageFormat.Png);

            PrerenderMatrix(lastView, lastContent);
        }

        #endregion

        #region IContrastThreshold

        /// <summary>
        /// Sets the contrast threshold.
        /// </summary>
        /// <param name="threshold">The threshold.</param>
        /// <returns>
        /// the new set threshold
        /// </returns>
        public int SetContrastThreshold(int threshold)
        {
            return ImageRenderer.SetContrastThreshold(threshold);
        }

        /// <summary>
        /// Gets the contrast threshold.
        /// </summary>
        /// <returns>
        /// the threshold
        /// </returns>
        public int GetContrastThreshold()
        {
            return ImageRenderer.GetContrastThreshold();
        }

        #endregion
    }
}
