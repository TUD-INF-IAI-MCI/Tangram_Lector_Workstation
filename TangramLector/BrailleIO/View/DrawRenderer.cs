using BrailleIO;
using BrailleIO.Interface;
using BrailleIO.Renderer;
using System;
using System.Drawing;
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
    class DrawRenderer : BrailleIOHookableRendererBase, IBrailleIOContentRenderer, ITouchableRenderer, IDisposable, IContrastThreshold, IBrailleIOPanningRendererInterfaces // AbstractCachingRendererBase, ITouchableRenderer, IDisposable, IContrastThreshold
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

        IViewBoxModel _lastView = null;
        
        OoDrawModel lastContent = null;
        bool[,] _lastResult = new bool[0,0];

        #region IBrailleIOContentRenderer

        /// <summary>
        /// Renders a content object into an boolean matrix;
        /// while <c>true</c> values indicating raised pins and <c>false</c> values indicating lowered pins
        /// </summary>
        /// <param name="view">The frame to render in. This gives access to the space to render and other parameters. Normally this is a BrailleIOViewRange.</param>
        /// <param name="matrix">The content to render.</param>
        /// <returns>
        /// A two dimensional boolean M x N matrix (bool[M,N]) where M is the count of rows (this is height)
        /// and N is the count of columns (which is the width).
        /// Positions in the Matrix are of type [i,j]
        /// while i is the index of the row (is the y position)
        /// and j is the index of the column (is the x position).
        /// In the matrix <c>true</c> values indicating raised pins and <c>false</c> values indicating lowered pins
        /// </returns>
        public virtual bool[,] RenderMatrix(IViewBoxModel view, bool[,] matrix)
        {
            return _lastResult;
        }

        /// <summary>
        /// Renders the real matrix.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="CallHooksOnCacherendering">if set to <c>true</c> [call hooks on cache rendering].</param>
        /// <returns></returns>
        public bool[,] RenderMatrix(IViewBoxModel view, object content)
        {
            bool[,] result = new bool[0, 0];

            callAllPreHooks(ref view, ref content, null);

            if (content is OoDrawModel)
            {
                lastContent = content as OoDrawModel;

                if (view is BrailleIOViewRange &&
                    ((BrailleIOViewRange)view).IsVisible() &&
                    ((BrailleIOViewRange)view).Parent != null &&
                    ((BrailleIOViewRange)view).Parent.IsVisible())
                {
                    using (Image capt = ((OoDrawModel)content).LastScreenCapturing)
                    {

                        // fix zoom -1 = fit to available space
                        if (capt != null
                            && ((IZoomable)view).GetZoom() < 0)
                        {
                            var Bounds = capt.Size;
                            var factor = Math.Min(
                               view.ContentBox.Height / (double)Bounds.Height + 0.000000001,
                                view.ContentBox.Width / (double)Bounds.Width + 0.000000001
                                );
                            ((IZoomable)view).SetZoom(factor);
                        }

                        // set the contrast threshold
                        ImageRenderer.SetThreshold(
                                ((IContrastThreshold)view).GetContrastThreshold());

                        ImageRenderer.Invert = ((BrailleIOViewRange)view).InvertImage;


                        result = ImageRenderer.RenderMatrix(view, capt);
                    }
                }
            }

            if(result.Length > 0)
            {
                _lastResult = result;
            }
            else // if something went wrong with the image rendering or capturing
            {
                // repeat the last result
                result = _lastResult;
            }

            callAllPostHooks(view, content, ref result, null);

            return result;
        }

        #endregion

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
            Point p = WindowManager.Instance.GetTapPositionOnScreen(x, x, _lastView as BrailleIOViewRange);

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

            // PrerenderMatrix(_lastView, lastContent);
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

        #region IBrailleIOPanningRendererInterfaces

        bool _doesPanning = false;
        /// <summary>
        /// Indicates to the combining renderer if this renderer handles panning by its own or not.
        /// <c>true</c> means the renderer has already handled panning (offsets) and returns the correct result.
        /// <c>false</c> means the render does not handle panning (offset), returns the whole rendering result
        /// and the combination renderer has to take care about the panning (offsets)
        /// </summary>
        public virtual bool DoesPanning
        {
            get { return _doesPanning; }
            protected set { _doesPanning = value; }
        }

        #endregion

    }
}
