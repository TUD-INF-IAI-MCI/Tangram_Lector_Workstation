using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BrailleIO.Renderer;
using System.Threading.Tasks;
using System.Threading;
using BrailleIO;
using System.Drawing;
using tud.mci.tangram.Uia;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.TangramLector.BrailleIO.Model;

namespace tud.mci.tangram.TangramLector.BrailleIO.View
{
    /// <summary>
    /// DRAW application observer renderer
    /// </summary>
    /// <seealso cref="BrailleIO.Interface.BrailleIOHookableRendererBase" />
    /// <seealso cref="BrailleIO.Interface.IBrailleIOContentRenderer" />
    class DrawRenderer : AbstractCachingRendererBase, ITouchableRenderer, IDisposable
    {
        #region Members
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

        public virtual object GetContentAtPosition(int x, int y)
        {




            //Point p = GetTapPositionOnScreen(x, x, lastView as BrailleIOViewRange);

            ////check if a OpenOffice Window is presented
            //if (ScreenObserver != null && ScreenObserver.Whnd != null)
            //{
            //    var observed = isObservedOpebnOfficeDrawWindow(ScreenObserver.Whnd.ToInt32());
            //    if (observed != null)
            //    {
            //        handleSelectedOoAccItem(observed, p, e);
            //        return;
            //    }
            //}

            //// if no OpenOfficeDraw Window could be found
            //var element = UiaPicker.GetElementFromScreenPosition(p.X, p.Y);

            //if (element != null)
            //{
            //    string className = element.Current.ClassName;
            //    if (String.IsNullOrEmpty(className) || className.Equals(OO_DOC_WND_CLASS_NAME))
            //    {
            //        if (handleTouchRequestForOO(element, p, e)) return;
            //    }

            //    //generic handling
            //    if (e.ReleasedGenericKeys.Count < 2 &&
            //        e.PressedGenericKeys.Count < 1 &&
            //        e.ReleasedGeneralKeys.Contains(BrailleIO_DeviceButton.Gesture))
            //    {
            //        UiaPicker.SpeakElement(element);
            //    }
            //}



            return null;



        }

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
    }
}
