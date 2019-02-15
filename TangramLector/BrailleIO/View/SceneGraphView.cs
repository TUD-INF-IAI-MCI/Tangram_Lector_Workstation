using BrailleIO.Interface;
using BrailleIO.Renderer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TangramLector.OO;
using tud.mci.tangram.TangramLector.BrailleIO.Controller;
using tud.mci.tangram.TangramLector.BrailleIO.Model;

namespace tud.mci.tangram.TangramLector.View
{
    class SceneGraphView : AbstractCachingRendererBase
    {


        MatrixBrailleRenderer brlRenderer = new MatrixBrailleRenderer(RenderingProperties.ADD_SPACE_FOR_SCROLLBARS | RenderingProperties.IGNORE_LAST_LINESPACE | RenderingProperties.RETURN_REAL_WIDTH);
        OoDrawModel _lastDrawModel = null;
        SceneGraphDialogController _sgc = null;
        internal SceneGraphDialogController SgController
        {
            get { return _sgc; }
            set {
                if (_sgc != value)
                {
                    var old = _sgc;
                    _sgc = value;
                    registerToEvents(_sgc);
                    unregisterFromEvents(old);
                    old.Dispose();
                } 
            }
        }

        ///// <summary>
        ///// Renders the current content
        ///// </summary>
        ///// <param name="view"></param>
        ///// <param name="content"></param>
        //public virtual void PrerenderMatrix(IViewBoxModel view, object content)
        //{
        //    if (view != null && content != null && content is OoDrawModel)
        //    {
        //        int trys = 0;
        //        Task t = new Task(() =>
        //        {
        //            while (IsRendering && trys++ < maxRenderingWaitTrys) { Thread.Sleep(renderingWaitTimeout); }
        //            IsRendering = true;
        //            ContentChanged = false;
        //            _cachedMatrix = _renderMatrix(view, content as OoDrawModel, CallHooksOnCacherendering);
        //            LastRendered = DateTime.Now;
        //            IsRendering = false;
        //        });
        //        t.Start();
        //    }
        //}



        protected override bool[,] renderMatrix(IViewBoxModel view, Object content, bool CallHooksOnCacherendering)
        {
            OoDrawModel ooDrawModel = content as OoDrawModel;            
            bool[,] m = new bool[0, 0];

            if (ooDrawModel == null) return m;

            int w = view.ContentBox.Width;
            // go through all elements of the drawing

            int levelIndent = 2;
            int lvl = 0;


            if (ooDrawModel != null && ooDrawModel.OoObserver != null)
            {
                var doc = ooDrawModel.OoObserver.GetActiveDocument();
                if (doc != null)
                {
                    var page = doc.GetActivePage();
                    if (page != null && page.shapeList != null)
                    {
                        List<bool[,]> lines = new List<bool[,]>();


                        // add Document title
                        bool[,] t = brlRenderer.RenderMatrix(w, "Doc: " + doc.Title, true);
                        lines.Add(t);

                        // add page number
                        bool[,] p = brlRenderer.RenderMatrix(w, "Page: " + page.GetPageNum(), true);
                        lines.Add(p);


                        lvl++;
                        // get all elements
                        var shapes = page.shapeList.ToArray();

                        string indetString = "  ";

                        foreach (var shape in shapes)
                        {
                            if (shape != null)
                            {
                                var shpText = indetString + OoElementSpeaker.GetElementDescriptionText(shape);
                                bool[,] shpM = brlRenderer.RenderMatrix(w, shpText);
                                lines.Add(shpM);
                            }
                        }



                        // build result

                        // check width and height





                    }
                }
            }

            return m;

        }


        #region build scene graph



        #endregion









        #region Events

        private void unregisterFromEvents(SceneGraphDialogController old)
        {
            
        }

        private void registerToEvents(SceneGraphDialogController _sgc)
        {

        }

        #endregion

    }
}
