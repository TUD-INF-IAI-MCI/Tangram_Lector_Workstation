using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using BrailleIO;
using BrailleIO.Interface;
using tud.mci.tangram;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector;
using tud.mci.tangram.TangramLector.OO;

namespace tud.mci.tangram.TangramLector.OO
{
    public class BrailleTextView : IBailleIORendererHook
    {
        #region Members

        /// <summary>
        /// The braille renderer used for the text replacement
        /// </summary>
        internal static readonly BrailleIO.Renderer.MatrixBrailleRenderer BrailleRenderer = new BrailleIO.Renderer.MatrixBrailleRenderer(
            new BrailleIO.Renderer.BrailleInterpreter.SimpleBrailleInterpreter()
            , BrailleIO.Renderer.RenderingProperties.RETURN_REAL_WIDTH
        );

        /// <summary>
        /// Activates or deactivates the renderer
        /// </summary>
        public bool Active { get; set; }

        private BrailleIOViewRange _lastView = null;

        #endregion

        public BrailleTextView()
        {
            Active = true;
        }

        #region IBailleIORendererHook

        void IBailleIORendererHook.PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams)
        {
            return;
        }

        void IBailleIORendererHook.PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            _lastView = view as BrailleIOViewRange;
            if (Active
                && CanBeActivated()
                )
            {
                List<TextElemet> visibleTexts = getVisibleTextElements(view);
                foreach (var visibleText in visibleTexts)
                {
                    renderTextFieldInMatrix(ref result, visibleText, ((BrailleIOViewRange)view).GetXOffset(), ((BrailleIOViewRange)view).GetYOffset(), ((BrailleIOViewRange)view).GetZoom());
                }
                visibleTexts = null;
            }
        }

        #endregion

        /// <summary>
        /// Check if current view is in print zoom and, therefore, BrailleTextRenderer can be activated.
        /// </summary>
        /// <returns></returns>
        public bool CanBeActivated()
        {
            if (_lastView != null && _lastView.GetZoom() == WindowManager.GetPrintZoomLevel())
            {
                return true;
            }
            return false;
        }

        private List<TextElemet> getVisibleTextElements(IViewBoxModel view)
        {
            List<TextElemet> texts = getAllTextElementsOfDoc();

            List<TextElemet> visibleTexts = new List<TextElemet>();
            foreach (var text in texts)
            {
                if (IsElementsBoundingBoxVisibleInView(view as BrailleIOViewRange, text.ScreenPosition))
                {
                    visibleTexts.Add(text);
                }
            }

            return visibleTexts;
        }

        #region Rendering

        private static void renderTextFieldInMatrix(ref bool[,] m, TextElemet text, int xOffset, int yOffset, double zoom)
        {
            if (m != null && !String.IsNullOrWhiteSpace(text.Text))
            {
                bool[,] matrix = m;
                Rectangle startpoint = BrailleTextView.MakeZoomBoundingBox(text.ScreenPosition, zoom);
                // calculate the center point and from that the staring point
                startpoint.X += startpoint.Width / 2;
                startpoint.X -= text.Matrix.GetLength(1) / 2;

                startpoint.Y += startpoint.Height / 2;
                startpoint.Y -= text.Matrix.GetLength(0) / 2;

                // paint a spacing frame around the text

                Parallel.For(0, text.Matrix.GetLength(1) + 1, (i) =>
                {
                    int x = startpoint.X + xOffset - 1 + i;
                    if (x > 0 && matrix.GetLength(1) > x)
                    {
                        try
                        {
                            int y1 = startpoint.Y + yOffset - 1;
                            if (y1 > 0 && y1 < matrix.GetLength(0)) matrix[y1, x] = false;

                            int y2 = startpoint.Y + yOffset + text.Matrix.GetLength(0) - 1; // remove the -1 if underlining should be possible
                            if (y2 > 0 && y2 < matrix.GetLength(0)) matrix[y2, x] = false;
                        }
                        catch { }
                    }
                }
                );

                Parallel.For(0, text.Matrix.GetLength(0), (j) =>
                {
                    int y = startpoint.Y + yOffset + j;

                    if (y > 0 && y < matrix.GetLength(0))
                    {
                        try
                        {
                            int x1 = startpoint.X + xOffset - 1;
                            if (x1 > 0 && x1 < matrix.GetLength(1)) matrix[y, x1] = false;
                        }
                        catch { }
                    }
                }
                );

                // add the rendered text in the result               

                Parallel.For(0, text.Matrix.GetLength(1), (i) =>
                {
                    int x = i + startpoint.X + xOffset;
                    if (x > 0 && matrix.GetLength(1) > x)
                    {
                        Parallel.For(0, text.Matrix.GetLength(0), (j) =>
                        {
                            int y = startpoint.Y + j + yOffset;
                            if (y > 0 && matrix.GetLength(0) > y)
                            {
                                try
                                {
                                    matrix[y, x] = text.Matrix[j, i];
                                }
                                catch { }
                            }
                        }
                        );
                    }
                }
                );
                m = matrix;
            }
        }

        #endregion

        ///// <summary>
        ///// Test if the given view is related to a filtered DRAW document
        ///// </summary>
        ///// <param name="view"></param>
        ///// <returns><c>true</c> if the view is related to a DRAW document</returns>
        //private bool isViewADrawDoc(BrailleIOViewRange view)
        //{
        //    if (view != null)
        //    {
        //        throw new NotImplementedException();
        //    }
        //    return false;
        //}

        #region Cache of text Elements

        private readonly Object cachingLock = new Object();
        private List<TextElemet> _cachedtextElements = new List<TextElemet>();
        private List<TextElemet> cachedtextElements
        {
            get
            {
                lock (cachingLock)
                {
                    return _cachedtextElements;
                }
            }
            set
            {
                lock (cachingLock)
                {
                    _cachedtextElements = value;
                }
            }
        }

        private DateTime lastCache = DateTime.Now;
        private readonly TimeSpan cachingPeriode = new TimeSpan(0, 0, 3);

        private volatile OoAccessibleDocWnd lastCachedDoc = null;

        /// <summary>
        /// Get a list of available Observers of elements containing text. This listed is cached.
        /// </summary>
        /// <returns>a cached list of text containing elements</returns>
        private List<TextElemet> getAllTextElementsOfDoc()
        {
            List<TextElemet> textElements = new List<TextElemet>();
            OoAccessibleDocWnd doc = null;
            try
            {
                doc = OoConnector.Instance.Observer.GetActiveDocument();
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] can't get active DRAW document", ex);
            }

            if (doc != null)
            {

                //check for caching
                DateTime now = DateTime.Now;

                if (doc != lastCachedDoc || (now - lastCache > cachingPeriode))
                {
                    // get text elements
                    var pageObs = doc.GetActivePage();

                    if (pageObs != null)
                    {
                        // clone the list to enable disposing of elements from the list without destroying the iterator
                        List<OoShapeObserver> shapes = new List<OoShapeObserver>(pageObs.shapeList);

                        if (shapes != null && shapes.Count > 0)
                        {
                            foreach (OoShapeObserver shape in shapes)
                            {
                                if (shape != null && !shape.Disposed )
                                {
                                   // if()
                                        if (shape.HasText && shape.IsVisible())
                                        {
                                            textElements.Add(new TextElemet(shape));
                                        }
                                        //else
                                        //{
                                        //    //if (!shape.IsValid()) 
                                        //    shape.Dispose();
                                        //}
                                }
                            }
                        }

                        // cache
                        lastCache = now;
                        lastCachedDoc = doc;
                        cachedtextElements = textElements;
                    }
                }
            }

            return cachedtextElements;
        }

        /// <summary>
        /// Resets the cache for text elements.
        /// </summary>
        public void ResetCache()
        {
            lastCachedDoc = null;
        }

        #endregion

        #region Element is visible

        /// <summary>
        /// Determines whether [is elements bounding box visible in view] [the specified view].
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="shape">The shape.</param>
        /// <returns>
        /// 	<c>true</c> if [is elements bounding box visible in view] [the specified view]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsElementsBoundingBoxVisibleInView(BrailleIOViewRange view, OoShapeObserver shape)
        {
            if (view != null && shape != null)
            {
                return IsElementsBoundingBoxVisibleInView(GetViewPort(view), shape, Math.Max(0.0000001, view.GetZoom()));
            }
            return false;
        }

        /// <summary>
        /// Determines whether [is elements bounding box visible in view] [the specified view].
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="screenPos">The screen pos.</param>
        /// <returns>
        /// 	<c>true</c> if [is elements bounding box visible in view] [the specified view]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsElementsBoundingBoxVisibleInView(BrailleIOViewRange view, Rectangle screenPos)
        {
            if (view != null && !screenPos.IsEmpty)
            {
                return IsElementsBoundingBoxVisibleInView(GetViewPort(view), screenPos, Math.Max(0.0000001, view.GetZoom()));
            }
            return false;
        }

        /// <summary>
        /// Determines whether [is elements bounding box visible in view] [the specified visible area].
        /// </summary>
        /// <param name="visibleArea">The visible area.</param>
        /// <param name="shape">The shape.</param>
        /// <returns>
        /// 	<c>true</c> if [is elements bounding box visible in view] [the specified visible area]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsElementsBoundingBoxVisibleInView(Rectangle visibleArea, OoShapeObserver shape, double zoomFactor = 1)
        {
            if (visibleArea != null && !visibleArea.IsEmpty && shape != null)
            {
                Rectangle screenPos = shape.GetRelativeScreenBoundsByDom();
                Rectangle relativeScreenPos = MakeZoomBoundingBox(screenPos, zoomFactor);
                return DoBoundingBoxesCollide(visibleArea, relativeScreenPos);
            }

            return false;
        }

        /// <summary>
        /// Determines whether [is elements bounding box visible in view] [the specified visible area].
        /// </summary>
        /// <param name="visibleArea">The visible area.</param>
        /// <param name="screenPos">The screen pos.</param>
        /// <param name="zoomFactor">The zoom factor.</param>
        /// <returns>
        /// 	<c>true</c> if [is elements bounding box visible in view] [the specified visible area]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsElementsBoundingBoxVisibleInView(Rectangle visibleArea, Rectangle screenPos, double zoomFactor = 1)
        {
            if (visibleArea != null && !visibleArea.IsEmpty)
            {
                Rectangle relativeScreenPos = MakeZoomBoundingBox(screenPos, zoomFactor);
                return DoBoundingBoxesCollide(visibleArea, relativeScreenPos);
            }

            return false;
        }

        /// <summary>
        /// Does two bounding boxes collide.
        /// </summary>
        /// <param name="bBox1">The b box1.</param>
        /// <param name="bBox2">The b box2.</param>
        /// <returns></returns>
        public static bool DoBoundingBoxesCollide(Rectangle bBox1, Rectangle bBox2)
        {
            /*
             * bBox1
             * ┌────────────────────┐
             * │                    │
             * │        ┌───────────┼───┐
             * │        │           │   │
             * └────────┼───────────┘   │
             *          │               │
             *          └───────────────┘
             *          bBox2
             */

            //check x
            if (bBox2.Left > bBox1.Right
                || bBox1.Left > bBox2.Right)
            {
                return false;
            }
            //check y
            if (bBox2.Top > bBox1.Bottom ||
                bBox1.Top > bBox2.Bottom
                )
            {
                return false;
            }
            return true;
        }


        /// <summary>
        /// Gets the view port of a given view. The viewPort is the 
        /// presented part of the content.
        /// 
        /// ┌────────────────────┐ -- content 
        /// │                    │
        /// │   ╔════════╗╌╌ viewPort
        /// │   ║        ║       │
        /// │   ╚════════╝       │
        /// │                    │
        /// └────────────────────┘
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        static Rectangle GetViewPort(BrailleIOViewRange view)
        {
            Rectangle viewPort = new Rectangle();

            if (view != null)
            {
                viewPort.Width = view.ContentBox.Width;
                viewPort.Height = view.ContentBox.Height;
                viewPort.X = -view.GetXOffset();
                viewPort.Y = -view.GetYOffset();
            }

            return viewPort;
        }

        internal static Rectangle MakeZoomBoundingBox(Rectangle bBox, double zoomFactor)
        {
            if (zoomFactor <= 0)
            {
                throw new ArgumentException("Zoom factor must be a positive value grater then zero.", "zoomFactor");
            }

            Rectangle newBBox = bBox;

            if (bBox != null && !bBox.IsEmpty)
            {
                newBBox.X = (int)Math.Round((double)newBBox.X * zoomFactor);
                newBBox.Y = (int)Math.Round((double)newBBox.Y * zoomFactor);
                newBBox.Width = (int)Math.Round((double)newBBox.Width * zoomFactor);
                newBBox.Height = (int)Math.Round((double)newBBox.Height * zoomFactor);
            }

            return newBBox;
        }

        #endregion
    }

    struct TextElemet
    {
        /// <summary>
        /// top left position of the <see cref="ObjectBoundingBox"/>
        /// </summary>
        public Point Position;
        /// <summary>
        /// rendered Braille text matrix
        /// </summary>
        public bool[,] Matrix;
        /// <summary>
        /// Position of the object in the content (in px)
        /// </summary>
        public Rectangle ScreenPosition;
        /// <summary>
        /// Screen position with respect to the print zoom level of the Window manager
        /// </summary>
        public Rectangle ObjectBoundingBox;
        /// <summary>
        /// Center position of the <see cref="ObjectBoundingBox"/>
        /// </summary>
        public Point Center;
        /// <summary>
        /// text value
        /// </summary>
        public String Text;
        /// <summary>
        /// the shape containing the text value
        /// </summary>
        public OoShapeObserver Shape;

        public TextElemet(OoShapeObserver shape)
        {
            Shape = shape;
            if (shape != null)
            {
                Text = shape.Text;
                ScreenPosition = shape.GetRelativeScreenBoundsByDom();
                ObjectBoundingBox = BrailleTextView.MakeZoomBoundingBox(ScreenPosition, WindowManager.GetPrintZoomLevel());
                Matrix = BrailleTextView.BrailleRenderer.RenderMatrix(ObjectBoundingBox.Width, Text);
                Position = new Point(ObjectBoundingBox.X, ObjectBoundingBox.Y);
                Center = new Point(ObjectBoundingBox.X + (ObjectBoundingBox.Width / 2), ObjectBoundingBox.Y + (ObjectBoundingBox.Height / 2));
            }
            else
            {
                Center = Position = new Point();
                Matrix = new bool[0, 0];
                ObjectBoundingBox = ScreenPosition = new Rectangle();
                Text = String.Empty;
            }
        }
    }
}

