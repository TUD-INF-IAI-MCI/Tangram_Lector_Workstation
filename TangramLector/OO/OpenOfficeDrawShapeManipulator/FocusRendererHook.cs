using System;
using System.Drawing;
using System.Threading.Tasks;
using BrailleIO.Interface;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.util;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    public class FocusRendererHook : IBailleIORendererHook
    {
        bool _paintControlPointReferenceLine = true;
        /// <summary>
        /// Gets or sets a value indicating whether to paint a dotted reference line 
        /// between a control point and its related edge point.
        /// </summary>
        /// <value>
        ///   <c>true</c> if set to <c>true</c> an reference line is drawn; otherwise set it to <c>false</c>.
        /// </value>
        public bool PaintControlPointReferenceLine
        {
            get { return _paintControlPointReferenceLine; }
            set { _paintControlPointReferenceLine = value; }
        }


        private bool _dashed = false;
        /// <summary>
        /// Initializes a new instance of the <see cref="FocusRendererHook"/> class.
        /// </summary>
        /// <param name="drawDashedBox">if set to <c>true</c> the box is drawn as a dashed outline box.</param>
        public FocusRendererHook(bool drawDashedBox = false)
        {
            _dashed = drawDashedBox;
        }


        /// <summary>
        /// The current bounding box to display
        /// </summary>
        private Rectangle _currentBoundingBox = new Rectangle(-1, -1, 0, 0);
        /// <summary>
        /// The current bounding box to display
        /// </summary>
        virtual public Rectangle CurrentBoundingBox
        {
            get { return _currentBoundingBox; }
            set { _currentBoundingBox = value; }
        }

        private OoPolygonPointsObserver _currentPolyPoint = null;
        /// <summary>
        /// Gets or sets the current polygon point.
        /// </summary>
        /// <value>
        /// The current polygon point.
        /// </value>
        public OoPolygonPointsObserver CurrentPolyPoint
        {
            get { return _currentPolyPoint; }
            set { _currentPolyPoint = value; }
        }

        OoPolygonPointsObserver _relatedPolyPoint = null;

        /// <summary>
        /// Gets the related polygon edge point to the <see cref="CurrentPolyPoint"/> if it is a Control point.
        /// </summary>
        /// <value>
        /// The related polygon edge point.
        /// </value>
        public PolyPointDescriptor RelatedPolyPoint
        {
            get
            {
                int i;
                PolyPointDescriptor relPoint = new PolyPointDescriptor();
                if (CurrentPolyPoint != null)
                {
                    var PointDescriptor = CurrentPolyPoint.Current(out i);
                    if (i > 0 && !PointDescriptor.IsEmpty && PointDescriptor.Flag == PolygonFlags.CONTROL)
                    {
                        // get previous point
                        relPoint = CurrentPolyPoint[i - 1];
                        if (relPoint.IsEmpty || relPoint.Flag == PolygonFlags.CONTROL)
                        {
                            if (CurrentPolyPoint.Count > (i + 1))
                            {
                                relPoint = CurrentPolyPoint[i + 1];
                                if (relPoint.IsEmpty || relPoint.Flag == PolygonFlags.CONTROL)
                                    relPoint = new PolyPointDescriptor();
                            }
                            else relPoint = new PolyPointDescriptor();
                        }
                    }
                }
                return relPoint;
            }
        }

        /// <summary>
        /// The current point
        /// </summary>
        private Point CurrentPoint
        {
            get
            {
                int trash;
                if (CurrentPolyPoint != null)
                {
                    var PointDescriptor = CurrentPolyPoint.Current(out trash);
                    return CurrentPolyPoint.TransformPointCoordinatesIntoScreenCoordinates(PointDescriptor);
                }
                else
                {
                    return new Point(-1, -1);
                }
            }
        }


        /// <summary>
        /// Sets the current bounding box by shape.
        /// </summary>
        /// <param name="currentSelectedShape">The current selected shape.</param>
        public void SetCurrentBoundingBoxByShape(OoShapeObserver currentSelectedShape)
        {
            if (currentSelectedShape != null && currentSelectedShape.IsValid(false))
            {
                CurrentBoundingBox = currentSelectedShape.GetAbsoluteScreenBoundsByDom();
            }
            else CurrentBoundingBox = new Rectangle(-1, -1, 0, 0);

        }

        /// <summary>
        /// Determine if this Hook should be active or not.
        /// </summary>
        public bool Active = false;

        private bool _doRenderBoundingBox = false;
        /// <summary>
        /// Determine if this Hook should render the Bounding box or not.
        /// </summary>
        public bool DoRenderBoundingBox
        {
            get { return _doRenderBoundingBox; }
            set { _doRenderBoundingBox = value; }
        }

        /// <summary>
        /// This hook function is called by an IBrailleIOHookableRenderer before he starts his rendering.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="content">The content.</param>
        /// <param name="additionalParams">Additional parameters.</param>
        void IBailleIORendererHook.PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams)
        {
        }

        // Result is addressed in [y, x] notation.
        void IBailleIORendererHook.PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            if (Active && _doRenderBoundingBox
                && WindowManager.Instance != null && !WindowManager.Instance.IsInMinimapMode())
            {
                doBlinkingBoundingBox(view, content, ref result, additionalParams);
            }
        }

        const int MIN_FRAME_SIZE = 4;
        /// <summary>
        /// Draw a blinking frame around the current selected shape.
        /// </summary>
        /// <param name="view"></param>
        /// <param name="content"></param>
        /// <param name="result"></param>
        /// <param name="additionalParams"></param>
        virtual protected void doBlinkingBoundingBox(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            //draw frame as bool pins
            if (_doRenderBoundingBox &&
                CurrentPoint.X < 0 && CurrentPoint.Y < 0 &&
                !(CurrentBoundingBox.Width * CurrentBoundingBox.Height < 1))
            {
                result = paintBoundingBoxMarker(view, result);
            }
            else
            {
                result = paintPolygonPointMarker(view, result);
            }
        }

        static int boundingBoxPadding = 2;

        virtual protected bool[,] paintBoundingBoxMarker(IViewBoxModel view, bool[,] target)
        {
            if (view is IZoomable && view is IPannable)
            {
                double zoom = ((IZoomable)view).GetZoom();
                int xOffset = ((IPannable)view).GetXOffset();
                int yOffset = ((IPannable)view).GetYOffset();

                Rectangle pageBounds = getPageBounds();

                if (pageBounds.Width > 0 && pageBounds.Height > 0)
                {
                    // coords of the shapes bounding box, relative to the whole captured image
                    Rectangle relbBox = new Rectangle(CurrentBoundingBox.X - pageBounds.X, CurrentBoundingBox.Y - pageBounds.Y, CurrentBoundingBox.Width, CurrentBoundingBox.Height);
                    // converted to braille output coords, as shown in the original view (with zoom factor and panning position applied)
                    Rectangle out_bBox = new Rectangle(
                        (int)(Math.Round(((double)relbBox.X * zoom) + xOffset - boundingBoxPadding)),     // x
                        (int)(Math.Round(((double)relbBox.Y * zoom) + yOffset - boundingBoxPadding)),     // y
                        (int)(Math.Round(((double)relbBox.Width * zoom) + 2 * boundingBoxPadding)),       // w
                        (int)(Math.Round(((double)relbBox.Height * zoom) + 2 * boundingBoxPadding)));     // h

                    // check for minimal height and width
                    if (out_bBox.Width < MIN_FRAME_SIZE)
                    {
                        int oldWidth = out_bBox.Width;
                        out_bBox.Width = MIN_FRAME_SIZE;
                        double change = out_bBox.Width - oldWidth;
                        out_bBox.X -= (int)Math.Round((change) / 2);
                    }

                    if (out_bBox.Height < MIN_FRAME_SIZE)
                    {
                        int oldHeight = out_bBox.Height;
                        out_bBox.Height = MIN_FRAME_SIZE;
                        double change = out_bBox.Height - oldHeight;

                        out_bBox.Y -= (int)Math.Round((change) / 2);
                    }

                    if (out_bBox.Width > 0 && out_bBox.Height > 0)
                    {
                        int result_x_max = target.GetLength(1) - 1;
                        int result_y_max = target.GetLength(0) - 1;
                        // rectangle coords within matrix:
                        int x1 = Math.Max(0, out_bBox.X);                                // from bBox x (or 0)
                        int x2 = Math.Min(out_bBox.X + out_bBox.Width, result_x_max);    // to bBox x + w (or x rightmost of matrix)
                        int y1 = Math.Max(0, out_bBox.Y);                                // from bBox y (or 0)
                        int y2 = Math.Min(out_bBox.Y + out_bBox.Height, result_y_max);   // to bBox y + h (or y bottom of matrix)

                        if (y2 >= y1 && x2 >= x1)
                        {
                            /* draw outer (inflated box) and inner (deflated box) lowered pins around the raised pins bounding box itself to improve contrast
                                *   ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○
                                * ○ ● ● ● ● ● ● ● ● ● ● ● ● ● ● ○
                                * ○ ● ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ● ○
                                * ○ ● ○                     ○ ● ○
                                * ○ ● ○                     ○ ● ○
                                * ○ ● ○                     ○ ● ○
                                * ○ ● ○                     ○ ● ○
                                * ○ ● ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ● ○
                                * ○ ● ● ● ● ● ● ● ● ● ● ● ● ● ● ○
                                *   ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ 
                            */

                            // draw horizontal lines
                            for (int x = x1; x < x2; x++)
                            {
                                if (x >= 0 && x <= result_x_max)
                                {
                                    if (y1 >= 0 && y1 <= result_y_max)
                                    {
                                        // top border
                                        if (y1 > 0) target[y1 - 1, x] = false;   // outer (lowered pins)
                                        if (y1 < result_y_max && x > x1 && x < x2) target[y1 + 1, x] = false;   // inner (lowered pins)
                                        target[y1, x] = (!_dashed || (x - x1) % 3 != 2) ? true : false;        // raised pins, except every 3rd in dashed mode

                                    }
                                    if (y2 >= 0 && y2 <= result_y_max)
                                    {
                                        // bottom border
                                        if (y2 > 0 && x > x1 && x < x2) target[y2 - 1, x] = false;   // inner (lowered pins)
                                        if (y2 < result_y_max) target[y2 + 1, x] = false;   // outer (lowered pins)
                                        target[y2, x] = (!_dashed || (x - x1) % 3 != 2) ? true : false;        // raised pins, except every 3rd in dashed mode
                                    }
                                }
                            }

                            // draw vertical lines
                            for (int y = y1; y < y2 + 1; y++)
                            {
                                if (y >= 0 && y <= result_y_max)
                                {
                                    if (x1 >= 0 && x1 <= result_x_max)
                                    {
                                        // left border 
                                        if (x1 > 0) target[y, x1 - 1] = false;  // outer (lowered pins)
                                        if (x1 < result_x_max && y > y1 && y < y2) target[y, x1 + 1] = false;  // inner (lowered pins)
                                        target[y, x1] = (!_dashed || (y - y1) % 3 != 2) ? true : false;        // raised pins, except every 3rd in dashed mode
                                    }
                                    if (x2 >= 0 && x2 <= result_x_max)
                                    {
                                        // right border 
                                        if (x2 > 0 && y > y1 && y < y2) target[y, x2 - 1] = false;  // inner (lowered pins)
                                        if (x2 < result_x_max) target[y, x2 + 1] = false;  // outer (lowered pins)
                                        target[y, x2] = (!_dashed || (y - y1) % 3 != 2) ? true : false;        // raised pins, except every 3rd in dashed mode
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return target;
        }

        virtual protected bool[,] paintPolygonPointMarker(IViewBoxModel view, bool[,] result)
        {
            // point;
            if (result != null
                //&& CurrentBoundingBox.X >= 0 && CurrentBoundingBox.Y >= 0
                && CurrentPoint.X > 0 && CurrentPoint.Y > 0
                )
            {
                int result_x_max = result.GetLength(1);
                int result_y_max = result.GetLength(0);

                if (result_x_max * result_y_max > 0)
                {

                    if (view is IZoomable && view is IPannable)
                    {
                        //Rectangle pageBounds = getPageBounds();
                        double zoom = ((IZoomable)view).GetZoom();
                        int xOffset = ((IPannable)view).GetXOffset();
                        int yOffset = ((IPannable)view).GetYOffset();
                        // coords of the shapes bounding box, relative to the whole captured image
                        // Rectangle relbBox = new Rectangle(CurrentBoundingBox.X - pageBounds.X, CurrentBoundingBox.Y - pageBounds.Y, CurrentBoundingBox.Width, CurrentBoundingBox.Height);

                        Point relPoint = new Point(
                            CurrentPoint.X
                            //- pageBounds.X
                            , CurrentPoint.Y
                            //- pageBounds.Y
                            );

                        // converted to braille output coords, as shown in the original view (with zoom factor and panning position applied)
                        Point out_bBox = new Point(
                            (int)Math.Round((relPoint.X * zoom) + xOffset - 1), // x
                            (int)Math.Round((relPoint.Y * zoom) + yOffset - 1)); // y

                        int x = out_bBox.X;
                        int y = out_bBox.Y;

                        /*
                         * DoRenderBoundingBox
                         *     true             |  false
                         *      ○               |  
                         *    ○ ● ○             |    ○
                         *  ○ ● + ● ○           |  ○ + ○
                         *    ○ ● ○             |    ○
                         *      ○               |
                         */

                        // cross
                        setSaveDot(x, y, ref result, _doRenderBoundingBox, result_x_max, result_y_max);
                        setSaveDot(x - 1, y, ref result, _doRenderBoundingBox, result_x_max, result_y_max);
                        setSaveDot(x + 1, y, ref result, _doRenderBoundingBox, result_x_max, result_y_max);
                        setSaveDot(x, y - 1, ref result, _doRenderBoundingBox, result_x_max, result_y_max);
                        setSaveDot(x, y + 1, ref result, _doRenderBoundingBox, result_x_max, result_y_max);

                        if (_doRenderBoundingBox)
                        {
                            // cross spacing
                            setSaveDot(x - 2, y, ref result, false, result_x_max, result_y_max);
                            setSaveDot(x + 2, y, ref result, false, result_x_max, result_y_max);

                            setSaveDot(x - 1, y - 1, ref result, false, result_x_max, result_y_max);
                            setSaveDot(x - 1, y + 1, ref result, false, result_x_max, result_y_max);

                            setSaveDot(x, y - 2, ref result, false, result_x_max, result_y_max);
                            setSaveDot(x, y + 2, ref result, false, result_x_max, result_y_max);

                            setSaveDot(x + 1, y - 1, ref result, false, result_x_max, result_y_max);
                            setSaveDot(x + 1, y + 1, ref result, false, result_x_max, result_y_max);
                        }


                        // paint dotted line
                        if (PaintControlPointReferenceLine) // && _doRenderBoundingBox
                        {
                            var relatedPoint = RelatedPolyPoint;
                            if (!relatedPoint.IsEmpty && CurrentPolyPoint != null)
                            {
                                Point relP = CurrentPolyPoint.TransformPointCoordinatesIntoScreenCoordinates(relatedPoint);
                                if (relP.X > 0)
                                {

                                    int x2 = (int)Math.Round((relP.X * zoom) + xOffset - 1); // x
                                    int y2 = (int)Math.Round((relP.Y * zoom) + yOffset - 1); // y


                                    // do Bresenham painting
                                    var success = doBresenham(x, y, x2, y2, ref result);
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Sets a binary dot in a matrix in a save way.
        /// </summary>
        /// <param name="x">The horizontal position.</param>
        /// <param name="y">The vertical position.</param>
        /// <param name="m">The matrix to manipulate.</param>
        /// <param name="value">value of the dot.</param>
        /// <param name="max_x">The maximum horizontal dimension.</param>
        /// <param name="max_y">The maximum vertical dimension.</param>
        /// <returns><c>true</c> if the dot could be set.</returns>
        protected static bool setSaveDot(int x, int y, ref bool[,] m, bool value, int max_x = 0, int max_y = 0)
        {
            if (m != null && x > -1 && y > -1)
            {
                try
                {
                    if (max_x < 1) max_x = m.GetLength(1);
                    if (max_y < 1) max_y = m.GetLength(0);

                    if (max_x * max_y > 0)
                    {
                        if (x < max_x && y < max_y)
                        {
                            m[y, x] = value;
                            return true;
                        }
                    }
                }
                catch { }
            }

            return false;
        }

        virtual protected Rectangle getPageBounds()
        {
            Rectangle pageBounds = new Rectangle();
            if (WindowManager.Instance != null)
            {
                ScreenObserver obs = WindowManager.Instance.DrawAppModel.ScreenObserver;
                if (obs != null)
                {
                    Object scPos = obs.ScreenPos;
                    if (scPos != null & scPos is Rectangle)
                    {
                        pageBounds.X = ((Rectangle)scPos).X;
                        pageBounds.Y = ((Rectangle)scPos).Y;
                        pageBounds.Width = ((Rectangle)scPos).Width;
                        pageBounds.Height = ((Rectangle)scPos).Height;
                    }
                }
            }

            return pageBounds;
        }

        
        /// <summary>
        /// Does an line interpolation between two points using the Bresenham algorithm.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <param name="x2">The x2.</param>
        /// <param name="y2">The y2.</param>
        /// <param name="m">The matrix to draw in.</param>
        public static bool doBresenham(int x, int y, int x2, int y2, ref bool[,] m)
        {
            int w = x2 - x;
            int h = y2 - y;
            int dx1 = 0, dy1 = 0, dx2 = 0, dy2 = 0;
            if (w < 0) dx1 = -1; else if (w > 0) dx1 = 1;
            if (h < 0) dy1 = -1; else if (h > 0) dy1 = 1;
            if (w < 0) dx2 = -1; else if (w > 0) dx2 = 1;
            int longest = Math.Abs(w);
            int shortest = Math.Abs(h);
            if (!(longest > shortest))
            {
                longest = Math.Abs(h);
                shortest = Math.Abs(w);
                if (h < 0) dy2 = -1; else if (h > 0) dy2 = 1;
                dx2 = 0;
            }
            int numerator = longest >> 1;
            for (int i = 0; i <= longest; i++)
            {

                // paint the point
                if (i%2 == 0 && m != null && y >= 0 && x >= 0 && m.GetLength(0) > y && m.GetLength(1) > x)
                {
                    m[y, x] = true;
                }

                numerator += shortest;
                if (!(numerator < longest))
                {
                    numerator -= longest;
                    x += dx1;
                    y += dy1;
                }
                else
                {
                    x += dx2;
                    y += dy2;
                }
            }
            return true;
        }
        
    }
}