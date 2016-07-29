using System;
using System.Drawing;
using System.Threading.Tasks;
using BrailleIO.Interface;
using tud.mci.tangram.controller.observer;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    public class FocusRendererHook : IBailleIORendererHook
    {
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
        public Rectangle CurrentBoundingBox;

        /// <summary>
        /// Sets the current bounding box by shape.
        /// </summary>
        /// <param name="currentSelectedShape">The current selected shape.</param>
        public void SetCurrentBoundingBoxByShape(OoShapeObserver currentSelectedShape)
        {
            if (currentSelectedShape != null && currentSelectedShape.IsValid())
            {
                CurrentBoundingBox = currentSelectedShape.GetAbsoluteScreenBoundsByDom();
            }
            else CurrentBoundingBox = new Rectangle(0, 0, 0, 0);
        }

        /// <summary>
        /// Determine if this Hook should be active or not.
        /// </summary>
        public bool Active = false;

        /// <summary>
        /// Determine if this Hook should render the Bounding box or not.
        /// </summary>
        public bool DoRenderBoundingBox = false;

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
            if (Active && DoRenderBoundingBox && WindowManager.Instance != null && !WindowManager.Instance.IsInMinimapMode())
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
        private void doBlinkingBoundingBox(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            //draw frame as bool pins
            if (DoRenderBoundingBox && !(CurrentBoundingBox.Width * CurrentBoundingBox.Height < 1) )
            {
                if (view is IZoomable && view is IPannable)
                {
                    double zoom = ((BrailleIO.Interface.IZoomable)view).GetZoom();
                    int xOffset = ((IPannable)view).GetXOffset();
                    int yOffset = ((IPannable)view).GetYOffset();

                    //if (WindowManager.Instance != null && WindowManager.Instance.ScreenObserver != null && WindowManager.Instance.ScreenObserver.ScreenPos is System.Drawing.Rectangle)
                    //{

                        Rectangle pageBounds = getPageBounds();

                        if (pageBounds.Width > 0 && pageBounds.Height > 0)
                        {
                            // coords of the shapes bounding box, relative to the whole captured image
                            Rectangle relbBox = new Rectangle(CurrentBoundingBox.X - pageBounds.X, CurrentBoundingBox.Y - pageBounds.Y, CurrentBoundingBox.Width, CurrentBoundingBox.Height);
                            // converted to braille output coords, as shown in the original view (with zoom factor and panning position applied)
                            Rectangle out_bBox = new Rectangle(
                                (int)(relbBox.X * zoom) + xOffset - 1, // x
                                (int)(relbBox.Y * zoom) + yOffset - 1, // y
                                (int)(relbBox.Width * zoom) + 2,       // w
                                (int)(relbBox.Height * zoom) + 2);     // h

                            // check for minimal height and width
                            if (out_bBox.Width < MIN_FRAME_SIZE)
                            {
                                int oldWidth = out_bBox.Width;
                                out_bBox.Width = MIN_FRAME_SIZE;
                                int change = out_bBox.Width - oldWidth;
                                out_bBox.X -= ((change) / 2);
                            }

                            if (out_bBox.Height < MIN_FRAME_SIZE)
                            {
                                int oldHeight = out_bBox.Height;
                                out_bBox.Height = MIN_FRAME_SIZE;
                                int change = out_bBox.Height - oldHeight;

                                out_bBox.Y -= ((change) / 2);
                            }

                            if (out_bBox.Width > 0 && out_bBox.Height > 0)
                            {
                                int result_x_max = result.GetLength(1) - 1;
                                int result_y_max = result.GetLength(0) - 1;
                                // rectangle coords within matrix:
                                int x1 = Math.Max(0, out_bBox.X);                                // from bBox x (or 0)
                                int x2 = Math.Min(out_bBox.X + out_bBox.Width, result_x_max);    // to bBox x + w (or x rightmost of matrix)
                                int y1 = Math.Max(0, out_bBox.Y);                                // from bBox y (or 0)
                                int y2 = Math.Min(out_bBox.Y + out_bBox.Height, result_y_max);   // to bBox y + h (or y bottom of matrix)

                                // TODO capture on blink one frame, one blink out frame of matrix withing bounding box to let the content blink after capturing without having to modify dom any more
                                // TODO: check if shape is in visible view port --> if shape is not visible, do not blink

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

                                    bool[,] target = result;

                                    // draw horizontal lines
                                    Parallel.For(x1, x2,
                                        (x) =>
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
                                       );

                                    // draw vertical lines
                                    Parallel.For(y1, y2 + 1,
                                        (y) =>
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
                                        );

                                    result = target;
                                }
                            }
                        }
                    //}
                }
            }
        }

        Rectangle getPageBounds()
        {
            Rectangle pageBounds = new Rectangle();
            if (WindowManager.Instance != null)
            {
                ScreenObserver obs = WindowManager.Instance.ScreenObserver;
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



    }
}