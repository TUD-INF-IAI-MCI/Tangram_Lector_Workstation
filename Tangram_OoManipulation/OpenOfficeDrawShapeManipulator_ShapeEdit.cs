using System;
using System.Drawing;
using System.IO;
using System.Linq;
using tud.mci.tangram.audio;
using tud.mci.tangram.util;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    /// <summary>
    /// Class for manipulating OpenOffice Draw document elements 
    /// </summary>
    public partial class OpenOfficeDrawShapeManipulator : AbstractSpecializedFunctionProxyBase
    {

        #region Modification Mode Carousel

        private TimeSpan doubleKlickTolerance = new TimeSpan(0, 0, 0, 0, 300);
        private DateTime lastRequest = DateTime.Now;

        private readonly int _maxMode = Enum.GetValues(typeof(ModificationMode)).Cast<int>().Max();

        /// <summary>
        /// Rotates through shape manipulation modes.
        /// </summary>
        /// <param name="shapeRecentlyCreated">Set true if function is called because of shape creation.</param>
        public void RotateThroughModes(bool shapeRecentlyCreated = false)
        {
            if (!IsShapeSelected /*LastSelectedShape == null*/)
            {
                return;
            }
            if (LastSelectedShapePolygonPoints != null)
            {
                Mode = ModificationMode.Move;
                //playError();
                return;
            }

            DateTime timestamp = DateTime.Now;

            object val = Convert.ChangeType(Mode, Mode.GetTypeCode());
            int m = Convert.ToInt32(val);

            // mode switch
            // should not rotate to UNKOWN = 0
            if (timestamp - lastRequest > doubleKlickTolerance)
            {
                m = Math.Max(1, (++m) % (_maxMode + 1));
            }
            else
            { // double click ... rotate backwards
                m = (m - 2) % (_maxMode + 1);
                if (m <= 0) m = _maxMode + m;
            }

            lastRequest = timestamp;

            Mode = (ModificationMode)Enum.ToObject(typeof(ModificationMode), m);
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MODE SWITCH] to the new mode: " + Mode);
            comunicateModeSwitch(Mode, shapeRecentlyCreated);
        }

        private void comunicateModeSwitch(ModificationMode mode, bool noAudio = false)
        {
            String audio = getAudioFeedback(mode);
            String detail = getDetailRegionFeedback(mode);

            if (!noAudio) AudioRenderer.Instance.PlaySoundImmediately(audio);
            sentTextFeedback(detail);
        }

        #endregion

        #region Manipulation

        private double getZoomLevel()
        {
            if (Zoomable != null)
            {
                return Zoomable.GetZoomLevel();
            }

            return 1;
        }

        private int getStep()
        {
            // get the right step - depending on the zoom level?!
            //
            // you have to move at least one visible pixel
            // visible pixel depends on the resolution of the output device
            // you need at least 10 dpi for displaying Braille
            // one inch is 25.4 mm --> one dot is circa 2.54 mm in zoom level 1.0
            // need over-sampling to see every change -->  3 mm change at least
            //
            // zoom factor is relation between one pixel and one dot
            // resolution of a screen should be circa 96 dpi
            // one pixel is circa 0.265 mm -->  3.78 p/mm
            // --> size in pixel / 3.78 = size in mm

            double zoom = getZoomLevel();
            //double change = (300 / zoom) / 3.78;
            double change = (150 / zoom) / 3.78;
            return (int)Math.Round(change, MidpointRounding.AwayFromZero);
        }

        private int getSizeSteps() { return getStep() * 2; }

        private int getLargeDegree() { return 1500; }

        private int getSmallDegree() { return 100; }

        private bool emptyShapeForManipulationError()
        {
            if (!IsShapeSelected /*LastSelectedShape == null*/)
            {
                return true;
            }
            return false;
        }

        #region Mode Based Functions

        private bool handleUP()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " up");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    //Mode = ModificationMode.Move;
                    //moveShapeVertical(-getStep());
                    return false;
                case ModificationMode.Move:
                    moveShapeVertical(-getStep());
                    break;
                case ModificationMode.Scale:
                    scaleHeight(getSizeSteps());
                    break;
                case ModificationMode.Rotate:
                    rotateLeft(-getSmallDegree());
                    break;
                case ModificationMode.Fill:
                    changeFillStyle(-1);
                    return true;
                // FIXME: getDetailregionFeedback is not correct --> detail region content is set in changeFillStyle function
                case ModificationMode.Line:
                    changeLineWidth(50);
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleDOWN()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " down");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    //Mode = ModificationMode.Move;
                    //moveShapeVertical(getStep());
                    return false;
                case ModificationMode.Move:
                    moveShapeVertical(getStep());
                    break;
                case ModificationMode.Scale:
                    scaleHeight(-getSizeSteps());
                    break;
                case ModificationMode.Rotate:
                    rotateLeft(getSmallDegree());
                    break;
                case ModificationMode.Fill:
                    changeFillStyle(1);
                    return true; // getDetailregionFeedback is not correct --> detail region content is set in changeFillStyle function
                case ModificationMode.Line:
                    changeLineWidth(-50);
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleLEFT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " left");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    //Mode = ModificationMode.Move;
                    //moveShapeHorizontal(-getStep());
                    return false;
                case ModificationMode.Move:
                    moveShapeHorizontal(-getStep());
                    break;
                case ModificationMode.Scale:
                    scaleWidth(-getSizeSteps());
                    break;
                case ModificationMode.Rotate:
                    rotateLeft(getLargeDegree());
                    break;
                case ModificationMode.Fill:
                    changeFillStyle(-1);
                    return true; // getDetailregionFeedback is not correct --> detail region content is set in changeFillStyle function
                case ModificationMode.Line:
                    changeLineStyle(-1);
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleRIGHT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " right");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    //Mode = ModificationMode.Move;
                    //moveShapeHorizontal(getStep());
                    return false;
                case ModificationMode.Move:
                    moveShapeHorizontal(getStep());
                    break;
                case ModificationMode.Scale:
                    scaleWidth(getSizeSteps());
                    break;
                case ModificationMode.Rotate:
                    rotateRight(getLargeDegree());
                    break;
                case ModificationMode.Fill:
                    changeFillStyle(1);
                    return true; // getDetailregionFeedback is not correct --> detail region content is set in changeFillStyle function
                case ModificationMode.Line:
                    changeLineStyle(1);
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleUP_RIGHT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " up right");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    return false;
                case ModificationMode.Move:
                    int step = getStep();
                    moveShape(step, -step);
                    break;
                case ModificationMode.Scale:
                    break;
                case ModificationMode.Rotate:
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleUP_LEFT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " up left");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    return false;
                case ModificationMode.Move:
                    int step = getStep();
                    moveShape(-step, -step);
                    break;
                case ModificationMode.Scale:
                    break;
                case ModificationMode.Rotate:
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleDOWN_RIGHT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " down right");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    return false;
                case ModificationMode.Move:
                    int step = getStep();
                    moveShape(step, step);
                    break;
                case ModificationMode.Scale:
                    break;
                case ModificationMode.Rotate:
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        private bool handleDOWN_LEFT()
        {
            if (emptyShapeForManipulationError()) return false;
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[MANIPULATION] " + Mode.ToString() + " down left");
            switch (Mode)
            {
                case ModificationMode.Unknown:
                    return false;
                case ModificationMode.Move:
                    int step = getStep();
                    moveShape(-step, step);
                    break;
                case ModificationMode.Scale:
                    break;
                case ModificationMode.Rotate:
                    break;
                default:
                    break;
            }
            sentTextFeedback(getDetailRegionFeedback(Mode));
            return true;
        }

        #endregion

        # region Fillmodes

        /// <summary>        
        /// Initializes the patters for filling forms.
        /// Patters are contained in OpenOffice extension TangramToolbar.oxt.
        /// Path C:\Users\voegler\AppData\Roaming\OpenOffice\4\user\uno_packages\cache\uno_packages\xxx\TangramToolbar.oxt\bitmap-pattern
        private void initializePatters()
        {
            string[] arrPatterns;
            try
            {
                string patternName = "";
                PatterDic.Add(LL.GetTrans("tangram.oomanipulation.current_fill_pattern.no_pattern"), "no_pattern");
                PatterDic.Add(LL.GetTrans("tangram.oomanipulation.current_fill_pattern.white_pattern"), "white_pattern");
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                Logger.Instance.Log(LogPriority.DEBUG, "Pattern loader", "Application folder: " + appData);
#if LIBRE
                var foundDirs = Directory.GetDirectories(appData + "\\LibreOffice", "TangramToolbar_LO*.oxt", SearchOption.AllDirectories);
#else
                var foundDirs = Directory.GetDirectories(appData + "\\OpenOffice", "TangramToolbar_OO*.oxt", SearchOption.AllDirectories);
#endif

                Logger.Instance.Log(LogPriority.DEBUG, "Pattern loader", "directories inside application folder: " + (foundDirs != null ? foundDirs.Length.ToString() : "null"));
                if (foundDirs != null && foundDirs.Length > 0)
                {
                    var openOfficePath = foundDirs[foundDirs.Length - 1];
                    Logger.Instance.Log(LogPriority.DEBUG, "Pattern loader", "used directory: " + openOfficePath);
                    arrPatterns = Directory.GetFiles(openOfficePath, "*.png", SearchOption.AllDirectories);
                    //remove files named name_TS.png
                    foreach (string pattern in arrPatterns)
                    {
                        if (!pattern.ToLower().EndsWith("_ts.png"))  // only files without _ts.png
                        {
                            //add entry
                            patternName = Path.GetFileNameWithoutExtension(pattern);
                            string name = LL.GetTrans("tangram.oomanipulation.current_fill_pattern." + patternName);
                            if (String.IsNullOrWhiteSpace(name)) name = pattern;
                            while (PatterDic.ContainsKey(name))
                            {
                                name += "*";
                            }
                            PatterDic.Add(name, pattern);
                            Logger.Instance.Log(LogPriority.DEBUG, "Pattern loader", "add pattern " + name);
                        }
                    }
                }
                if (PatterDic.Count == 0)
                {
                    AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.no_Tangram_extension_found"));
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Unhandled exception occurred while loading pattern from extension TangramToolbar.oxt", ex);
                Logger.Instance.Log(LogPriority.DEBUG, "Pattern loader", "[ERROR] can't load patterns: " + ex);
            }
        }

        private void changeFillStyle(int p)
        {
            if (IsShapeSelected && LastSelectedShape != null && LastSelectedShape.IsValid() && PatterDic.Count > 0)
            {
                try
                {
                    string bitmapName = LastSelectedShape.GetBackgroundBitmapName();

                    fillStyleNum += p;

                    if (fillStyleNum < 0)
                    {
                        fillStyleNum = PatterDic.Count - 1;
                    }
                    if (fillStyleNum > PatterDic.Count - 1)
                    {
                        fillStyleNum = 0;
                    }

                    string pattern = PatterDic.Keys.ElementAt(fillStyleNum); //LL.GetTrans("tangram.oomanipulation.current_fill_pattern." + bitmapName);
                    bitmapName = PatterDic.Values.ElementAt(fillStyleNum);
                    if (bitmapName == "no_pattern")
                    {
                        LastSelectedShape.FillStyle = tud.mci.tangram.util.FillStyle.NONE;
                    }
                    else if (bitmapName == "white_pattern")
                    {
                        LastSelectedShape.FillStyle = tud.mci.tangram.util.FillStyle.SOLID;
                        LastSelectedShape.FillColor = OoUtils.ConvertToColorInt(System.Drawing.Color.White);
                    }
                    else
                    {
                        LastSelectedShape.FillStyle = tud.mci.tangram.util.FillStyle.BITMAP;
                        LastSelectedShape.SetBackgroundBitmap(PatterDic[pattern], tud.mci.tangram.util.BitmapMode.REPEAT, bitmapName);
                    }

                    if (pattern.Contains("_"))
                    {
                        string[] patternSplit = pattern.Split(new Char[] { '_' });
                        pattern = patternSplit[1];
                    }
                    if (String.IsNullOrWhiteSpace(pattern)) pattern = Path.GetFileNameWithoutExtension(bitmapName);
                    AudioRenderer.Instance.PlaySoundImmediately(pattern);
                    sentTextFeedback(LL.GetTrans("tangram.oomanipulation.current_fill_pattern", pattern));
                }
                catch (Exception ex)
                {

                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("changeFillStyle cannot changed some value are null or empty");
            }
        }

        #endregion

        #region Line Styles

        private void changeLineStyle(int p)
        {
            if (IsShapeSelected /*LastSelectedShape != null*/)
            {
                string lineStyleName = "";

                if (lineStyleNum <= linestyleNames.Length - 1 && lineStyleNum >= 0)
                {
                    lineStyleNum += p;
                }
                if (lineStyleNum < 0)
                {
                    lineStyleNum = linestyleNames.Length - 1;
                }
                else if (lineStyleNum > linestyleNames.Length - 1)
                {
                    lineStyleNum = 0;
                }
                lineStyleName = linestyleNames[lineStyleNum];


                //if (LastSelectedShape.Children.Count > 0) // group object
                //{
                //    AudioRenderer.Instance.PlaySoundImmediately("Ändern der Linie nicht möglich");
                //    return;
                //}

                switch (lineStyleName)
                {
                    case "solid":
                        solidStyle();
                        AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.linestyle.solid"));
                        break;
                    case "dashed_line":
                        dashedStyle();
                        AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.linestyle.dashed"));
                        break;
                    case "dotted_line":
                        dottedStyle();
                        AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.oomanipulation.linestyle.dotted"));
                        break;
                    default:
                        break;
                }
            }
        }

        private string getCurrentLineStyle()
        {
            if (IsShapeSelected && LastSelectedShape != null)
            {
                string lineStyle = LastSelectedShape.GetLineStyleName();
                switch (lineStyle)
                {
                    case "solid":
                        return LL.GetTrans("tangram.oomanipulation.linestyle.solid");
                    case "dashed_line":
                        return LL.GetTrans("tangram.oomanipulation.linestyle.dashed");
                    case "dotted_line":
                        return LL.GetTrans("tangram.oomanipulation.linestyle.dotted");
                }
            }
            return "";
        }


        private void changeLineWidth(int p)
        {
            // test whether style dotted if yes update dashStyle
            if (LastSelectedShape != null)
            {

                short dots = 0;
                short dashes = 0;
                int dotLen = 0;
                int dashLen = 0;
                int distance = 0;
                int oldLineWidth = LastSelectedShape.LineWidth;
                int newLinewidth = Math.Max(0, oldLineWidth + p);
                bool error = false;
                if (oldLineWidth == newLinewidth)
                {
                    error = true;
                }
                else
                {
                    tud.mci.tangram.util.DashStyle dashstyles = tud.mci.tangram.util.DashStyle.RECT;
                    LastSelectedShape.GetLineDash(out dashstyles, out dots, out dotLen, out dashes, out dashLen, out distance);
                    LastSelectedShape.LineWidth = LastSelectedShape.LineWidth + p;
                    if (dots > 0 && LastSelectedShape.LineStyle != tud.mci.tangram.util.LineStyle.SOLID)
                    {
                        dottedStyle();
                    }
                }

                float lineWidth = (float)LastSelectedShape.LineWidth;
                string message = LL.GetTrans("tangram.oomanipulation.mm", ((float)(lineWidth / 100)).ToString());
                if (error)
                {
                    AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.End);
                    AudioRenderer.Instance.PlaySound(message);
                }
                else
                {
                    AudioRenderer.Instance.PlaySoundImmediately(message);
                }
            }
        }

        private void solidStyle()
        {
            if (LastSelectedShape != null)
            {
                LastSelectedShape.LineStyle = tud.mci.tangram.util.LineStyle.SOLID;
            }
        }

        private void dashedStyle()
        {
            if (LastSelectedShape != null)
            {
                LastSelectedShape.LineStyle = tud.mci.tangram.util.LineStyle.DASH;
                LastSelectedShape.SetLineDash(tud.mci.tangram.util.DashStyle.ROUND, 0, 0, 1, 1000, 500);
            }
        }

        private void dottedStyle()
        {
            if (LastSelectedShape != null)
            {
                LastSelectedShape.LineStyle = tud.mci.tangram.util.LineStyle.DASH;
                int linewidth = LastSelectedShape.LineWidth;
                LastSelectedShape.SetLineDash(tud.mci.tangram.util.DashStyle.ROUND, 1, linewidth, 0, 0, 500);
            }
        }

        #endregion

        #region Moving

        private void moveShapeHorizontal(int steps)
        {
            moveShape(steps, 0);
        }

        private void moveShapeVertical(int steps)
        {
            moveShape(0, steps);
        }

        private void moveShape(int horizontalSteps, int verticalSteps)
        {
            bool success = false;
            if (LastSelectedShapePolygonPoints != null)
            {
                success = movePolygonPoint(horizontalSteps, verticalSteps);
            }
            else if (IsShapeSelected && LastSelectedShape != null)
            {
                var pos = LastSelectedShape.Position;
                if (!pos.IsEmpty)
                {
                    pos.X += horizontalSteps;
                    pos.Y += verticalSteps;
                    LastSelectedShape.Position = pos;
                    success = true;
                }
            }

            if (success) { playEdit(); }
            else { playError(); }
        }

        private bool movePolygonPoint(int horizontalSteps, int verticalSteps)
        {
            bool success = false;
            int i;
            var point = LastSelectedShapePolygonPoints.Current(out i);

            // if an e.g. adjustment value for custom shapes is set.
            if (point.Flag == util.PolygonFlags.CUSTOM && point.Value != null)
            {
                if (point.Value is double) // point is position
                {
                    double val = (double)point.Value;
                    double newVal = val + (horizontalSteps == 0 ? verticalSteps : horizontalSteps);
                    point.Value = newVal;
                }
                else if (point.Value is int) // point is radius
                {
                    int rad = (int)point.Value;
                    int change = (horizontalSteps == 0 ? getLargeDegree() : getSmallDegree()) / 100;

                    bool invert = horizontalSteps < 0 || verticalSteps > 0;

                    int newRad = rad + (invert ? -change : change);
                    point.Value = newRad;
                }

            }

            point.X += horizontalSteps;
            point.Y += verticalSteps;

            // Special treatment for closed forms:
            // first and last point have to be the same!
            // FIXME: check if this works for poly polygons !
            if (!point.Flag.HasFlag(tud.mci.tangram.util.PolygonFlags.CUSTOM) && (i == 0 || i == LastSelectedShapePolygonPoints.Count - 1)
                && LastSelectedShapePolygonPoints.IsClosed() // is closed polygon
                )
            {
                success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, 0, false);
                success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, LastSelectedShapePolygonPoints.Count - 1, false);
                success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
            }
            else
            {
                success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, i, true);
            }
            return success;
        }

        #endregion

        #region Scaling

        private static int _minSize = 100;
        /// <summary>
        /// Gets the minimum length size of a shape.
        /// This value can be set by the app.config key "Tangram_MinimumShapeSize".
        /// </summary>
        /// <value>
        /// The minimum size of the shape in 100th/mm.
        /// </value>
        public static int MinimumShapeSize
        {
            get { return _minSize; }
            private set { _minSize = Math.Max(1, value); }
        }

        private static double _squareLockDistFactor = 0.5;
        /// <summary>
        /// Gets the factor of the applied step size for changing to identify the intention to make both dimensions equal (square).
        /// </summary>
        /// <value>
        /// The square lock distance factor.
        /// </value>
        public static double SquareLockDistanceFactor
        {
            get { return OpenOfficeDrawShapeManipulator._squareLockDistFactor; }
            private set { OpenOfficeDrawShapeManipulator._squareLockDistFactor = value; }
        }

        private void scaleWidth(int steps) { scale(steps, 0); }

        private void scaleHeight(int steps) { scale(0, steps); }

        private void scale(int widthSteps, int heightSteps)
        {
            if (IsShapeSelected && LastSelectedShape != null)
            {
                // switch width and height if shape is rotated
                var rotation = LastSelectedShape.Rotation;
                if ((rotation >= 4500 && rotation <= 13500) ||
                    (rotation >= 22500 && rotation <= 31500))
                {
                    if (!PolygonHelper.IsFreeform(LastSelectedShape.DomShape))
                    {
                        int tmp = widthSteps;
                        widthSteps = heightSteps;
                        heightSteps = tmp;
                    }
                }

                bool error = false;
                // old Bounds
                var bounds = LastSelectedShape.Bounds;
                var size = LastSelectedShape.Size;


                bool wdch = Math.Abs(widthSteps) > Math.Abs(heightSteps) ? true : false; // hint if width or height was changed
                int step = wdch ? widthSteps : heightSteps;
                int squarTolerence = (int)(step * SquareLockDistanceFactor);


                size.Width += widthSteps;
                size.Height += heightSteps;

                // check if new size square like or not
                if (step != 0 && Math.Abs(size.Width - size.Height) <= Math.Abs(squarTolerence)) // difference between both dimensions is less then tolerance 
                {
                    // check which part should be adapted
                    if (wdch) // with was changed so adapt
                    {
                        size.Width = size.Height;
                    }
                    else // height was changed
                    {
                        size.Height = size.Width;
                    }
                }

                if (size.Height < MinimumShapeSize)
                {
                    size.Height = MinimumShapeSize;
                    heightSteps = MinimumShapeSize - LastSelectedShape.Size.Height;
                    error = true;
                }

                if (size.Width < MinimumShapeSize)
                {
                    size.Width = MinimumShapeSize;
                    widthSteps = MinimumShapeSize - LastSelectedShape.Size.Width;
                    error = true;
                }

                // reposition
                Point pos = new Point();
                if (bounds.IsEmpty)
                {
                    pos = LastSelectedShape.Position;

                    pos.X -= (widthSteps / 2);
                    pos.Y -= (heightSteps / 2);

                    if (pos.X < 0) pos.X = 0;
                    if (pos.Y < 0) pos.Y = 0;

                    LastSelectedShape.Size = size;
                    LastSelectedShape.Position = pos;
                }
                else
                {
                    pos.X = bounds.Left;
                    pos.Y = bounds.Top;

                    Point oldCenter = new Point(
                        (int)(bounds.Left + bounds.Width / 2.0),
                        (int)(bounds.Top + bounds.Height / 2.0)
                        );

                    LastSelectedShape.Size = size; // change size to new one

                    int relX = 0, relY = 0;
                    var newBounds = LastSelectedShape.Bounds;

                    if (newBounds.IsEmpty)
                    {
                        // TODO: should not happen
                    }
                    else
                    {
                        relX = (int)(oldCenter.X - (newBounds.Left + newBounds.Width * 0.5));
                        relY = (int)(oldCenter.Y - (newBounds.Top + newBounds.Height * 0.5));
                    }

                    Point p = LastSelectedShape.Position;
                    p.X += relX;
                    p.Y += relY;
                    LastSelectedShape.Position = p;

                }

                if (error) playError();
                else playEdit();
            }
        }

        #endregion

        #region Rotating

        private void rotateLeft(int degres) { rotate(degres); }

        private void rotateRight(int degres) { rotate(-degres); }

        private void rotate(int degres)
        {
            if (IsShapeSelected && LastSelectedShape != null)
            {
                int rotation = LastSelectedShape.Rotation;
                LastSelectedShape.Rotation = (rotation + degres) - (rotation % 100);
                playEdit();
                int rot = (LastSelectedShape.Rotation / 100);
                String detail = LL.GetTrans("tangram.oomanipulation.rotated", rot.ToString("0."));
                sendDetailInfo(detail);
                play(LL.GetTrans("tangram.oomanipulation.degrees", rot.ToString("0.")));
            }
        }

        #endregion

        #region Delete

        private void deleteSelectedObject()
        {
            bool success = false;
            if (LastSelectedShapePolygonPoints != null)
                success = deleteSelectedPolyPoint();
            else success = deleteSelectedShape();
        }

        private bool deleteSelectedShape()
        {
            bool success = false;

            // FIXME: show MessageBox
            var _shape = LastSelectedShape;
            if (_shape != null)
            {
                String n = String.Empty;
                String t = n;

                if (_shape != null) // kill the shape without validation
                {
                    n = _shape.Name;
                    t = _shape.Title;

                    success = _shape.Delete();

                    if (success)
                    {
                        LastSelectedShape = null;

                        var name = LL.GetTrans("tangram.oomanipulation.delete.element", n + " " + t);
                        var audio = LL.GetTrans("tangram.oomanipulation.element_speaker.delete.element", n);
                        sendDetailInfo(name);
                        sentAudioFeedback(audio);
                    }
                    else
                    {
                        String msg = LL.GetTrans("tangram.oomanipulation.command.not_successfull", LL.GetTrans("tangram.oomanipulation.command.DELETE"));

                        if (AudioRenderer.Instance != null)
                        {
                            AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Critical);
                            AudioRenderer.Instance.PlaySound(msg);
                        }
                        sentTextNotification(msg);
                    }
                }
            }

            return success;
        }

        private bool deleteSelectedPolyPoint()
        {
            bool success = false;

            if (LastSelectedShapePolygonPoints != null)
            {
                success = LastSelectedShapePolygonPoints.IsBezier() ? deleteBezierPoint() : deletePolygonPoint();

                if (success) playEdit();
                else playError();
            }
            return success;
        }

        private bool deletePolygonPoint()
        {
            bool success = false;
            int i = -1;
            var _p = LastSelectedShapePolygonPoints.Current(out i);
            if (i >= 0)
            {
                // Special treatment for closed forms:
                // first and last point have to be the same!
                // FIXME: check if this works for poly polygons !
                if ((i == 0 || i == LastSelectedShapePolygonPoints.Count - 1)
                    && LastSelectedShapePolygonPoints.IsClosed() // is closed polygon
                    )
                {
                    if (i == 0) // first point
                    {
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(0, false);
                        // prepare the new end point!
                        var point = LastSelectedShapePolygonPoints.First();
                        success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, LastSelectedShapePolygonPoints.Count - 1, false);
                        success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
                    }
                    else
                    { // Last
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i, false);
                        // prepare the new start point!
                        var point = LastSelectedShapePolygonPoints.Last();
                        success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, 0, false);
                        success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
                    }
                }
                else
                {
                    success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i, true);
                }
            }
            return success;
        }

        private bool deleteBezierPoint()
        {
            bool success = false;
            int i = -1;
            var _p = LastSelectedShapePolygonPoints.Current(out i);
            if (i >= 0)
            {
                // Special treatment for closed forms:
                // first and last point have to be the same!
                // FIXME: check if this works for poly polygons !
                if ((i == 0 || i == 1 || i == LastSelectedShapePolygonPoints.Count - 1 || i == LastSelectedShapePolygonPoints.Count - 2)
                    && LastSelectedShapePolygonPoints.IsClosed() // is closed polygon
                    )
                {
                    if (i == 0 || i == 1) // first point or its control
                    {
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(0, false); // point
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(0, false); // control
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(0, false); // control
                        // prepare the new end point!
                        var point = LastSelectedShapePolygonPoints.First();
                        success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, LastSelectedShapePolygonPoints.Count - 1, false);
                        // TODO: adapt control ?
                        success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
                    }
                    else
                    {   // Last
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 2, false); // control
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 2, false); // control
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 2, false); // point

                        // prepare the new start point!
                        var point = LastSelectedShapePolygonPoints.Last();
                        success = LastSelectedShapePolygonPoints.UpdatePolyPointDescriptor(point, 0, false);
                        // TODO: adapt control ?
                        success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
                    }
                }
                else
                {
                    // check point type - delete only works for edge points!
                    if (_p.Flag.HasFlag(util.PolygonFlags.NORMAL))
                    {
                        // TODO: what to do if structure is not ... c | p | c ...
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 1, false); // control
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 1, false); // point
                        success = LastSelectedShapePolygonPoints.RemovePolygonPoints(i - 1, false); // control
                        success = LastSelectedShapePolygonPoints.WritePointsToPolygon();
                    }
                    // TODO: delete control points?
                }
            }
            return success;
        }

        #endregion

        #region Order / Z-Index

        /// <summary>
        /// Brings the modifying shape one step higher.
        /// </summary>
        /// <returns><c>true</c> if the shape could brought one level higher; otherwise <c>false</c>.</returns>
        private bool bringToFront()
        {
            if (IsShapeSelected && LastSelectedShape != null)
            {
                int currZIndex = LastSelectedShape.ZOrder;
                LastSelectedShape.ZOrder = currZIndex + 1;
                return LastSelectedShape.ZOrder > currZIndex;
            }

            return false;
        }

        /// <summary>
        /// Sent the modifying shape one step deeper.
        /// </summary>
        /// <returns><c>true</c> if the shape could sent one level deeper; otherwise <c>false</c>.</returns>
        private bool sentToBackground()
        {
            if (IsShapeSelected && LastSelectedShape != null)
            {
                int currZIndex = LastSelectedShape.ZOrder;
                LastSelectedShape.ZOrder = currZIndex - 1;
                return LastSelectedShape.ZOrder < currZIndex;
            }
            return false;
        }

        #endregion

        #endregion

        #region audio and detail region feedback

        private String getAudioFeedback(ModificationMode mode)
        {
            String audio = "";//"Element ";

            switch (mode)
            {
                case ModificationMode.Unknown:
                    break;
                case ModificationMode.Move:
                    audio += LL.GetTrans("tangram.oomanipulation.manipulation.move");
                    if (IsShapeSelected && LastSelectedShape != null && LastSelectedShape.Position != null)
                    {
                        audio += ", " + LL.GetTrans("tangram.oomanipulation.current") + ": "
                            + LL.GetTrans("tangram.oomanipulation.manipulation.move.status.audio"
                            , ((float)((float)LastSelectedShape.Position.X / 1000)).ToString("0.#")
                            , ((float)((float)LastSelectedShape.Position.Y / 1000)).ToString("0.#"));
                    }
                    break;
                case ModificationMode.Scale:
                    audio += LL.GetTrans("tangram.oomanipulation.manipulation.scale");
                    if (IsShapeSelected && LastSelectedShape != null && LastSelectedShape.Size != null)
                    {
                        audio += ", " + LL.GetTrans("tangram.oomanipulation.current") + ": "
                            + LL.GetTrans("tangram.oomanipulation.manipulation.scale.status.audio"
                            , ((float)((float)LastSelectedShape.Size.Height / 1000)).ToString("0.#")
                            , ((float)((float)LastSelectedShape.Size.Width / 1000)).ToString("0.#"));
                    }
                    break;
                case ModificationMode.Rotate:
                    audio += LL.GetTrans("tangram.oomanipulation.manipulation.rotate");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        audio += ", " + LL.GetTrans("tangram.oomanipulation.current") + ": "
                            + LL.GetTrans("tangram.oomanipulation.manipulation.rotate.status.audio"
                            , (LastSelectedShape.Rotation / 100).ToString("0."));
                    }
                    break;
                case ModificationMode.Fill:
                    audio += LL.GetTrans("tangram.oomanipulation.manipulation.filling.audio");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        if (LastSelectedShape.GetProperty("FillBitmap") == null) audio += ", " + LL.GetTrans("tangram.oomanipulation.manipulation.filling.status.none");
                        else
                        {
                            audio += ", " + LL.GetTrans("tangram.oomanipulation.current") + ": "
                                + LL.GetTrans("tangram.oomanipulation.manipulation.filling.status"
                                , LastSelectedShape.GetBackgroundBitmapName());
                        }
                    }
                    break;
                case ModificationMode.Line:
                    audio += LL.GetTrans("tangram.oomanipulation.manipulation.line.audio");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        audio += ", " + LL.GetTrans("tangram.oomanipulation.current") + ": "
                            + LL.GetTrans("tangram.oomanipulation.manipulation.line.status.audio"
                            , ((float)((float)LastSelectedShape.LineWidth / 100)).ToString("0.#")
                            , getCurrentLineStyle());
                    }
                    break;
                default:
                    break;
            }
            return audio;
        }

        private String getDetailRegionFeedback(ModificationMode mode)
        {
            String detail = "";//Bearbeitung: ";

            switch (mode)
            {
                case ModificationMode.Unknown:
                    break;
                case ModificationMode.Move:

                    if (LastSelectedShapePolygonPoints != null) { return GetPointText(LastSelectedShapePolygonPoints); }

                    detail += LL.GetTrans("tangram.oomanipulation.manipulation.move");
                    if (IsShapeSelected && LastSelectedShape != null && LastSelectedShape.Position != null)
                    {
                        detail += " - " + LL.GetTrans("tangram.oomanipulation.manipulation.move.status"
                            , ((float)((float)LastSelectedShape.Position.X / 1000)).ToString("0.#")
                            , ((float)((float)LastSelectedShape.Position.Y / 1000)).ToString("0.#"));
                    }
                    break;
                case ModificationMode.Scale:
                    detail += LL.GetTrans("tangram.oomanipulation.manipulation.scale");
                    if (IsShapeSelected && LastSelectedShape != null && LastSelectedShape.Size != null)
                    {
                        detail += " - " + LL.GetTrans("tangram.oomanipulation.manipulation.scale.status"
                            , ((float)((float)LastSelectedShape.Size.Height / 1000)).ToString("0.#")
                            , ((float)((float)LastSelectedShape.Size.Width / 1000)).ToString("0.#"));
                    }
                    break;
                case ModificationMode.Rotate:
                    detail += LL.GetTrans("tangram.oomanipulation.manipulation.rotate");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        detail += " - " + LL.GetTrans("tangram.oomanipulation.degrees", (LastSelectedShape.Rotation / 100).ToString("0."));
                    }
                    break;
                case ModificationMode.Fill:
                    detail += LL.GetTrans("tangram.oomanipulation.manipulation.filling");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        if (LastSelectedShape.GetProperty("FillBitmap") != null)
                        {
                            detail += " - " + LastSelectedShape.GetBackgroundBitmapName(); // TODO: incorrect name of bitmap after changing it the first time (only returns "Bitmape 1" ect.)
                        }
                    }
                    break;
                case ModificationMode.Line:
                    detail += LL.GetTrans("tangram.oomanipulation.manipulation.line");
                    if (IsShapeSelected && LastSelectedShape != null)
                    {
                        detail += " - " + LL.GetTrans("tangram.oomanipulation.manipulation.line.status"
                            , ((float)((float)LastSelectedShape.LineWidth / 100)).ToString("0.#")
                            , getCurrentLineStyle());
                    }
                    break;
                default:
                    break;
            }
            return detail;
        }

        #endregion

    }
}