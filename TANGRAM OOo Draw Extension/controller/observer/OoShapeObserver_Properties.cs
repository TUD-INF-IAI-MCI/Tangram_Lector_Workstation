using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.document;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.text;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Observes a XShape or an XShapes and his children
    /// </summary>
    public partial class OoShapeObserver : PropertiesEventForwarderBase, IUpdateable, IDisposable, IDisposingObserver
    {
        #region Property Access

        /// <summary>
        /// This property specifies the name of the font style.
        /// </summary>
        /// <value>The name of the char font.</value>
        virtual public String CharFontName
        {
            get { return getStringProperty("CharFontName"); }
            set { setStringProperty("CharFontName", value); }
        }

        /// <summary>
        /// Gets or sets the description for this object.
        /// </summary>
        /// <value>The description.</value>
        virtual public String Description
        {
            get { return getStringProperty("Description"); }
            set { setStringProperty("Description", value); }
        }

        /// <summary>
        /// this selects how a area is filled with a single bitmap.
        /// This property corresponds to the properties FillBitmapStretch and FillBitmapTile.
        /// 
        /// If set to BitmapMode::REPEAT, the property FillBitmapStretch is set 
        /// to false, and the property FillBitmapTile is set to true.
        /// If set to BitmapMode::STRETCH, the property FillBitmapStretch is set
        /// to true, and the property FillBitmapTile is set to false.
        /// If set to BitmapMode::NO_REPEAT, both properties FillBitmapStretch
        /// and FillBitmapTile are set to false.
        /// </summary>
        /// <value>The fill bitmap mode.</value>
        virtual public util.BitmapMode FillBitmapMode
        {
            get
            {
                var mode = GetProperty("FillBitmapMode");
                if (mode != null && mode is unoidl.com.sun.star.drawing.BitmapMode)
                {
                    return (util.BitmapMode)mode;
                }
                return util.BitmapMode.NO_REPEAT;
            }
            set { SetProperty("FillBitmapMode", (unoidl.com.sun.star.drawing.BitmapMode)value); }
        }

        //TODO: translate into c# color
        /// <summary>
        /// If the property FillStyle is set to FillStyle::SOLID, this is the color used. 
        /// typedef Color
        /// Defining Type
        ///     long
        /// Description
        ///     describes an RGB color value with an optional alpha channel. 
        ///     The byte order is from high to low:
        ///         alpha channel - (doesn't work)
        ///         red
        ///         green
        ///         blue
        /// </summary>
        /// <value>The color of the fill.</value>
        virtual public int FillColor
        {
            get { return getIntProperty("FillColor"); }
            set { setIntProperty("FillColor", value); }
        }

        /// <summary>
        /// Gets or sets the fill style.
        /// This enumeration selects the style the area will be filled with.  
        /// </summary>
        /// <value>The fill style.</value>
        virtual public tud.mci.tangram.util.FillStyle FillStyle
        {
            get
            {
                //TODO: check if this is to restrict
                if (OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_FILL))
                {
                    return (tud.mci.tangram.util.FillStyle)((int)GetProperty("FillStyle"));
                }
                return tud.mci.tangram.util.FillStyle.NONE;
            }
            set
            {
                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    var children = GetChilderen();
                    for (int i = 0; i < children.Count; i++)
                    {
                        OoShapeObserver child = children.ElementAt(i);

                        if (child != null)
                        {
                            child.FillStyle = value;
                        }
                    }
                }
                if(!PolygonHelper.IsLine(Shape))
                    SetProperty("FillStyle", (unoidl.com.sun.star.drawing.FillStyle)value);
            }
        }

        //TODO: check the meaning of the value
        /// <summary>
        ///  This is the transparency of the filled area.
        ///  This property is only valid if the property FillStyle is set to FillStyle::SOLID. 
        /// </summary>
        /// <value>The fill transparency.</value>
        virtual public short FillTransparence
        {
            get { return (short)getIntProperty("FillTransparence"); }
            set { setIntProperty("FillTransparence", value); }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// This is the ID of the Layer to which this Shape is attached. 
        /// </summary>
        /// <value>The layer ID.</value>
        virtual public int LayerID
        {
            get { return getIntProperty("LayerID"); }
            set { setIntProperty("LayerID", value); }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// This is the name of the Layer to which this Shape is attached.  
        /// </summary>
        /// <value>The name of the layer.</value>
        virtual public String LayerName
        {
            get { return getStringProperty("LayerName"); }
        }

        //TODO: translate into c# color
        /// <summary>
        /// This property contains the line color.
        /// typedef Color
        /// Defining Type
        ///     long
        /// Description
        ///     describes an RGB color value with an optional alpha channel. 
        ///     The byte order is from high to low:
        ///         alpha channel - (doesn't work)
        ///         red
        ///         green
        ///         blue
        /// </summary>
        /// <value>The color of the line.</value>
        virtual public int LineColor
        {
            get { return getIntProperty("LineColor"); }
            set { setIntProperty("LineColor", value); }
        }

        /// <summary>
        /// A LineDash defines a non-continuous line.
        /// </summary>
        /// <param name="dashstyle">This sets the style of this LineDash</param>
        /// <param name="dots">This is the number of dots in this LineDash</param>
        /// <param name="dotLen">This is the length of a dot</param>
        /// <param name="dashes">This is the number of dashes</param>
        /// <param name="dashLen">This is the length of a single dash</param>
        /// <param name="distance">This is the distance between the dots</param>
        virtual public void SetLineDash(tud.mci.tangram.util.DashStyle dashstyle, short dots, int dotLen, short dashes, int dashLen, int distance)
        {
            try
            {
                LineDash dash = new LineDash((unoidl.com.sun.star.drawing.DashStyle)((int)dashstyle), dots, dotLen, dashes, dashLen, distance);

                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    var children = GetChilderen();
                    for (int i = 0; i < children.Count; i++)
                    {
                        OoShapeObserver child = children.ElementAt(i);

                        if (child != null)
                        {
                            child.SetProperty("LineDash", dash);
                            child.SetLineDash(dashstyle, dots, dotLen, dashes, dashLen, distance);
                        }
                    }
                }
                SetProperty("LineDash", dash);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Can't set line dash: " + ex);
            }
        }
        /// <summary>
        /// Gets a LineDash defines a non-continuous line.
        /// </summary>
        /// <param name="dashstyle">This sets the style of this LineDash</param>
        /// <param name="dots">This is the number of dots in this LineDash</param>
        /// <param name="dotLen">This is the length of a dot</param>
        /// <param name="dashes">This is the number of dashes</param>
        /// <param name="dashLen">This is the length of a single dash</param>
        /// <param name="distance">This is the distance between the dots</param>
        /// <returns>true and false</returns>
        virtual public bool GetLineDash(out tud.mci.tangram.util.DashStyle dashstyle, out  short dots, out  int dotLen, out  short dashes, out int dashLen, out int distance)
        {
            dashstyle = tud.mci.tangram.util.DashStyle.RECT;
            dots = dashes = 0;
            dotLen = dashLen = distance = 0;
            var ld = GetProperty("LineDash");
            if (ld != null && ld is LineDash)
            {
                LineDash linedash = ((LineDash)ld);
                dashstyle = (tud.mci.tangram.util.DashStyle)((int)linedash.Style);
                dots = linedash.Dots;
                dotLen = linedash.DotLen;
                dashes = linedash.Dashes;
                dashLen = linedash.DashLen;
                distance = linedash.Distance;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Gets the internal name of the LineStyle property of the shape.
        /// </summary>
        /// <returns></returns>
        virtual public string GetLineStyleName()
        {
            if (LineStyle.Equals(tud.mci.tangram.util.LineStyle.SOLID)) return "solid";
            else if (LineStyle.Equals(tud.mci.tangram.util.LineStyle.DASH))
            {
                LineDash lineDash = GetProperty("LineDash") as LineDash;
                if (lineDash != null)
                {
                    if (lineDash.Dots == 0) return "dashed_line";
                    else if (lineDash.Dots == 1) return "dotted_line";
                }
            }
            return "";
        }

        /// <summary>
        /// specifies the appearance of the lines of a shape
        /// </summary>
        virtual public tud.mci.tangram.util.LineStyle LineStyle
        {
            get
            {
                tud.mci.tangram.util.LineStyle ls = tud.mci.tangram.util.LineStyle.UNKNOWN;

                try
                {
                    var lineStyle = GetProperty("LineStyle");
                    if (lineStyle != null) ls = (tud.mci.tangram.util.LineStyle)(int)lineStyle;
                }
                catch { }

                return ls;
            }
            set
            {
                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    var children = GetChilderen();
                    for (int i = 0; i < children.Count; i++)
                    {
                        OoShapeObserver child = children.ElementAt(i);

                        if (child != null)
                        {
                            child.LineStyle = value;
                        }
                    }
                }
                SetProperty("LineStyle", (unoidl.com.sun.star.drawing.LineStyle)((int)value));
            }
        }

        //TODO: check the meaning of the value
        /// <summary>
        /// This property contains the extent of transparency.
        /// </summary>
        /// <value>The line transparency.</value>
        virtual public int LineTransparence
        {
            get { return getIntProperty("LineTransparence"); }
            set { setIntProperty("LineTransparence", value); }
        }

        /// <summary>
        /// This property contains the width of the line in 1/100th mm. 
        /// </summary>
        /// <value>The width of the line.</value>
        virtual public int LineWidth
        {
            get
            {
                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    return GetChild(0).LineWidth;
                }
                return getIntProperty("LineWidth");
            }
            set
            {
                value = Math.Max(0, value);
                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    var children = GetChilderen();
                    for (int i = 0; i < children.Count; i++)
                    {
                        OoShapeObserver child = children.ElementAt(i);

                        if (child != null)
                        {
                            child.LineWidth = value;
                        }
                    }
                }
                setIntProperty("LineWidth", value);
            }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// With this set to true, this Shape cannot be moved interactively in the user interface.  
        /// </summary>
        /// <value><c>true</c> if [move protect]; otherwise, <c>false</c>.</value>
        virtual public bool MoveProtect
        {
            get { return getBoolProperty("MoveProtect"); }
            set { setBoolProperty("MoveProtect", value); }
        }

        private String _lastName = String.Empty;
        readonly Object _porpLock = new Object();
        /// <summary>
        /// [ OPTIONAL ]
        /// This is the name of this Shape. 
        /// Will be the id - should be unique
        /// </summary>
        /// <value>The name.</value>
        virtual public String Name
        {
            get
            {
                lock (_porpLock)
                {
                    return _lastName = getStringProperty("Name");
                }
            }
            set
            {
                lock (_porpLock)
                {
                    if (!OoUtils.ElementSupportsService(Shape, OO.Services.DRAW_SHAPE_TEXT))
                    // uncommenting still crashes, 
                    {
                        XModel docModel = (XModel)GetDocument();
                        if (docModel != null)
                        {
                            XActionLockable docModelLockable = docModel as XActionLockable;
                            //docModel.lockControllers();
                            if (docModelLockable != null) docModelLockable.addActionLock();
                            try
                            {
                                if (Shape is XNamed /* && !OoUtils.ElementSupportsService(Shape, OO.Services.DRAW_SHAPE_TEXT)*/)
                                {
                                    try
                                    {
                                        ((XNamed)Shape).setName(value);
                                    }
                                    catch (unoidl.com.sun.star.lang.DisposedException) { Dispose(); }
                                    catch
                                    {
                                        setStringProperty("Name", value);
                                    }
                                }
                                else
                                {
                                    setStringProperty("Name", value);
                                }
                                _lastName = value;

                            }
                            catch (Exception ex)
                            {
                                Logger.Instance.Log(LogPriority.DEBUG, this, ex);
                            }
                            finally
                            {
                                if (docModelLockable != null) docModelLockable.removeActionLock();
                                // docModel.unlockControllers();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// this property stores the navigation order of this shape. If this value is negative, the navigation order for this shapes page is equal to the z-order.  
        /// </summary>
        /// <value>The navigation order.</value>
        virtual public int NavigationOrder
        {
            get { return getIntProperty("NavigationOrder"); }
            set { setIntProperty("NavigationOrder", value); }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// If this is false, the Shape is not visible on printer outputs. 
        /// </summary>
        /// <value><c>true</c> if printable; otherwise, <c>false</c>.</value>
        virtual public bool Printable
        {
            get { return getBoolProperty("Printable"); }
            set { setBoolProperty("Printable", value); }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// With this set to true, this Shape may not be sized interactively in the user interface. 
        /// </summary>
        /// <value><c>true</c> if [size protect]; otherwise, <c>false</c>.</value>
        virtual public bool SizeProtect
        {
            get { return getBoolProperty("SizeProtect"); }
            set { setBoolProperty("SizeProtect", value); }
        }


        String _cachedText = String.Empty;
        DateTime _lastTextGetting = DateTime.Now;
        static readonly TimeSpan _textGettingSpan = new TimeSpan(0, 0, 2);
        volatile bool _textgetting = false;
        /// <summary>
        /// if the Shape is XTextRange (containing some text)
        /// returns the string that is included in this text range.
        /// </summary>
        /// <value>
        /// the whole string of characters of this piece of text is replaced. 
        /// All styles are removed when applying this method.
        /// </value>
        virtual public String Text
        {
            get
            {
                if (Shape is XTextRange)
                {
                    if (_textgetting || DateTime.Now - _lastTextGetting < _textGettingSpan) return _cachedText;

                    _textgetting = true;
                    try
                    {
                        XTextRange xTextRange = Shape as XTextRange;
                        TimeLimitExecutor.WaitForExecuteWithTimeLimit(50, () =>
                        {
                            if (xTextRange != null) _cachedText = xTextRange.getString();
                        }, "GetShapeText_" + this.GetHashCode());

                        _lastTextGetting = DateTime.Now;
                        return _cachedText;
                    }
                    finally
                    {
                        _textgetting = false;
                    } 
                }
                return _cachedText;
            }
            set
            {
                if (Shape is XTextRange)
                {
                    XTextRange xTextRange = Shape as XTextRange;
                    if (xTextRange != null) xTextRange.setString(value); 
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is a text element.
        /// </summary>
        /// <value><c>true</c> if this instance is text; otherwise, <c>false</c>.</value>
        virtual public bool IsText
        {
            get
            {
                try
                {
                    return OoUtils.ElementSupportsService(Shape, OO.Services.DRAW_SHAPE_TEXT);
                }
                catch (System.Threading.ThreadAbortException) { }
                catch (unoidl.com.sun.star.lang.DisposedException) { Dispose(); }
                catch (Exception)
                {
                    //TODO: invalidate
                    return false;
                }
                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has text.
        /// </summary>
        /// <value><c>true</c> if this instance has text; otherwise, <c>false</c>.</value>
        virtual public bool HasText
        {
            get
            {
                return !String.IsNullOrEmpty(Text);
            }
        }

        /// <summary>
        /// Gets or sets the title of the object.
        /// </summary>
        /// <value>The title.</value>
        virtual public String Title
        {
            get { return getStringProperty("Title"); }
            set { setStringProperty("Title", value); }
        }

        /// <summary>
        /// provides a human-readable name (which can be presented at the UI) for a category.  
        /// </summary>
        /// <value>The UI name singular.</value>
        virtual public String UINameSingular
        {
            get { return getStringProperty("UINameSingular"); }
        }

        /// <summary>
        /// provides a human-readable name (which can be presented at the UI) for a category.  
        /// </summary>
        /// <value>The UI name plural.</value>
        virtual public String UINamePlural
        {
            get { return getStringProperty("UINamePlural"); }
        }

        /// <summary>
        /// [ OPTIONAL ]
        /// If this is false, the Shape is not visible on screen outputs. Please note that the Shape may still be visible when printed, see Printable.  
        /// </summary>
        /// <value><c>true</c> if visible; otherwise, <c>false</c>.</value>
        virtual public bool Visible
        {
            get { return getBoolProperty("Visible"); }
            set { setBoolProperty("Visible", value); }
        }

        /// <summary>
        /// ZOrder 	[ OPTIONAL ]
        /// is used to query or change the ZOrder of this Shape.  
        /// </summary>
        /// <value>The Z order.</value>
        virtual public int ZOrder
        {
            get { return getIntProperty("ZOrder"); }
            set { setIntProperty("ZOrder", value); }
        }

        /// <summary>
        /// the current position of this object.
        /// </summary>
        /// <value>the position of the top left edge in 100/th mm.</value>
        virtual public System.Drawing.Point Position
        {
            get
            {
                unoidl.com.sun.star.awt.Point pos = null;
                if (Shape != null) pos = Shape.getPosition();
                return pos != null ? new System.Drawing.Point(pos.X, pos.Y) : new System.Drawing.Point();
            }
            set
            {
                try
                {
                    if (!MoveProtect)
                    {
                        try
                        {

                            var oldPos = Position;
                            OoUtils.AddActionToUndoManager(
                                GetDocument() as XUndoManagerSupplier,
                                new ActionUndo(
                                        "Move Shape " + Name,
                                        () => { Position = oldPos; },
                                        () => { Position = value; }
                                    )
                                );

                            Shape.setPosition(value.IsEmpty ? new unoidl.com.sun.star.awt.Point(0, 0) : new unoidl.com.sun.star.awt.Point(value.X, value.Y));
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// the size of this object.
        /// </summary>
        /// <value>the size in 100/th mm.</value>
        virtual public System.Drawing.Size Size
        {
            get
            {
                unoidl.com.sun.star.awt.Size size = null;
                if (Shape != null) size = Shape.getSize();
                return size != null ? new System.Drawing.Size(size.Width, size.Height) : new System.Drawing.Size();
            }
            set
            {
                if (!SizeProtect)
                    if (Shape != null)
                        try
                        {
                            //if (IsTransformed)
                            //{
                            //    var trans = TransformationMatrix;
                            //}

                            var oldSize = Size;
                            OoUtils.AddActionToUndoManager(
                                GetDocument() as XUndoManagerSupplier,
                                new ActionUndo(
                                        "Scale Shape " + Name,
                                        () => { Size = oldSize; },
                                        () => { Size = value; }
                                    )
                                );

                            Shape.setSize(value.IsEmpty ? new unoidl.com.sun.star.awt.Size(0, 0) : new unoidl.com.sun.star.awt.Size(value.Width, value.Height));
                        }
                        catch { }
            }
        }

        /// <summary>
        /// Gets the bounds (position and size) of the shape on the DRAW page ind 100th/mm.
        /// </summary>
        /// <value>
        /// The bounds.
        /// </value>
        virtual public System.Drawing.Rectangle Bounds
        {
            get
            {
                if (Shape != null)
                    return OoDrawUtils.GetBounds(Shape);
                return new System.Drawing.Rectangle();
            }
        }



        /// <summary>
        /// This is the angle for rotation of this shape in 1/100th of a degree. 
        /// The shape is rotated counter-clockwise around the center of the bounding box.
        /// </summary>
        /// <value>The rotation in 1/100th of a degree.</value>
        virtual public int Rotation
        {
            get
            {
                //if (
                //    OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR)
                //    || OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_CUSTOM)
                //    )
                //{
                return getIntProperty("RotateAngle");
                //}
                //else return 0;
            }
            set
            {
                //if (OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR)
                //    || OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_CUSTOM))
                //{
                setIntProperty("RotateAngle", value % 36000);
                //}
            }
        }

        /// <summary>
        /// This is the amount of shearing for this shape in 1/100th of a degree.
        /// The shape is sheared clockwise around the center of the bounding box.
        /// </summary>
        /// <value>The shear angle in 1/100th of a degree.</value>
        virtual public int Shear
        {
            get
            {
                if (OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR)
                    || OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_CUSTOM))
                {
                    return getIntProperty("ShearAngle");
                }
                else return 0;
            }
            set
            {
                if (OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR)
                    || OoUtils.ElementSupportsService(Shape, OO.Services.DRAWING_PROPERTIES_CUSTOM))
                {
                    setIntProperty("ShearAngle", value);
                }
            }
        }

        #region Transformation

        /// <summary>
        /// Gets a value indicating whether this instance is transformed.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is transformed; otherwise, <c>false</c>.
        /// </value>
        virtual public bool IsTransformed
        {
            //FIXME: check if this is always true
            get { return OoUtils.GetProperty(Shape, "Transformation") != null; }
        }

        /// <summary>
        /// Gets the transformation matrix.
        /// </summary>
        /// <value>The transformation matrix.</value>
        virtual public double[,] TransformationMatrix
        {
            get
            {
                double[,] transformMatrix = new double[,] { };
                unoidl.com.sun.star.drawing.HomogenMatrix3 transformProp = OoUtils.GetProperty(Shape, "Transformation") as unoidl.com.sun.star.drawing.HomogenMatrix3;
                if (transformProp != null)
                {
                    transformMatrix = OoDrawUtils.ConvertHomogenMatrix3ToMatrix(transformProp);
                }
                return transformMatrix;
            }
        }

        #endregion

        #region Background
        /// <summary>
        /// Sets the background image bitmap.
        /// </summary>
        /// <param name="imagePath">The local path to the bitmap.</param>
        /// <param name="Name">The name of the Fill Bitmap.</param>
        /// <returns>
        /// 	<c>true</c> if the background image could been set, otherwise <c>false</c>
        /// </returns>
        virtual public bool SetBackgroundBitmap(String imagePath, String Name = "") { return SetBackgroundBitmap(imagePath, tud.mci.tangram.util.BitmapMode.REPEAT, Name); }
        /// <summary>
        /// Sets the background image bitmap.
        /// </summary>
        /// <param name="imagePath">The local path to the bitmap.</param>
        /// <param name="mode">The fill mode <see cref="FillBitmapMode"/>.</param>
        /// <param name="Name">The name of the Fill Bitmap.</param>
        /// <returns>
        /// 	<c>true</c> if the background image could been set, otherwise <c>false</c>
        /// </returns>
        virtual public bool SetBackgroundBitmap(String imagePath, tud.mci.tangram.util.BitmapMode mode, String Name = "")
        {
            XBitmap bitmap = null;
            bitmap = OoUtils.GetGraphicFromUrl(imagePath) as XBitmap;
            return SetBackgroundBitmap(bitmap, mode, Name);
        }

        /// <summary>
        /// Get the name of the used fill bitmap style
        /// </summary>
        /// <returns>Returns name of the used fill bitmap style</returns>
        virtual public string GetBackgroundBitmapName()
        {
            return getStringProperty("FillBitmapName");
        }


        /// <summary>
        /// Sets the background image bitmap.
        /// </summary>
        /// <param name="bitmap">The bitmap.</param>
        /// <param name="Name">The name of the Fill Bitmap.</param>
        /// <returns>
        /// 	<c>true</c> if the background image could been set, otherwise <c>false</c>
        /// </returns>
        virtual internal bool SetBackgroundBitmap(XBitmap bitmap, String Name = "") { return SetBackgroundBitmap(bitmap, tud.mci.tangram.util.BitmapMode.REPEAT, Name); }
        /// <summary>
        /// Sets the background image bitmap.
        /// </summary>
        /// <param name="bitmap">The bitmap.</param>
        /// <param name="mode">The fill mode <see cref="FillBitmapMode"/>.</param>
        /// <param name="Name">The name of the Fill Bitmap.</param>
        /// <returns><c>true</c> if the background image could been set, otherwise <c>false</c></returns>
        virtual internal bool SetBackgroundBitmap(XBitmap bitmap, tud.mci.tangram.util.BitmapMode mode, String Name = "")
        {
            try
            {
                int count = this.ChildCount;
                if (count > 0) // group object
                {
                    var children = GetChilderen();
                    for (int i = 0; i < children.Count; i++)
                    {
                        OoShapeObserver child = children.ElementAt(i);

                        if (child != null)
                        {
                            child.SetBackgroundBitmap(bitmap, mode, Name);
                        }
                    }
                }

                if (bitmap != null)
                {
                    SetProperty("FillBitmap", bitmap);
                    FillBitmapMode = mode;
                    setStringProperty("FillBitmapName", Name);
                }
                return true;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] Can't set background image: " + ex);
            }
            return false;
        }


        /// <summary>
        /// If the property FillStyle is set to FillStyle::BITMAP, this is the bitmap used.
        /// </summary>
        /// <value>The fill bitmap.</value>
        virtual internal XBitmap FillBitmap
        {
            set { SetProperty("FillBitmap", value); }
            get
            {
                var bitmap = GetProperty("FillBitmap");
                if (bitmap != null && bitmap is XBitmap)
                {
                    return bitmap as XBitmap;
                }
                return null;
            }
        }

        #endregion

        #region Polygon Points

        OoPolygonPointsObserver _ppObs = null;
        /// <summary>
        /// Gets the polygon points.
        /// </summary>
        /// <returns>a warper for helping handling polygon points or <c>null</c> if this is not a shape that have polygon points.</returns>
        virtual public OoPolygonPointsObserver GetPolygonPointsObserver()
        {
            if (_ppObs == null && PolygonHelper.IsFreeform(Shape))
            {
                _ppObs = new OoPolygonPointsObserver(this);
            }
            return _ppObs;
        }

        /// <summary>
        /// Gets the direct access to the poly polygon points.
        /// </summary>
        /// <returns>list of polygon point lists</returns>
        virtual public List<List<PolyPointDescriptor>> GetPolyPolygonPoints()
        {
            var points = PolygonHelper.GetPolyPoints(Shape);

            return points;
        }

        /// <summary>
        /// Sets the poly polygon points to the shape.
        /// </summary>
        /// <param name="points">The points.</param>
        virtual public void SetPolyPolygonPoints(List<List<PolyPointDescriptor>> points)
        {
            //System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            //watch.Restart();
            try
            {
                Task t = new Task(
                    () =>
                    {
                        LockValidation = true;
                        try { PolygonHelper.SetPolyPoints(Shape, points, false, GetDocument()); }
                        finally { LockValidation = false; }
                    }
                );
                t.Start();
            }
            finally
            {
                //watch.Stop();
                //System.Diagnostics.Debug.WriteLine("Ticks for handling Polygon setting Points:" + watch.ElapsedTicks + " / Milliseconds:" + watch.ElapsedMilliseconds);
            }
        }

        #endregion

        #region Generic

        /// <summary>
        /// Try to return the value of the requested property.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>the value as object</returns>
        public Object GetProperty(string name)
        {
            return OoUtils.GetProperty(Shape as XPropertySet, name);
        }

        /// <summary>
        /// Try to sets the property. The function doesn't give feedback 
        /// if the property is valid or available
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        public bool SetProperty(string name, Object value)
        {
            return OoUtils.SetPropertyUndoable(Shape as XPropertySet, name, value, GetDocument() as XUndoManagerSupplier);
        }

        #region String

        private String getStringProperty(string name)
        {
            Object prop = OoUtils.GetProperty(Shape as XPropertySet, name);
            string sProp = prop != null ? prop.ToString() : String.Empty;
            return sProp;
        }

        private bool setStringProperty(string name, string value)
        {
            string val = OoUtils.GetStringProperty(Shape as XPropertySet, name);
            if (val.Equals(value)) return true;
            try
            {
                //OoUtils.SetPropertyUndoable(Shape, name, value, getDocument() as XUndoManagerSupplier);                
                OoUtils.SetStringProperty(Shape as XPropertySet, name, value);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write("exception during setting property " + ex);
            }
            string newVal = getStringProperty(name);
            return newVal.Equals(value);
        }

        #endregion

        #region boolean

        private bool getBoolProperty(string name)
        {
            Object prop = OoUtils.GetProperty(Shape as XPropertySet, name);
            return prop != null && prop is bool ? (bool)prop : false;
        }

        private bool setBoolProperty(string name, bool value)
        {
            if (OoUtils.GetBooleanProperty(Shape as XPropertySet, name).Equals(value)) return true;
            OoUtils.SetPropertyUndoable(Shape as XPropertySet, name, value, GetDocument() as XUndoManagerSupplier);
            bool newVal = getBoolProperty(name);
            return newVal.Equals(value);
        }

        #endregion

        #region integer

        private int getIntProperty(string name)
        {
            return OoUtils.GetIntProperty(Shape as XPropertySet, name);
        }

        private bool setIntProperty(string name, int value)
        {

            int oldVal = getIntProperty(name);
            //if (OoUtils.GetIntProperty(Shape, name).Equals(value)) return true;
            OoUtils.SetPropertyUndoable(Shape as XPropertySet, name, value, GetDocument() as XUndoManagerSupplier);
            int newVal = getIntProperty(name);

            return newVal.Equals(value);
        }

        #endregion

        #endregion

        #endregion

        #region Group Stuff

        /// <summary>
        /// Gets a value indicating whether this instance is a group of shapes.
        /// </summary>
        /// <value><c>true</c> if this instance is a group; otherwise, <c>false</c>.</value>
        virtual public bool IsGroup { get { return Shape is XShapes; } }

        /// <summary>
        /// Gets a value indicating whether this instance has children.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance has children; otherwise, <c>false</c>.
        /// </value>
        virtual public bool HasChildren { get { return IsGroup && ((XShapes)Shape).hasElements(); } }

        /// <summary>
        /// Gets the amount of child if in this group if this is a group.
        /// </summary>
        /// <value>The number of children.</value>
        virtual public int ChildCount { get { return HasChildren ? (((XShapes)Shape).getCount()) : 0; } }

        /// <summary>
        /// Gets a value indicating whether this instance is a child. This is nearly always true. Check if the parent is a group.
        /// </summary>
        /// <value><c>true</c> if this instance is child; otherwise, <c>false</c>. This is nearly always true. Check if the parent is a group.</value>
        virtual public bool IsChild { get { return (Shape is XChild) && (((XChild)Shape).getParent() != null); } }

        /// <summary>
        /// Gets a value indicating whether this instance is a member of a group. [CURRENTLY DOES NOT WORK PROPPERLY]
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is a group member; otherwise, <c>false</c>.
        /// </value>
        virtual public bool IsGroupMember
        {
            get
            {
                bool child = false;
                //if (Parent != null) return true;
                if (Shape is XChild)
                {
                    var parent = ((XChild)Shape).getParent();
                    if (parent != null && !(parent is XDrawPage) && parent is XShapes && parent != Shape)
                    {
                        // determine if parent is not the document and not the page shape
                        child = true;
                    }
                }
                if (child)
                {
                    if (Parent == null)
                    {
                        //TODO: update parent
                    }
                }
                else
                {
                    if (Parent != null) Parent = null;
                }



                return child;
            }
        }

        #endregion

    }
}
