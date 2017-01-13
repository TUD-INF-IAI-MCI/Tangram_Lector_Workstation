using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.drawing;

namespace tud.mci.tangram.controller.observer
{
    public class OoCustomShapeObserver : OoShapeObserver, INameBuilder, IUpdateable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OoShapeObserver"/> class.
        /// </summary>
        /// <param name="s">The XShape to observe.</param>
        /// <param name="page">The observer for the page the shape is located on.</param>
        public OoCustomShapeObserver(XShape s, OoDrawPageObserver page) : this(s, page, null) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="OoShapeObserver"/> class.
        /// </summary>
        /// <param name="s">The XShape to observe.</param>
        /// <param name="page">The observer for the page the shape is located on.</param>
        /// <param name="parent">The observer for the parent shape.</param>
        public OoCustomShapeObserver(XShape s, OoDrawPageObserver page, OoShapeObserver parent)
            : base(s, page, parent)
        {
            // getCustomShapeHandles();
            //GetAdjustmentValues();
        }


        #region INameBuilder

        string _uid = String.Empty;
        string _baseUid = String.Empty;
        /// <summary>
        /// Builds a unique, human-readable name for the object.
        /// </summary>
        /// <returns>
        /// A unique name for the object.
        /// </returns>
        string INameBuilder.BuildName()
        {
            _uid = Name;
            if (String.IsNullOrWhiteSpace(_uid))
            {
                var type = GetCustomShapeType();
                if (String.IsNullOrWhiteSpace(type))
                    _uid = UINameSingular;
                else
                    _uid = type;
            }

            _baseUid = _uid;
            return _uid;
        }

        int _tryies = 0;
        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <returns>
        /// A new try for a unique name for the object.
        /// </returns>
        string INameBuilder.RebuildName()
        {
            if (String.IsNullOrWhiteSpace(_baseUid))
                ((INameBuilder)this).BuildName();
            _uid = _baseUid + "_" + ++_tryies;
            return _uid;
        }

        /// <summary>
        /// Rebuilds the name based on the previously build name - e.g. because the name already exists.
        /// </summary>
        /// <param name="startIndex">The start index for the next try - e.g. 4 objects of the same type already exists.</param>
        /// <returns>
        /// A new try for a unique name for the object.
        /// </returns>
        string INameBuilder.RebuildName(int startIndex)
        {
            _tryies = startIndex;
            return ((INameBuilder)this).RebuildName();
        }

        #endregion

        #region CustomShapePropreties

        Dictionary<String, PropertyValue> _cachedCustonProperties = null;
        Dictionary<String, PropertyValue> getCustomProperties(bool cached = true)
        {
            if (_cachedCustonProperties == null || !cached)
            {
                updateCustomShapeProperties();
            }
            return _cachedCustonProperties;
        }

        PropertyValue[] _csprop = null;
        PropertyValue[] updateCustomShapeProperties()
        {
            _csprop = OoUtils.GetProperty(Shape, "CustomShapeGeometry") as PropertyValue[];
            if (_csprop != null)
            {
                _cachedCustonProperties = OoUtils.GetPropertyvalueDictionary(_csprop);
            }
            return _csprop;
        }

        string _csType = String.Empty;
        /// <summary>
        /// Gets the type of the custom shape.
        /// This is normally some kind of human readable name for the custom shape.
        /// </summary>
        /// <returns>The 'Type' property from the custom shape properties.</returns>
        public string GetCustomShapeType()
        {
            if (String.IsNullOrEmpty(_csType))
            {
                var props = getCustomProperties();
                if (props.ContainsKey("Type"))
                {
                    _csType = props["Type"].Value.Value.ToString();
                }
            }
            return _csType;
        }


        /// <summary>
        /// Gets the adjustment values for this custom shape. 
        /// With this the interaction handle controls can be manipulated to adapt the shape.
        /// Property: 'CustomShapeGeometry'.'AdjustmentValues' .
        /// </summary>
        /// <returns>List of available adjustment controls.</returns>
        public List<CustomShapeAdjustment> GetAdjustmentValues()
        {
            List<CustomShapeAdjustment> adjs = new List<CustomShapeAdjustment>();
            var props = getCustomProperties(false);
            if (props.ContainsKey("AdjustmentValues"))
            {
                var adjustments = props["AdjustmentValues"].Value.Value as unoidl.com.sun.star.drawing.EnhancedCustomShapeAdjustmentValue[];
                if (adjustments != null && adjustments.Length > 0)
                {
                    foreach (EnhancedCustomShapeAdjustmentValue adjustment in adjustments)
                    {
                        CustomShapeAdjustment csAdj = new CustomShapeAdjustment(adjustment);
                        adjs.Add(csAdj);
                    }
                }
            }
            return adjs;
        }


        //List<CustomShapeHandle> getCustomShapeHandles()
        //{
        //    List<CustomShapeHandle> hndls = new List<CustomShapeHandle>();
        //    var props = getCustomProperties(false);
        //    if (props.ContainsKey("Handles"))
        //    {
        //        var handles = props["Handles"].Value.Value as PropertyValue[][];
        //        if (handles != null && handles.Length > 0)
        //        {
        //            foreach (PropertyValue[] handle in handles)
        //            {
        //                CustomShapeHandle csh = new CustomShapeHandle(handle);
        //                hndls.Add(csh);

        //                // FIXME: only for fixing
        //                // -- DOES NOT WORK AT ALL: only moves the handle but does not affect the shape :-(
        //                // bool succ = csh.Move(100, 100);
        //            }
        //        }
        //    }

        //    //// FIXME: only a test -- DOES NOT WORK AT ALL: only moves the handle but does not affect the shape :-(
        //    //// write back to shape
        //    //PropertyValue tmp1 = buildHanldePropertyValues(hndls);
        //    //PropertyValue[] tmp2 = writeToCustomShapeProperties("Handles", tmp1);
        //    //bool succss = writeHandlesToShape(tmp2);

        //    return hndls;
        //}


        #endregion

        #region IUpdateable

        new void Update()
        {
            base.Update();
            updateCustomShapeProperties();
        }

        #endregion


        #region Helper

        internal bool SetAdjustmentValues(List<CustomShapeAdjustment> adjs)
        {
            // write back to shape
            PropertyValue tmp1 = buildAdjustmentPropertyValue(adjs);
            PropertyValue[] tmp2 = writeToCustomShapeProperties("AdjustmentValues", tmp1);
            bool success = writeCustomShapeGeometryToShape(tmp2);
            return success;
        }

        /// <summary>
        /// Builds a property value from a CustomShapeAdjustment list to store it in a property.
        /// </summary>
        /// <param name="adjustments">The adjustments.</param>
        /// <returns>a property value to store in a property e.g. CustomShapeGeometry.AdjustmentValues .</returns>
        internal static PropertyValue buildAdjustmentPropertyValue(List<CustomShapeAdjustment> adjustments)
        {
            //Handle        long                  0                
            //Name          string                AdjustmentValues 
            //State         .beans.PropertyState  DIRECT_VALUE     
            //Value         any                   -Sequence-   

            PropertyValue handlesPropValue = new PropertyValue();
            handlesPropValue.Name = "AdjustmentValues";
            handlesPropValue.Handle = 0;
            handlesPropValue.State = PropertyState.DIRECT_VALUE;

            if (adjustments != null && adjustments.Count > 0)
            {
                unoidl.com.sun.star.drawing.EnhancedCustomShapeAdjustmentValue[] value = new EnhancedCustomShapeAdjustmentValue[adjustments.Count];
                for (int i = 0; i < adjustments.Count; i++)
                {
                    value[i] = adjustments[i].ToEnhancedCustomShapeAdjustmentValue();
                }

                handlesPropValue.Value = Any.GetAsOne(value);
            }

            return handlesPropValue;
        }

        /// <summary>
        /// Writes a special property to the cached customProperty set and returns the updated values.
        /// ATTENTION: the properties are not stored to the shape by this function.
        /// </summary>
        /// <param name="propName">Name of the property to update.</param>
        /// <param name="propValue">The property value to set.</param>
        /// <returns>the whole updated 'CustomShapeGeometry' set for storing it to the shape.</returns>
        PropertyValue[] writeToCustomShapeProperties(string propName, PropertyValue propValue)
        {
            var properties = updateCustomShapeProperties();
            if (properties != null && properties.Length > 0)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    if (properties[i] != null && properties[i].Name.Equals(propName))
                    {
                        properties[i].Value = propValue.Value;
                        return properties;
                    }
                }
            }
            return properties;
        }

        /// <summary>
        /// Writes the custom shape geometry properties to the shape.
        /// The undo manager is used to make this manipulation undo- / redo-able.
        /// </summary>
        /// <param name="propertyValues">The property values.</param>
        /// <returns></returns>
        bool writeCustomShapeGeometryToShape(PropertyValue[] propertyValues)
        {
            return OoUtils.SetPropertyUndoable(
                Shape,
                "CustomShapeGeometry",
                propertyValues,
                this.GetDocument() as unoidl.com.sun.star.document.XUndoManagerSupplier);
        }

        #endregion

        #region Polygon Points

        OoPolygonPointsObserver _ppObs = null;

        /// <summary>
        /// Gets the polygon points.
        /// </summary>
        /// <returns>a warper for helping handling polygon points or <c>null</c> if this is not a shape that have polygon points.</returns>
        override public OoPolygonPointsObserver GetPolygonPointsObserver()
        {
            if (_ppObs == null)
            {
                _ppObs = new CustomShapeAjustmentPintsObserver(this);
            }
            return _ppObs;
        }

        /// <summary>
        /// Gets the direct access to the poly polygon points.
        /// </summary>
        /// <returns>list of polygon point lists</returns>
        override public List<List<PolyPointDescriptor>> GetPolyPolygonPoints()
        {
            var obs = GetPolygonPointsObserver();
            if (obs != null)
            {
                return obs.CachedPolyPointList;
            }
            return new List<List<PolyPointDescriptor>>();
        }

        /// <summary>
        /// Sets the poly polygon points to the shape.
        /// </summary>
        /// <param name="points">The points.</param>
        override public void SetPolyPolygonPoints(List<List<PolyPointDescriptor>> points)
        {
            throw new NotImplementedException();
        }

        #endregion






    }

    /// <summary>
    /// Struct wrapping Adjustment control information.
    /// </summary>
    public struct CustomShapeAdjustment
    {
        double _value;
        /// <summary>
        /// Gets or sets the value of this one-dimensional adjustment value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public Double Value { get { return _value; } set { _value = value; } }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomShapeAdjustment"/> struct.
        /// </summary>
        /// <param name="val">The property value for 'AdjustmentValues' from the 'CustomShapeGeometry' property.</param>
        internal CustomShapeAdjustment(EnhancedCustomShapeAdjustmentValue val)
        {
            _value = 0;
            if (val != null && val.Value.hasValue())
            {
                Value = (double)val.Value.Value;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomShapeAdjustment"/> struct.
        /// </summary>
        /// <param name="val">The value.</param>
        public CustomShapeAdjustment(double val)
        {
            _value = val;
        }

        /// <summary>
        /// To the enhanced custom shape adjustment value.
        /// </summary>
        /// <returns></returns>
        internal EnhancedCustomShapeAdjustmentValue ToEnhancedCustomShapeAdjustmentValue()
        {
            return new EnhancedCustomShapeAdjustmentValue(Any.Get(Value), PropertyState.DIRECT_VALUE, "");
        }
    }



    //public struct CustomShapeHandle
    //{
    //    // Position
    //    // If the property Polar is set, 
    //    // then the first value specifies the radius and 
    //    // the second parameter the angle of the handle.
    //    // Otherwise, if the handle is not polar, 
    //    // the first parameter specifies the horizontal handle position, 
    //    // the vertical handle position is described by the second parameter.

    //    int x;
    //    /// <summary>
    //    /// Gets or sets the x position of the handle.
    //    /// </summary>
    //    /// <value>
    //    /// The x position.
    //    /// </value>
    //    public int X
    //    {
    //        get { return x; }
    //        set
    //        {
    //            x = Math.Min(RangeXMaximum, Math.Max(RangeXMinimum, value));
    //        }
    //    }
    //    /// <summary>
    //    /// determines if the handle is changeable in horizontal direction.
    //    /// </summary>
    //    public readonly bool IsHorizontalMovable;


    //    int y;
    //    /// <summary>
    //    /// Gets or sets the y position of the handle.
    //    /// </summary>
    //    /// <value>
    //    /// The y position.
    //    /// </value>
    //    public int Y
    //    {
    //        get { return y; }
    //        set
    //        {
    //            y = Math.Min(RangeYMaximum, Math.Max(RangeYMinimum, value));
    //        }
    //    }
    //    /// <summary>
    //    /// determines if the handle is changeable in vertical direction.
    //    /// </summary>
    //    public readonly bool IsVerticalMovable;

    //    /// <summary>
    //    /// The minimum y value
    //    /// If the attribute RangeYMinimum is set, it specifies the vertical minimum range of the handle.
    //    /// </summary>
    //    public readonly int RangeYMinimum;
    //    /// <summary>
    //    /// The maximum y value
    //    /// If the attribute RangeYMaximum is set, it specifies the vertical maximum range of the handle.
    //    /// </summary>
    //    public readonly int RangeYMaximum;

    //    /// <summary>
    //    /// The minimum x value
    //    /// If the attribute RangeXMinimum is set, it specifies the horizontal minimum range of the handle.
    //    /// </summary>
    //    public readonly int RangeXMinimum;
    //    /// <summary>
    //    /// The maximum x value.
    //    /// If the attribute RangeXMaximum is set, it specifies the horizontal maximum range of the handle.
    //    /// </summary>
    //    public readonly int RangeXMaximum;



    //    // MirroredX
    //    /// <summary>
    //    ///  Specifies if the x position of the handle is mirrored.
    //    /// </summary>
    //    /// <returns>
    //    ///   <c>true</c> if [is mirrored x]; otherwise, <c>false</c>.
    //    /// </returns>
    //    public bool IsMirroredX()
    //    {
    //        if (propertyDict != null && propertyDict.ContainsKey("MirroredX"))
    //        {
    //            return (bool)propertyDict["MirroredX"].Value.Value;
    //        }
    //        return false;
    //    }

    //    // MirroredY
    //    // Specifies if the y position of the handle is mirrored.
    //    public bool IsMirroredY()
    //    {
    //        if (propertyDict != null && propertyDict.ContainsKey("MirroredY"))
    //        {
    //            return (bool)propertyDict["MirroredY"].Value.Value;
    //        }
    //        return false;
    //    }

    //    // Polar
    //    // If this attribute is set, the handle is a polar handle.
    //    // The property specifies the center position of the handle. 
    //    // If this attribute is set, the attributes RangeX and RangeY are ignored, 
    //    // instead the attribute RadiusRange is used.

    //    // RadiusRangeMaximum
    //    // If this attribute is set, it specifies the maximum radius range that can be used for a polar handle.

    //    // RadiusRangeMinimum
    //    // If this attribute is set, it specifies the minimum radius range that can be used for a polar handle.


    //    // long RefAngle
    //    // RefAngle, if this attribute is set, it specifies the index of the adjustment value which is connected to the angle of the handle.

    //    // long RefR
    //    // RefR, if this attribute is set, it specifies the index of the adjustment value which is connected to the radius of the handle.

    //    // long RefX
    //    // RefX, if this attribute is set, it specifies the index of the adjustment value which is connected to the horizontal position of the handle.

    //    // long RefY
    //    // RefY, if this attribute is set, it specifies the index of the adjustment value which is connected to the vertical position of the handle.


    //    // boolean Switched
    //    // Specifies if the handle directions are swapped if the shape is taller than wide.


    //    PropertyValue[] properties;
    //    Dictionary<String, PropertyValue> propertyDict;

    //    #region Constructor

    //    /// <summary>
    //    /// Initializes a new instance of the <see cref="CustomShapeHandle"/> struct.
    //    /// </summary>
    //    /// <param name="vals">The vals.</param>
    //    internal CustomShapeHandle(PropertyValue[] vals)
    //    {
    //        propertyDict = OoUtils.GetPropertyvalueDictionary(vals);
    //        properties = vals;


    //        // position
    //        x = 0;
    //        y = 0;
    //        IsHorizontalMovable = false;
    //        IsVerticalMovable = false;

    //        if (propertyDict.ContainsKey("Position"))
    //        {
    //            var pos = propertyDict["Position"].Value.Value;
    //            if (pos != null && pos is EnhancedCustomShapeParameterPair)
    //            {
    //                var xVal = ((EnhancedCustomShapeParameterPair)pos).First.Value.Value;
    //                x = (int)xVal;
    //                IsHorizontalMovable = ((EnhancedCustomShapeParameterPair)pos).First.Type == 2;

    //                var yVal = ((EnhancedCustomShapeParameterPair)pos).Second.Value.Value;
    //                y = (int)yVal;
    //                IsVerticalMovable = ((EnhancedCustomShapeParameterPair)pos).Second.Type == 2;
    //            }
    //        }

    //        // x Range
    //        RangeXMinimum = x;
    //        RangeXMaximum = x;
    //        if (propertyDict.ContainsKey("RangeXMinimum") && propertyDict.ContainsKey("RangeXMaximum"))
    //        {
    //            var xMin = propertyDict["RangeXMinimum"].Value.Value;
    //            if (xMin != null && xMin is EnhancedCustomShapeParameter)
    //            {
    //                RangeXMinimum = (int)(((EnhancedCustomShapeParameter)xMin).Value.Value);
    //            }
    //            var xMax = propertyDict["RangeXMaximum"].Value.Value;
    //            if (xMax != null && xMax is EnhancedCustomShapeParameter)
    //            {
    //                RangeXMaximum = (int)(((EnhancedCustomShapeParameter)xMax).Value.Value);
    //            }
    //        }

    //        // y Range
    //        RangeYMinimum = y;
    //        RangeYMaximum = y;
    //        if (propertyDict.ContainsKey("RangeYMinimum") && propertyDict.ContainsKey("RangeYMaximum"))
    //        {
    //            var yMin = propertyDict["RangeYMinimum"].Value.Value;
    //            if (yMin != null && yMin is EnhancedCustomShapeParameter)
    //            {
    //                RangeYMinimum = (int)(((EnhancedCustomShapeParameter)yMin).Value.Value);
    //            }
    //            var yMax = propertyDict["RangeYMaximum"].Value.Value;
    //            if (yMax != null && yMax is EnhancedCustomShapeParameter)
    //            {
    //                RangeYMaximum = (int)(((EnhancedCustomShapeParameter)yMax).Value.Value);
    //            }
    //        }
    //    }

    //    #endregion


    //    #region Modification

    //    /// <summary>
    //    /// Moves the handle in certain directions.
    //    /// </summary>
    //    /// <param name="horizontal">The horizontal change (X).</param>
    //    /// <param name="vertical">The vertical change (Y).</param>
    //    /// <returns><c>true</c> if some changes happens; otherwise <c>false</c>.</returns>
    //    public bool Move(int horizontal, int vertical)
    //    {
    //        int xOld = X;
    //        int yOld = Y;

    //        X += horizontal;
    //        Y += vertical;

    //        bool update = Y != yOld || X != xOld;

    //        if (update)
    //        {
    //            var newProp = GetPositionPropertyValue();
    //            updateProperty("Position", newProp);
    //        }

    //        return update;

    //    }

    //    private void updateProperty(string p, PropertyValue newProp)
    //    {
    //        if (propertyDict.ContainsKey(p))
    //        {
    //            propertyDict[p] = newProp;
    //            foreach (var item in properties)
    //            {
    //                if (item != null && item.Name.Equals(p))
    //                {
    //                    item.Value = newProp.Value;
    //                }
    //            }
    //        }
    //    }

    //    #endregion






    //    #region Export

    //    internal PropertyValue[] ToPropertyValues()
    //    {
    //        return properties;
    //    }



    //    // [5] = {[Handles, unoidl.com.sun.star.beans.PropertyValue]}        
    //    //  State = DIRECT_VALUE
    //    //  Handle = 0
    //    //  Name = "Handles"
    //    //  Value = {uno.Any { Type= unoidl.com.sun.star.beans.PropertyValue[][], Value=unoidl.com.sun.star.beans.PropertyValue[][]}}
    //    //      Type = {Name = "PropertyValue[][]" FullName = "unoidl.com.sun.star.beans.PropertyValue[][]"}
    //    //      Value = {unoidl.com.sun.star.beans.PropertyValue[1][]}
    //    //
    //    //
    //    //      -----> THIS is for the export !!
    //    //
    //    //          [0] = {unoidl.com.sun.star.beans.PropertyValue[3]}
    //    //              [0] = {unoidl.com.sun.star.beans.PropertyValue}
    //    //                  Handle	0	            int
    //    // 		            Name	"Position"	    string
    //    //  		        State	DIRECT_VALUE	unoidl.com.sun.star.beans.PropertyState
    //    // 		            Value	{uno.Any { Type= unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair, Value=unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair}}	uno.Any
    //    //                      Type	{Name = "EnhancedCustomShapeParameterPair" FullName = "unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair"}	System.Type {System.RuntimeType}
    //    //                      Value	{unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair}	object {unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair}
    //    //                          First	{unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}	unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter
    //    //                          Second	{unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}	unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter
    //    // 	            [1]	{unoidl.com.sun.star.beans.PropertyValue}	unoidl.com.sun.star.beans.PropertyValue
    //    // 		            Handle	0	int
    //    //                  Name	"RangeXMinimum"	string
    //    //                  State	DIRECT_VALUE	unoidl.com.sun.star.beans.PropertyState
    //    //                  Value	{uno.Any { Type= unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter, Value=unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}}	uno.Any
    //    //                      Type	{Name = "EnhancedCustomShapeParameter" FullName = "unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter"}	System.Type {System.RuntimeType}
    //    //                      Value	{unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}	object {unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}
    //    //                          Type	0	short
    //    //                          Value	{uno.Any { Type= System.Int32, Value=0}}	uno.Any
    //    //                              Type	{Name = "Int32" FullName = "System.Int32"}	System.Type {System.RuntimeType}
    //    //                              Value	0	object {int}
    //    //              [2]	{unoidl.com.sun.star.beans.PropertyValue}	unoidl.com.sun.star.beans.PropertyValue
    //    //                  Handle	0	int
    //    // 		            Name	"RangeXMaximum"	string
    //    //                  State	DIRECT_VALUE	unoidl.com.sun.star.beans.PropertyState
    //    //                  Value	{uno.Any { Type= unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter, Value=unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}}	uno.Any
    //    //                      Type	{Name = "EnhancedCustomShapeParameter" FullName = "unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter"}	System.Type {System.RuntimeType}
    //    //                      Value	{unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}	object {unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter}
    //    //                          Type	0	short
    //    //                          Value	{uno.Any { Type= System.Int32, Value=10800}}	uno.Any
    //    //                              Type	{Name = "Int32" FullName = "System.Int32"}	System.Type {System.RuntimeType}
    //    // 		                        Value	10800	object {int}
    //    //




    //    //internal unoidl.com.sun.star.beans.PropertyValue[] ToPropertyValues(bool useXRange = true, bool useYRange = true)
    //    //{
    //    //    //unoidl.com.sun.star.beans.PropertyValue propVal = new PropertyValue();
    //    //    //propVal.Name = "Handle";
    //    //    //propVal.Handle = 0;
    //    //    //propVal.State = PropertyState.DIRECT_VALUE;

    //    //    int size = 1 + (useXRange ? 2 : 0) + (useYRange ? 2 : 0);
    //    //    unoidl.com.sun.star.beans.PropertyValue[] propVals = new  unoidl.com.sun.star.beans.PropertyValue[size];
    //    //    int i = 0;
    //    //    // "Position"
    //    //    propVals[i++] = GetPositionPropertyValue(useXRange, useYRange);



    //    //    return propVals;
    //    //}

    //    internal unoidl.com.sun.star.beans.PropertyValue GetPositionPropertyValue()
    //    {
    //        return GetPositionPropertyValue(IsHorizontalMovable, IsVerticalMovable);
    //    }

    //    internal unoidl.com.sun.star.beans.PropertyValue GetPositionPropertyValue(bool xRangeChangeable = true, bool yRangeChangeable = true)
    //    {
    //        unoidl.com.sun.star.beans.PropertyValue propVal = new PropertyValue();
    //        propVal.Name = "Position";
    //        propVal.Handle = 0;
    //        propVal.State = PropertyState.DIRECT_VALUE;

    //        /*
    //         * If the property Polar is set, then 
    //         * the first value specifies the radius and 
    //         * the second parameter the angle of the handle.
    //         * 
    //         * Otherwise, if the handle is not polar, 
    //         * the first parameter specifies the horizontal handle position, 
    //         * the vertical handle position is described by the second parameter.
    //         */

    //        // build value
    //        unoidl.com.sun.star.drawing.EnhancedCustomShapeParameterPair value = new EnhancedCustomShapeParameterPair(
    //            // First
    //            new unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter(Any.Get(X), (short)(xRangeChangeable ? 2 : 0)), // 0 = not movable | 2 = changeable
    //            // Second
    //            new unoidl.com.sun.star.drawing.EnhancedCustomShapeParameter(Any.Get(Y), (short)(yRangeChangeable ? 2 : 0))
    //            );
    //        propVal.Value = Any.Get(value);

    //        return propVal;
    //    }

    //    #endregion


    //}

    //enum EnhancedCustomShapeParameterType : short
    //{
    //    NONE = 0,
    //    CHANGEABLE = 2
    //}


    #region PolygonPointObserverForAjestmentPoints

    /// <summary>
    /// Wrapper class for handling poly polygon and poly Bezier function points.
    /// </summary>
    /// <seealso cref="tud.mci.tangram.models.Interfaces.IUpdateable" />
    /// <seealso cref="System.Collections.IEnumerable" />
    public class CustomShapeAjustmentPintsObserver : OoPolygonPointsObserver
    {
        #region Member

        private readonly Object _lock = new Object();
        private List<List<PolyPointDescriptor>> _cachedPolyPointList = new List<List<PolyPointDescriptor>>(1);
        /// <summary>
        /// Gets or sets the last checked out poly point list.
        /// </summary>
        /// <value>
        /// The last checked out poly point list.
        /// </value>
        new public List<List<PolyPointDescriptor>> CachedPolyPointList
        {
            get { lock (_lock) return _cachedPolyPointList; }
            private set { lock (_lock)_cachedPolyPointList = value; }
        }

        /// <summary>
        /// Gets the geometry poly point list.
        /// </summary>
        /// <value>
        /// The geometry poly point list.
        /// </value>
        new public List<List<PolyPointDescriptor>> GeometryPolyPointList
        {
            get
            {
                if (Shape != null)
                {
                    var AdjVals = ((OoCustomShapeObserver)Shape).GetAdjustmentValues();
                    var adjPPoints = TransformAdjustmentDescriptorsInPolyPointDescriptors(AdjVals);
                    return new List<List<PolyPointDescriptor>> { adjPPoints };
                }
                return null;
            }
        }

        ///// <summary>
        ///// Corresponding shape.
        ///// </summary>
        ///// <value>The shape.</value>
        //public OoShapeObserver Shape { get; private set; }

        volatile bool _updating = false;
        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoPolygonPointsObserver" /> class.
        /// </summary>
        /// <param name="shape">The corresponding shape.</param>
        /// <exception cref="ArgumentNullException">shape</exception>
        /// <exception cref="ArgumentException">shape must be a custom shape.</exception>
        public CustomShapeAjustmentPintsObserver(OoCustomShapeObserver shape)
        {
            if (shape == null) throw new ArgumentNullException("shape");
            // if (!PolygonHelper.IsFreeform(shape.Shape)) throw new ArgumentException("shape must be a polygon or a bezier curve", "shape");

            Shape = shape;
            Shape.BoundRectChangeEventHandlers += Shape_BoundRectChangeEventHandlers;
            Shape.ObserverDisposing += Shape_ObserverDisposing;
            Update();

        }

        private void Shape_BoundRectChangeEventHandlers()
        {
            Update();
        }

        void Shape_ObserverDisposing(object sender, EventArgs e)
        {
            Shape = null;
            CachedPolyPointList = null;
        }

        #endregion

        #region Tests Overrides

        /// <summary>
        /// Returns true if the parent Shape is valid.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </returns>
        override public bool IsValid()
        {
            try
            {
                return Shape != null && Shape is OoCustomShapeObserver && Shape.Shape != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Determines whether this CustomShape is a closed one or an open one.
        /// </summary>
        /// <param name="cached">if set to <c>true</c> the previously cached result will be returned.</param>
        /// <returns>
        ///   Custom shapes always return <c>false</c>.
        /// </returns>
        override public bool IsClosed(bool cached = true) { return false; }

        /// <summary>
        /// Determines whether this CustomShape is a Bezier defined form or not.
        /// </summary>
        /// <param name="cached">if set to <c>true</c> the previously cached result will be returned.</param>
        /// <returns>
        ///   Custom shapes always return <c>true</c> because they only contain CONTROL points.
        /// </returns>
        override public bool IsBezier(bool cached = true) { return true; }

        /// <summary>
        /// Determines whether the polygon description is a group of several polygons.
        /// </summary>
        /// <returns>
        ///   always return <c>false</c>.
        /// </returns>
        override public bool IsPolyPolygon() { return false; }

        #endregion

        #region Modification Overrides

        /// <summary>
        /// Writes the points to polygons properties.
        /// </summary>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). 
        /// <returns><c>true</c> if the points could be written to the properties</returns>
        override public bool WritePointsToPolygon(bool geometry = false)
        {
            if (IsValid()
                && CachedPolyPointList != null && CachedPolyPointList.Count > 0
                && CachedPolyPointList[0] != null && CachedPolyPointList[0].Count > 0)
            {
                OoShapeObserver.LockValidation = true;
                try
                {

                    List<CustomShapeAdjustment> adj = new List<CustomShapeAdjustment>(CachedPolyPointList[0].Count);
                    foreach (var item in CachedPolyPointList[0])
                    {
                        var a = GetAdjustmentFromPointDescriptor(item);
                        adj.Add(a);
                    }

                    bool success = ((OoCustomShapeObserver)Shape).SetAdjustmentValues(adj);
                    Update();
                    return success;
                }
                finally
                {
                    OoShapeObserver.LockValidation = false;
                }
            }
            return false;
        }

        /// <summary>
        /// Sets the poly point list.
        /// </summary>
        /// <param name="newPoints">The new points.</param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). 
        /// <returns><c>true</c> if the points could be written to the properties</returns>
        override public bool SetPolyPointList(List<List<PolyPointDescriptor>> newPoints, bool geometry = false)
        {
            if (newPoints.Count != CachedPolyPointList.Count || newPoints[0].Count != CachedPolyPointList[0].Count)
            {
                throw new ArgumentException("The amount of adjustment points is fixed - you cannot change this!", "newPoints");
            }


            if (IsValid())
            {
                OoShapeObserver.LockValidation = true;
                try
                {

                    throw new NotImplementedException();


                    //        bool success = PolygonHelper.SetPolyPoints(Shape.Shape, newPoints, geometry, Shape.GetDocument());
                    //        Update();
                    //        return success;
                }
                finally
                {
                    OoShapeObserver.LockValidation = false;
                }
            }
            return false;
        }

        #endregion

        #region Point Handling

        /// <summary>
        /// Gets the polygon points of a certain polygon of the poly polygon descriptor.
        /// </summary>
        /// <param name="index">The index of the polygon in the poly polygon descriptor.</param>
        /// <returns>the list of polygon point descriptors or an empty list if there is no polygon at the requested index.</returns>
        override public List<PolyPointDescriptor> GetPolygonPoints(int index)
        {
            try
            {
                _updating = true;
                if (CachedPolyPointList != null && CachedPolyPointList.Count > index)
                    return CachedPolyPointList[index];
                return null;
            }
            finally { _updating = false; }
        }

        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The index of the point inside the polygon description.</param>
        /// <param name="polygonIndex">Index of the polygon in a poly polygon group.</param>
        /// <returns>Point descriptor for the requested position or empty Point descriptor</returns>
        override public PolyPointDescriptor GetPolyPointDescriptor(int index, int polygonIndex)
        {
            if (polygonIndex == 0)
            {
                try
                {
                    _updating = true;
                    if (CachedPolyPointList != null && CachedPolyPointList.Count > 0 
                        && CachedPolyPointList[0] != null && CachedPolyPointList[0].Count > index)
                        return CachedPolyPointList[0][index];
                }
                finally { _updating = false; }
            }
            return new PolyPointDescriptor();
        }
        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The one-dimensional index of the point inside the polygon description.</param>
        /// <returns>
        /// Point descriptor for the requested position or empty Point descriptor
        /// </returns>
        override public PolyPointDescriptor GetPolyPointDescriptor(int index)
        {
            return GetPolyPointDescriptor(index, 0);
        }

        /// <summary>
        /// Updates a poly point descriptor.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The index of the poly point to update.</param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon. Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly; 
        /// otherwise you have to call <see cref="WritePointsToPolygon"/> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns><c>true</c> if the point could be updated</returns>
        override public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (ppd.Value != null && ppd.Value is double && IsValid())
            {
                try
                {
                    _updating = true;
                    if (CachedPolyPointList != null && CachedPolyPointList.Count > polygonIndex)
                    {
                        if (CachedPolyPointList[polygonIndex] == null) CachedPolyPointList[polygonIndex] = new List<PolyPointDescriptor>();
                        if (CachedPolyPointList[polygonIndex].Count > index)
                        {
                            CachedPolyPointList[polygonIndex][index] = ppd;
                            if (updateDirectly) { return WritePointsToPolygon(geometry); }
                            else { return true; }
                        }
                    }
                }
                finally { _updating = false; }
            }
            return false;
        }
        /// <summary>
        /// Updates a poly point descriptor.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The one-dimensional index of the poly point to update.</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update" />function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   <c>true</c> if the point could be updated
        /// </returns>
        override public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, bool updateDirectly = true, bool geometry = false)
        {
            return UpdatePolyPointDescriptor(ppd, index, 0, updateDirectly, geometry);
        }

        /// <summary>
        /// Adds a poly point descriptor.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The index of the poly point to add to. A index &lt;= 0 will add the point as fist; an index &gt; the length of the list will add it as last. Default is <c>Int32.MaxValue</c></param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon.  Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update" />function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to add a new CustomShape adjustment point!
        /// </returns>
        override public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }
        /// <summary>
        /// Adds a poly point descriptor.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The one-dimensional index of the poly point to add to. A index &lt;= 0 will add the point as fist; an index &gt; the length of the list will add it as last. Default is <c>Int32.MaxValue</c></param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update" />function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to add a new CustomShape adjustment point!
        /// </returns>
        override public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }

        /// <summary>
        /// Adds or updates a poly point descriptor. Point will be updated if it already exist; otherwise it will be added.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The index of the poly point to add to. A index &lt;= 0 will add the point as fist; an index &gt;= the length of the list will add it as last. Default is <c>Int32.MaxValue</c></param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon. Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update" />function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   <c>true</c> if the point could be added
        /// </returns>
        override public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (ppd.Value != null && ppd.Value is double && IsValid())
            {
                try
                {
                    _updating = true;
                    if (CachedPolyPointList != null && CachedPolyPointList.Count > polygonIndex)
                    {
                        if (CachedPolyPointList[polygonIndex] != null && CachedPolyPointList[polygonIndex].Count > index)
                        {
                            return UpdatePolyPointDescriptor(ppd, index, polygonIndex, updateDirectly, geometry);
                        }
                        //else
                        //{
                        //   // return AddPolyPointDescriptor(ppd, index, polygonIndex, updateDirectly, geometry);
                        //}
                    }
                }
                finally { _updating = false; }
            }

            return false;
        }
        /// <summary>
        /// Adds or updates a poly point descriptor. Point will be updated if it already exist; otherwise it will be added.
        /// </summary>
        /// <param name="ppd">The descriptor for a poly polygon point.</param>
        /// <param name="index">The one-dimensional index of the poly point to add to. A index &lt;= 0 will add the point as fist; an index &gt;= the length of the list will add it as last. Default is <c>Int32.MaxValue</c></param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update" />function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   <c>true</c> if the point could be added
        /// </returns>
        override public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            return AddOrUpdatePolyPointDescriptor(ppd, index, 0, updateDirectly, geometry);
        }

        /// <summary>
        /// Removes a poly point descriptor at a certain index. 
        /// </summary>
        /// <param name="index">The index of the poly point to remove.</param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon.  Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon"/> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c> </param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to remove a CustomShape adjustment point!
        /// </returns>
        override public bool RemovePolyPointDescriptor(int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }
        /// <summary>
        /// Removes a poly point descriptor at a certain index. 
        /// </summary>
        /// <param name="index">The one-dimensional index of the poly point to remove.</param>
        /// <param name="updateDirectly">if set to <c>true</c> the point lists will be updated in the shape properties directly;
        /// otherwise you have to call <see cref="WritePointsToPolygon"/> manually.
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c> </param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to remove a CustomShape adjustment point!
        /// </returns>
        override public bool RemovePolyPointDescriptor(int index, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }


        /// <summary>
        /// Removes a polygon (list of poly point descriptor).
        /// </summary>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon. Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the polygon lists will be updated in the shape properties directly; 
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually. 
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to remove a CustomShape adjustment point!
        /// </returns>
        override public bool RemovePolygonPoints(int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }

        /// <summary>
        /// Adds a polygon (list of poly point descriptor).
        /// </summary>
        /// <param name="points">The descriptor list for all polygon points of one polygon.</param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon. Default is <c>Int32.MaxValue</c></param>
        /// <param name="updateDirectly">if set to <c>true</c> the polygon lists will be updated in the shape properties directly; 
        /// otherwise you have to call <see cref="WritePointsToPolygon"/> manually. 
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c> </param>
        /// <returns>
        ///   always return <c>false</c> - it is not possible to add a new CustomShape adjustment point!
        /// </returns>
        override public bool AddPolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            return false;
        }

        /// <summary>
        /// Updates a polygon (list of poly point descriptor).
        /// If the new list is <c>null</c>, the polygon will be removed from the list.
        /// </summary>
        /// <param name="points">The descriptor list for all polygon points of one polygon or <c>null</c> if it should be removed.</param>
        /// <param name="polygonIndex">Index of the polygon if this is a poly polygon. Default is 0</param>
        /// <param name="updateDirectly">if set to <c>true</c> the polygon lists will be updated in the shape properties directly; 
        /// otherwise you have to call <see cref="WritePointsToPolygon" /> manually. 
        /// NOTICE: disable this if you manipulate a whole bunch of points. Updating all changed point at once is much faster!
        /// ATTENTION: your changes could bee overwritten by calling the <see cref="Update"/>function.
        /// Default is <c>true</c></param>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). Default is <c>false</c></param>
        /// <returns>
        ///   <c>true</c> if the polygon points could be added
        /// </returns>
        override public bool UpdatePolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {
                    try
                    {
                        _updating = true;
                        if (polygonIndex < CachedPolyPointList.Count)
                        {
                            if (points != null && CachedPolyPointList[polygonIndex].Count == points.Count)
                            {
                                CachedPolyPointList[polygonIndex] = points;
                            }
                            else
                            {
                                //CachedPolyPointList.RemoveAt(polygonIndex);
                                return false;
                            }

                            if (updateDirectly) { return WritePointsToPolygon(geometry); }
                            else { return true; }
                        }
                    }
                    finally { _updating = false; }
                }
            }
            return false;
        }

        #endregion

        #region IUpdateable

        /// <summary>
        /// Updates this instance and his related Objects. Especially the cached poly polygon points
        /// </summary>
        override public void Update()
        {
            if (_updating) return;
            try
            {
                _updating = true;
                if (IsValid())
                {
                    OoShapeObserver.LockValidation = true;
                    try
                    {
                        CachedPolyPointList = GeometryPolyPointList;
                    }
                    finally
                    {
                        OoShapeObserver.LockValidation = false;
                        _updating = false;
                    }
                }
                else
                {
                    CachedPolyPointList = null;
                }
            }
            finally
            {
                _updating = false;
            }
        }

        #endregion


        #region Point Access Overrides

        #region Array Style Access Override

        /// <summary>
        /// Gets the count of poly polygon points.
        /// </summary>
        /// <value>
        /// The count of poly polygon points.
        /// </value>
        override public int Count
        {
            get
            {
                if (IsValid() && CachedPolyPointList != null 
                    && CachedPolyPointList.Count > 0 
                    && CachedPolyPointList[0] != null)
                {
                    return CachedPolyPointList[0].Count;
                }
                return 0;
            }
        }

        /// <summary>
        /// Gets the amount of polygons inside this poly polygon.
        /// </summary>
        /// <value>
        /// The polygon count.
        /// </value>
        override public int PolygonCount
        {
            get
            {
                if (IsValid() && CachedPolyPointList != null 
                    && CachedPolyPointList.Count > 0)
                {
                    return CachedPolyPointList[0].Count;
                }
                return 0;
            }
        }

        #endregion

        #endregion

        #region Helper Functions

        /// <summary>
        /// Transforms the adjustment descriptors in a list of poly point descriptors.
        /// </summary>
        /// <param name="adjustments">The adjustments.</param>
        /// <returns>A corresponding list of simulated poly point descriptors.</returns>
        public static List<PolyPointDescriptor> TransformAdjustmentDescriptorsInPolyPointDescriptors(List<CustomShapeAdjustment> adjustments)
        {
            if (adjustments != null && adjustments.Count > 0)
            {
                List<PolyPointDescriptor> points = new List<PolyPointDescriptor>(adjustments.Count);
                foreach (CustomShapeAdjustment adj in adjustments)
                {
                    var adjPPoint = GetPointDescriptorForAdjustment(adj);
                    points.Add(adjPPoint);
                }
                return points;
            }
            return new List<PolyPointDescriptor>();
        }

        /// <summary>
        /// Generates a point descriptor for an adjustment.
        /// </summary>
        /// <param name="adj">The adjustment value.</param>
        /// <returns>An corresponding PolyPointDescriptor.</returns>
        public static PolyPointDescriptor GetPointDescriptorForAdjustment(CustomShapeAdjustment adj)
        {
            int p = (int)Math.Round(adj.Value, 0);
            var point = new PolyPointDescriptor(p, p, util.PolygonFlags.CUSTOM);
            point.Value = adj.Value;
            return point;
        }

        /// <summary>
        /// Gets an adjustment from a poly point descriptor.
        /// </summary>
        /// <param name="p">The poly point descriptor.</param>
        /// <returns>An adjustment for the poly point descriptor.</returns>
        public static CustomShapeAdjustment GetAdjustmentFromPointDescriptor(PolyPointDescriptor p)
        {
            var adj = new CustomShapeAdjustment(((double)p.Value));
            return adj;
        }


        #endregion

    }

    #endregion

}
