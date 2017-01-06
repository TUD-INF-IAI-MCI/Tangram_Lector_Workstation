using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.drawing;

namespace tud.mci.tangram.controller.observer
{
    class OoCustomShapeObserver : OoShapeObserver, INameBuilder, IUpdateable
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
            getCustomShapeHandles();
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

        void updateCustomShapeProperties()
        {
            var csprop = OoUtils.GetProperty(Shape, "CustomShapeGeometry");
            if (csprop != null)
            {
               _cachedCustonProperties = OoUtils.GetPropertyvalueDictionary(csprop as PropertyValue[]);
            }
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


        List<CustomShapeHandle> getCustomShapeHandles()
        {
            List<CustomShapeHandle> hndls = new List<CustomShapeHandle>();
            var props = getCustomProperties(false);
            if (props.ContainsKey("Handles"))
            {
                var handles = props["Handles"].Value.Value as PropertyValue[][];
                if (handles != null && handles.Length > 0)
                {
                    foreach (PropertyValue[] handle in handles)
                    {
                        CustomShapeHandle csh = new CustomShapeHandle(handle);
                        hndls.Add(csh);
                    }
                }
            }
            return hndls;
        }


        #endregion

        #region IUpdateable

        new void Update()
        {
            base.Update();
            updateCustomShapeProperties();
        }

        #endregion


        #region Helper



        #endregion

    }

    public struct CustomShapeHandle
    {
        int x;
        /// <summary>
        /// Gets or sets the x position of the handle.
        /// </summary>
        /// <value>
        /// The x position.
        /// </value>
        public int X
        {
            get { return x; }
            set {
                x = Math.Min(RangeXMaximum, Math.Max(RangeXMinimum, value)); 
            }
        }

        int y;
        /// <summary>
        /// Gets or sets the y position of the handle.
        /// </summary>
        /// <value>
        /// The y position.
        /// </value>
        public int Y
        {
            get { return y; }
            set {
                y = Math.Min(RangeYMaximum, Math.Max(RangeYMinimum, value)); 
            }
        }

        /// <summary>
        /// The minimum y value
        /// </summary>
        public readonly int RangeYMinimum;
        /// <summary>
        /// The maximum y value
        /// </summary>
        public readonly int RangeYMaximum;

        /// <summary>
        /// The minimum x value
        /// </summary>
        public readonly int RangeXMinimum;
        /// <summary>
        /// The maximum y value
        /// </summary>
        public readonly int RangeXMaximum;

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomShapeHandle"/> struct.
        /// </summary>
        /// <param name="vals">The vals.</param>
        internal CustomShapeHandle(PropertyValue[] vals)
        {
            var props = OoUtils.GetPropertyvalueDictionary(vals);

            // position
            x = 0;
            y = 0;
            if (props.ContainsKey("Position"))
            {
                var pos = props["Position"].Value.Value;
                if (pos != null && pos is EnhancedCustomShapeParameterPair)
                {
                    var xVal = ((EnhancedCustomShapeParameterPair)pos).First.Value.Value;
                   x = (int) xVal;
                   var yVal = ((EnhancedCustomShapeParameterPair)pos).Second.Value.Value;
                   y = (int) yVal;
                }
            }

            // x Range
            RangeXMinimum = x;
            RangeXMaximum = x;
            if (props.ContainsKey("RangeXMinimum") && props.ContainsKey("RangeXMaximum"))
            {
                var xMin = props["RangeXMinimum"].Value.Value;
                if (xMin != null && xMin is EnhancedCustomShapeParameter)
                {
                    RangeXMinimum = (int)(((EnhancedCustomShapeParameter)xMin).Value.Value);
                }
                var xMax = props["RangeXMaximum"].Value.Value;
                if (xMax != null && xMax is EnhancedCustomShapeParameter)
                {
                    RangeXMaximum = (int)(((EnhancedCustomShapeParameter)xMax).Value.Value);
                }
            }

            // y Range
            RangeYMinimum = y;
            RangeYMaximum = y;
            if (props.ContainsKey("RangeYMinimum") && props.ContainsKey("RangeYMaximum"))
            {
                var yMin = props["RangeYMinimum"].Value.Value;
                if (yMin != null && yMin is EnhancedCustomShapeParameter)
                {
                    RangeXMinimum = (int)(((EnhancedCustomShapeParameter)yMin).Value.Value);
                }
                var yMax = props["RangeYMaximum"].Value.Value;
                if (yMax != null && yMax is EnhancedCustomShapeParameter)
                {
                    RangeXMaximum = (int)(((EnhancedCustomShapeParameter)yMax).Value.Value);
                }
            }
        }

        #endregion

        #region Requests




        #endregion


    }

}
