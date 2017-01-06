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
}
