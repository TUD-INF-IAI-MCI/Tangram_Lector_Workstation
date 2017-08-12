using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using tud.mci.tangram.util;

namespace tud.mci.tangram.controller.observer
{
    /// <summary>
    /// Wrapper class for handling poly polygon and poly Bezier function points.
    /// </summary>
    /// <seealso cref="tud.mci.tangram.models.Interfaces.IUpdateable" />
    /// <seealso cref="System.Collections.IEnumerable" />
    public class OoPolygonPointsObserver : IUpdateable, IEnumerable
    {
        #region Member
        /// <summary>
        /// Gets a value indicating whether this instance is empty.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is empty; otherwise, <c>false</c>.
        /// </value>
        public bool IsEmpty { get; private set; }

        volatile bool _updating = false;
        private readonly Object _lock = new Object();
        private List<List<PolyPointDescriptor>> _cachedPolyPointList;
        /// <summary>
        /// Gets or sets the last checked out poly point list.
        /// </summary>
        /// <value>
        /// The last checked out poly point list.
        /// </value>
        virtual public List<List<PolyPointDescriptor>> CachedPolyPointList
        {
            get { lock (_lock) return _cachedPolyPointList; }
            protected set { lock (_lock)_cachedPolyPointList = value; }
        }

        /// <summary>
        /// Gets the geometry poly point list.
        /// </summary>
        /// <value>
        /// The geometry poly point list.
        /// </value>
        virtual public List<List<PolyPointDescriptor>> GeometryPolyPointList { get { return IsValid() ? PolygonHelper.GetPolyPoints(Shape.Shape, true) : null; } }

        /// <summary>
        /// Corresponding shape.
        /// </summary>
        /// <value>The shape.</value>
        virtual public OoShapeObserver Shape { get; protected set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoPolygonPointsObserver"/> class.
        /// ATTENTION: is only for derivation and internal use. 
        /// This constructor leaves the class as a non-functional object.
        /// </summary>
        protected OoPolygonPointsObserver() { IsEmpty = true; }

        /// <summary>
        /// Initializes a new instance of the <see cref="OoPolygonPointsObserver" /> class.
        /// </summary>
        /// <param name="shape">The corresponding shape.</param>
        /// <exception cref="ArgumentNullException">shape</exception>
        /// <exception cref="ArgumentException">shape must be a polygon or a bezier curve;shape</exception>
        public OoPolygonPointsObserver(OoShapeObserver shape)
        {
            if (shape == null) throw new ArgumentNullException("shape");
            if (!PolygonHelper.IsFreeform(shape.Shape)) throw new ArgumentException("shape must be a polygon or a bezier curve", "shape");

            Shape = shape;
            Shape.BoundRectChangeEventHandlers += Shape_BoundRectChangeEventHandlers;
            Shape.ObserverDisposing += Shape_ObserverDisposing;
            Update();
            IsEmpty = false;
        }

        void Shape_ObserverDisposing(object sender, EventArgs e)
        {
            Shape = null;
            CachedPolyPointList = null;
        }

        void Shape_BoundRectChangeEventHandlers()
        {
            //System.Diagnostics.Debug.WriteLine("[P] ------ Bound rect changed observed in Polygon Points observer!!!");
            Update();
        }

        #endregion

        #region Tests

        /// <summary>
        /// Returns true if the parent Shape is valid.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool IsValid()
        {
            try
            {
                return Shape != null && Shape.Shape != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        bool _isClosed = false;
        /// <summary>
        /// Determines whether this Freeform is a closed one or an open one.
        /// </summary>
        /// <param name="cached">if set to <c>true</c> the previously cached result will be returned.</param>
        /// <returns>
        ///   <c>true</c> if the Freeform is a closed one; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool IsClosed(bool cached = true)
        {
            if (IsValid() && !cached) _isClosed = PolygonHelper.IsClosedFreeForm(Shape.Shape);
            return _isClosed;
        }

        bool _isBezier = false;
        /// <summary>
        /// Determines whether this Freeform is a Bezier defined form or not.
        /// </summary>
        /// <param name="cached">if set to <c>true</c> the previously cached result will be returned.</param>
        /// <returns>
        ///   <c>true</c> if the Freeform is a Bezier defined one; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool IsBezier(bool cached = true)
        {
            if (IsValid() && !cached)
            {
                if (OoUtils.ElementSupportsService(Shape.Shape, OO.Services.DRAW_POLY_POLYGON_BEZIER_DESCRIPTOR))
                    _isBezier = true;
            }
            return _isBezier;
        }


        /// <summary>
        /// The tolerated distance between points to be taken as equal.
        /// </summary>
        const int POINT_EQUAL_TOLERANCE = 2;

        bool _firstIsLast = false;
        /// <summary>
        /// Check if first and last point are equal. This is e.g. the case for Beziere curves.
        /// </summary>
        /// <param name="cached">if set to <c>true</c> [cached].</param>
        /// <returns></returns>
        virtual public bool FirstPointEqualsLastPoint(bool cached = false)
        {
            // if (IsClosed())
            // {
                if (IsBezier()) return _firstIsLast = true;
                else if (!cached)
                {
                    if (this.Count > 1)
                    {
                        var firts = this[0];
                        var last = this[this.Count - 1];

                        if (firts.Flag == PolygonFlags.NORMAL && last.Flag == PolygonFlags.NORMAL &&
                            Math.Abs(firts.X - last.X) <= POINT_EQUAL_TOLERANCE &&
                            Math.Abs(firts.Y - last.Y) <= POINT_EQUAL_TOLERANCE)
                        {
                            return _firstIsLast = true;
                        }
                    }
                    else return _firstIsLast = false;
                }
            //}
            //else _firstIsLast = false;
            return _firstIsLast;
        }

        #endregion

        #region Modification

        /// <summary>
        /// Writes the points to polygons properties.
        /// </summary>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). 
        /// <returns><c>true</c> if the points could be written to the properties</returns>
        virtual public bool WritePointsToPolygon(bool geometry = false)
        {
            if (IsValid())
            {
                OoShapeObserver.LockValidation = true;
                try
                {
                    bool success = PolygonHelper.SetPolyPoints(Shape.Shape, CachedPolyPointList, geometry, Shape.GetDocument());
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
        virtual public bool SetPolyPointList(List<List<PolyPointDescriptor>> newPoints, bool geometry = false)
        {
            if (IsValid())
            {
                OoShapeObserver.LockValidation = true;
                try
                {
                    bool success = PolygonHelper.SetPolyPoints(Shape.Shape, newPoints, geometry, Shape.GetDocument());
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
        /// Determines whether the polygon description is a group of several polygons.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this is polygon descriptor is a group of more than one polygon; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool IsPolyPolygon()
        {
            return CachedPolyPointList != null && CachedPolyPointList.Count > 1;
        }

        #endregion

        #region Point Handling

        /// <summary>
        /// Gets the polygon points of a certain polygon of the poly polygon descriptor.
        /// </summary>
        /// <param name="index">The index of the polygon in the poly polygon descriptor.</param>
        /// <returns>the list of polygon point descriptors or an empty list if there is no polygon at the requested index.</returns>
        virtual public List<PolyPointDescriptor> GetPolygonPoints(int index)
        {
            try
            {
                _updating = true;
                if (CachedPolyPointList != null && CachedPolyPointList.Count > index)
                    return CachedPolyPointList[index];
                return new List<PolyPointDescriptor>();
            }
            finally { _updating = false; }
        }

        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The index of the point inside the polygon description.</param>
        /// <param name="polygonIndex">Index of the polygon in a poly polygon group.</param>
        /// <returns>Point descriptor for the requested position or empty Point descriptor</returns>
        virtual public PolyPointDescriptor GetPolyPointDescriptor(int index, int polygonIndex)
        {
            try
            {
                _updating = true;
                return CachedPolyPointList != null && CachedPolyPointList.Count > polygonIndex && CachedPolyPointList[polygonIndex] != null && CachedPolyPointList[polygonIndex].Count > index ?
                    CachedPolyPointList[polygonIndex][index] : new PolyPointDescriptor();
            }
            finally { _updating = false; }
        }
        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The one-dimensional index of the point inside the polygon description.</param>
        /// <returns>
        /// Point descriptor for the requested position or empty Point descriptor
        /// </returns>
        virtual public PolyPointDescriptor GetPolyPointDescriptor(int index)
        {
            int i, pi;
            Transform1DindexTo2DpolyPolygonIndex(index, out i, out pi);
            return GetPolyPointDescriptor(i, pi);
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
        virtual public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
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
        virtual public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, bool updateDirectly = true, bool geometry = false)
        {
            int i, pi;
            Transform1DindexTo2DpolyPolygonIndex(index, out i, out pi);
            return UpdatePolyPointDescriptor(ppd, i, pi, updateDirectly, geometry);
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
        ///   <c>true</c> if the point could be added
        /// </returns>
        virtual public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {
                    try
                    {
                        _updating = true;
                        if (CachedPolyPointList.Count > polygonIndex)
                        {
                            if (CachedPolyPointList[polygonIndex] == null) CachedPolyPointList[polygonIndex] = new List<PolyPointDescriptor>();
                            if (index < CachedPolyPointList[polygonIndex].Count) { CachedPolyPointList[polygonIndex].Insert(Math.Max(index, 0), ppd); }
                            else { CachedPolyPointList[polygonIndex].Add(ppd); }
                        }
                        else
                        {
                            CachedPolyPointList.Add(
                                new List<PolyPointDescriptor>(){
                                ppd
                            }
                            );
                        }

                        if (updateDirectly) { return WritePointsToPolygon(geometry); }
                        else { return true; }
                    }
                    finally { _updating = false; }
                }

            }
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
        ///   <c>true</c> if the point could be added
        /// </returns>
        virtual public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            int i, pi;
            Transform1DindexTo2DpolyPolygonIndex(index, out i, out pi);
            return AddPolyPointDescriptor(ppd, i, pi, updateDirectly, geometry);
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
        virtual public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
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
                        else
                        {
                            return AddPolyPointDescriptor(ppd, index, polygonIndex, updateDirectly, geometry);
                        }
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
        virtual public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            int i, pi;
            Transform1DindexTo2DpolyPolygonIndex(index, out i, out pi);
            return AddOrUpdatePolyPointDescriptor(ppd, i, pi, updateDirectly, geometry);
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
        /// <returns><c>true</c> if the point could be removed</returns>
        virtual public bool RemovePolyPointDescriptor(int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                try
                {
                    _updating = true;
                    if (CachedPolyPointList != null && CachedPolyPointList.Count > polygonIndex)
                    {
                        if (CachedPolyPointList[polygonIndex] != null && CachedPolyPointList[polygonIndex].Count > index)
                        {
                            CachedPolyPointList[polygonIndex].RemoveAt(index);

                            if (updateDirectly) { return WritePointsToPolygon(geometry); }
                            else { return true; }
                        }
                    }
                }
                finally { _updating = false; }
            }

            return true;
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
        /// <returns><c>true</c> if the point could be removed</returns>
        virtual public bool RemovePolyPointDescriptor(int index, bool updateDirectly = true, bool geometry = false)
        {
            int i, pi;
            Transform1DindexTo2DpolyPolygonIndex(index, out i, out pi);
            return RemovePolyPointDescriptor(i, pi, updateDirectly, geometry);
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
        /// <returns><c>true</c> if the polygon points could be added</returns>
        virtual public bool AddPolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null && points != null)
                {
                    try
                    {
                        _updating = true;
                        if (polygonIndex < CachedPolyPointList.Count)
                        {
                            CachedPolyPointList.Insert(Math.Max(0, polygonIndex), points);
                        }
                        else
                        {
                            CachedPolyPointList.Add(points);
                        }

                        if (updateDirectly) { return WritePointsToPolygon(geometry); }
                        else { return true; }
                    }
                    finally { _updating = false; }
                }
            }
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
        virtual public bool UpdatePolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
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
                            if (points != null)
                            {
                                CachedPolyPointList[polygonIndex] = points;
                            }
                            else
                            {
                                CachedPolyPointList.RemoveAt(polygonIndex);
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
        ///   <c>true</c> if the polygon points could be removed
        /// </returns>
        virtual public bool RemovePolygonPoints(int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {
                    try
                    {
                        _updating = true;
                        int pointIndx, plyIndx;
                        Transform1DindexTo2DpolyPolygonIndex(polygonIndex, out pointIndx, out plyIndx);

                        if (plyIndx < CachedPolyPointList.Count && pointIndx < CachedPolyPointList[plyIndx].Count)
                        {
                            CachedPolyPointList[plyIndx].RemoveAt(pointIndx);

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
        virtual public void Update()
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
                        CachedPolyPointList = PolygonHelper.GetPolyPoints(Shape.Shape);
                        IsClosed(false);
                        IsBezier(false);
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

        #region Helper Functions

        /// <summary>
        /// Transform1s a one-dimensional index into a two-dimensional index of this poly polygon descriptor. So a continuous indexing can be realized.
        /// </summary>
        /// <param name="index">The one-dimensional index.</param>
        /// <param name="pointIndex">Index of the point inside a polygon.</param>
        /// <param name="polygonIndex">Index of the polygon inside the poly polygon list.</param>
        public void Transform1DindexTo2DpolyPolygonIndex(int index, out int pointIndex, out int polygonIndex)
        {
            pointIndex = index;
            polygonIndex = 0;

            if (CachedPolyPointList != null)
            {
                while (polygonIndex < CachedPolyPointList.Count)
                {
                    if (pointIndex < CachedPolyPointList[polygonIndex].Count)
                    {
                        break;
                    }
                    else
                    {
                        pointIndex -= CachedPolyPointList[polygonIndex].Count;
                        polygonIndex++;
                    }
                }
            }
        }

        /// <summary>
        /// Transforms a point, which is in document coordinates, into pixel coordinates on screen,
        /// taking into account the current zoom level of the drawing application.
        /// </summary>
        /// <param name="polyPointDescriptor">The polygon point to transform.</param>
        /// <returns>
        /// The estimated pixel position of the point on the screen.
        /// </returns>
        public Point TransformPointCoordinatesIntoScreenCoordinates(PolyPointDescriptor polyPointDescriptor)
        {
            return TransformPointCoordinatesIntoScreenCoordinates(polyPointDescriptor.X, polyPointDescriptor.Y);
        }
        /// <summary>
        /// Transforms a point, which is in document coordinates, into pixel coordinates on screen, 
        /// taking into account the current zoom level of the drawing application.
        /// </summary>
        /// <param name="x">The x coordinate to transform.</param>
        /// <param name="y">The y coordinate to transform.</param>
        /// <returns>The estimated pixel position of the point on the screen.</returns>
        public Point TransformPointCoordinatesIntoScreenCoordinates(int x, int y)
        {
            Point p = new Point(x, y);

            try
            {
                if (IsValid()
                    && Shape.Page != null
                    && Shape.Page.PagesObserver != null)
                {
                    Point offset = Shape.Page.PagesObserver.ViewOffset;
                    int currentZoom = Shape.Page.PagesObserver.ZoomValue;
                    p.X = OoDrawUtils.ConvertToPixel(x - offset.X, currentZoom, OoDrawPagesObserver.PixelPerMeterX);
                    p.Y = OoDrawUtils.ConvertToPixel(y - offset.Y, currentZoom, OoDrawPagesObserver.PixelPerMeterY);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.ALWAYS, this, "[ERROR] when transforming polygon points" + ex);
            }

            return p;
        }

        /// <summary>
        /// Exports a string representation of polygons.
        /// Can be uses to import it again.
        /// </summary>
        /// <returns>A multi-line string of a polypolygon points definitions.</returns>
        public String ToExportString()
        {
            List<List<PolyPointDescriptor>> polys = CachedPolyPointList;
            String exp = PolygonHelper.BuildExportString(polys);
            return exp;
        }


        #endregion

        #region Point Access

        #region IEnumerable

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>An <see cref="IEnumerator"/> object that can be used to iterate through the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            if (!IsValid() || CachedPolyPointList == null)
                yield return null;

            for (int index = 0; index < CachedPolyPointList.Count; index++)
            {
                if (CachedPolyPointList[index] == null) continue;

                for (int index2 = 0; index2 < CachedPolyPointList[index].Count; index2++)
                {
                    // Yield each day of the week.
                    yield return CachedPolyPointList[index][index2];
                }
            }
        }

        #endregion

        #region Array Style Access

        /// <summary>
        /// Gets or sets the <see cref="PolyPointDescriptor" /> at the specified index.
        /// </summary>
        /// <value>
        /// The <see cref="PolyPointDescriptor" />.
        /// NOTICE: do not use the set functionality if you manipulate a whole bunch of points. 
        /// Updating all changed point at once is much faster! So use <see cref="AddOrUpdatePolyPointDescriptor"/>.
        /// </value>
        /// <param name="index">The index.</param>
        /// <returns>
        /// the <see cref="PolyPointDescriptor" /> at the specified one-dimensional index
        /// </returns>
        virtual public PolyPointDescriptor this[int index]
        {
            get
            {
                return GetPolyPointDescriptor(index);
            }
            set
            {
                AddOrUpdatePolyPointDescriptor(value, index);
            }
        }

        /// <summary>
        /// Gets the count of poly polygon points.
        /// </summary>
        /// <value>
        /// The count of poly polygon points.
        /// </value>
        virtual public int Count
        {
            get
            {
                int c = 0;
                if (IsValid() && CachedPolyPointList != null)
                {
                    foreach (var item in CachedPolyPointList)
                    {
                        if (item != null)
                            c += item.Count;
                    }
                }
                return c;
            }
        }

        /// <summary>
        /// Gets the amount of polygons inside this poly polygon.
        /// </summary>
        /// <value>
        /// The polygon count.
        /// </value>
        virtual public int PolygonCount
        {
            get
            {
                int c = 0;
                if (IsValid() && CachedPolyPointList != null)
                {
                    c = CachedPolyPointList.Count;
                }
                return c;
            }
        }

        #endregion

        #region Iterator Like Access

        int _iteratorIndex = -1;

        /// <summary>
        /// Resets the iterator.
        /// </summary>
        virtual public void ResetIterator()
        {
            _iteratorIndex = -1;
        }
        /// <summary>
        /// Gets the current index of the iterator.
        /// </summary>
        /// <returns>current index of the iterator</returns>
        virtual public int GetIteratorIndex()
        {
            return _iteratorIndex;
        }
        /// <summary>
        /// Determines whether this instance has points.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance has points; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool HasPoints() { return Count > 0; }
        /// <summary>
        /// Determines whether this instance has a next polygon point.
        /// </summary>
        /// <param name="ignoreLastDuplicate">if set to <c>true</c> to ignore the 
        /// last point of a closed Bezier shape because its exactly the same as 
        /// the first.</param>
        /// <returns>
        ///   <c>true</c> if this instance has a next point; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool HasNext(bool ignoreLastDuplicate) {
            if (ignoreLastDuplicate && _iteratorIndex == Count -2 && FirstPointEqualsLastPoint()) {
                return false;
            }
            else return (_iteratorIndex + 1) < Count; 
        }
        /// <summary>
        /// Determines whether this instance has a next polygon point.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance has a next point; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool HasNext() { return (_iteratorIndex + 1) < Count; }
        /// <summary>
        /// Gets the next polygon point of this instance based on the internal iterator.
        /// </summary>
        /// <returns>the next polygon point</returns>
        virtual public PolyPointDescriptor Next()
        {
            _iteratorIndex = (_iteratorIndex + 1);
            return GetPolyPointDescriptor(_iteratorIndex);
        }
        /// <summary>
        /// Get the first Point of this instance and resets the iterator.
        /// </summary>
        /// <returns>the first Point of this instance</returns>
        virtual public PolyPointDescriptor First()
        {
            ResetIterator();
            return Next();
        }

        /// <summary>
        /// Determines whether this instance has a previous point.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance has previous point; otherwise, <c>false</c>.
        /// </returns>
        virtual public bool HasPrevious() { return _iteratorIndex > 0; }
        /// <summary>
        /// Gets the previous polygon point of this instance.
        /// </summary>
        /// <returns>the previous point</returns>
        virtual public PolyPointDescriptor Previous()
        {
            _iteratorIndex = (_iteratorIndex - 1);
            return _iteratorIndex > 0 ? GetPolyPointDescriptor(_iteratorIndex) : new PolyPointDescriptor();
        }


        /// <summary>
        /// Get the last Point of this instance and sets the iterator to this element.
        /// </summary>
        /// <returns>the last point</returns>
        virtual public PolyPointDescriptor Last(bool ignoreLastDuplicate)
        {
            if (!ignoreLastDuplicate || Count < 2) return Last();

            _iteratorIndex = (Count - 2);
            if (HasNext(ignoreLastDuplicate))
                return Next();

            return _iteratorIndex > 0 ? GetPolyPointDescriptor(_iteratorIndex) : new PolyPointDescriptor();
        }

        /// <summary>
        /// Get the last Point of this instance and sets the iterator to this element.
        /// </summary>
        /// <returns>the last point</returns>
        virtual public PolyPointDescriptor Last()
        {
            _iteratorIndex = (Count - 1);
            return _iteratorIndex > 0 ? GetPolyPointDescriptor(_iteratorIndex) : new PolyPointDescriptor();
        }

        /// <summary>
        /// Gets the current selected polygon point of this instance without moving the internal iterator.
        /// </summary>
        /// <returns>the current iterated polygon point</returns>
        virtual public PolyPointDescriptor Current(out int currentIterator)
        {
            currentIterator = _iteratorIndex;
            return GetPolyPointDescriptor(currentIterator);
        }

        #endregion

        #endregion




    }
}
