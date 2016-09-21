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

        private readonly Object _lock = new Object();
        private List<List<PolyPointDescriptor>> _cachedPolyPointList;
        /// <summary>
        /// Gets or sets the last checked out poly point list.
        /// </summary>
        /// <value>
        /// The last checked out poly point list.
        /// </value>
        public List<List<PolyPointDescriptor>> CachedPolyPointList
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
        public List<List<PolyPointDescriptor>> GeometryPolyPointList { get { return IsValid() ? PolygonHelper.GetPolyPoints(Shape.Shape, true) : null; } }

        /// <summary>
        /// Corresponding shape.
        /// </summary>
        /// <value>The shape.</value>
        public OoShapeObserver Shape { get; private set; }

        #endregion

        #region Constructor

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

        #region Modification

        /// <summary>
        /// Returns true if the parent Shape is valid.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </returns>
        public bool IsValid()
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

        /// <summary>
        /// Writes the points to polygons properties.
        /// </summary>
        /// <param name="geometry">if set to <c>true</c> the 'Geometry' property is used (these are the untransformed coordinates of the polygon). 
        /// <returns><c>true</c> if the points could be written to the properties</returns>
        public bool WritePointsToPolygon(bool geometry = false)
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
        public bool SetPolyPointList(List<List<PolyPointDescriptor>> newPoints, bool geometry = false)
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
        public bool IsPolyPolygon()
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
        public List<PolyPointDescriptor> GetPolygonPoints(int index)
        {
            if (CachedPolyPointList != null && CachedPolyPointList.Count > index)
                return CachedPolyPointList[index];
            return new List<PolyPointDescriptor>();
        }

        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The index of the point inside the polygon description.</param>
        /// <param name="polygonIndex">Index of the polygon in a poly polygon group.</param>
        /// <returns>Point descriptor for the requested position or empty Point descriptor</returns>
        public PolyPointDescriptor GetPolyPointDescriptor(int index, int polygonIndex)
        {
            return CachedPolyPointList != null && CachedPolyPointList.Count > polygonIndex && CachedPolyPointList[polygonIndex] != null && CachedPolyPointList[polygonIndex].Count > index ?
                CachedPolyPointList[polygonIndex][index] : new PolyPointDescriptor();
        }
        /// <summary>
        /// Gets the poly point descriptor of a certain polygon at a certain index.
        /// </summary>
        /// <param name="index">The one-dimensional index of the point inside the polygon description.</param>
        /// <returns>
        /// Point descriptor for the requested position or empty Point descriptor
        /// </returns>
        public PolyPointDescriptor GetPolyPointDescriptor(int index)
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
        public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
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
        public bool UpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, bool updateDirectly = true, bool geometry = false)
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
        public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {

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
        public bool AddPolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
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
        public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
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
        public bool AddOrUpdatePolyPointDescriptor(PolyPointDescriptor ppd, int index = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
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
        /// <returns><c>true</c> if the point could be added</returns>
        public bool RemovePolyPointDescriptor(int index, int polygonIndex, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
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
        /// <returns><c>true</c> if the point could be added</returns>
        public bool RemovePolyPointDescriptor(int index, bool updateDirectly = true, bool geometry = false)
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
        public bool AddPolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = Int32.MaxValue, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null && points != null)
                {
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
        public bool UpdatePolygonPoints(List<PolyPointDescriptor> points, int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {
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
        ///   <c>true</c> if the polygon points could be added
        /// </returns>
        public bool RemovePolygonPoints(int polygonIndex = 0, bool updateDirectly = true, bool geometry = false)
        {
            if (IsValid())
            {
                if (CachedPolyPointList != null)
                {
                    if (polygonIndex < CachedPolyPointList.Count)
                    {
                        CachedPolyPointList.RemoveAt(polygonIndex);

                        if (updateDirectly) { return WritePointsToPolygon(geometry); }
                        else { return true; }
                    }
                }
            }

            return true;
        }


        #endregion

        #region IUpdateable

        bool _updating = false;
        /// <summary>
        /// Updates this instance and his related Objects. Especially the cached poly polygon points
        /// </summary>
        public void Update()
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

        //public Point TransformPointCoordinatesIntoScreenCoordinates(int x, int y)
        //{



        //}

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
        public PolyPointDescriptor this[int index]
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
        public int Count
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
        public int PolygonCount
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
        public void ResetIterator()
        {
            _iteratorIndex = -1;
        }
        /// <summary>
        /// Gets the current index of the iterator.
        /// </summary>
        /// <returns>current index of the iterator</returns>
        public int GetIteratorIndex()
        {
            return _iteratorIndex;
        }
        /// <summary>
        /// Determines whether this instance has points.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance has points; otherwise, <c>false</c>.
        /// </returns>
        public bool HasPoints() { return Count > 0; }
        /// <summary>
        /// Determines whether this instance has a next polygon point.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance has a next point; otherwise, <c>false</c>.
        /// </returns>
        public bool HasNext() { return (_iteratorIndex + 1) < Count; }
        /// <summary>
        /// Gets the next polygon point of this instance based on the internal iterator.
        /// </summary>
        /// <returns>the next polygon point</returns>
        public PolyPointDescriptor Next()
        {
            _iteratorIndex = (_iteratorIndex + 1);
            return GetPolyPointDescriptor(_iteratorIndex);
        }
        /// <summary>
        /// Get the first Point of this instance and resets the iterator.
        /// </summary>
        /// <returns>the first Point of this instance</returns>
        public PolyPointDescriptor First()
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
        public bool HasPrevious() { return _iteratorIndex > 0; }
        /// <summary>
        /// Gets the previous polygon point of this instance.
        /// </summary>
        /// <returns>the previous point</returns>
        public PolyPointDescriptor Previous()
        {
            _iteratorIndex = (_iteratorIndex - 1);
            return _iteratorIndex > 0 ? GetPolyPointDescriptor(_iteratorIndex) : new PolyPointDescriptor();
        }

        /// <summary>
        /// Get the last Point of this instance and sets the iterator to this element.
        /// </summary>
        /// <returns>the last point</returns>
        public PolyPointDescriptor Last()
        {
            _iteratorIndex = (Count - 1);
            return _iteratorIndex > 0 ? GetPolyPointDescriptor(_iteratorIndex) : new PolyPointDescriptor();
        }

        /// <summary>
        /// Gets the current selected polygon point of this instance without moving the internal iterator.
        /// </summary>
        /// <returns>the current iterated polygon point</returns>
        public PolyPointDescriptor Current(out int currentIterator)
        {
            currentIterator = _iteratorIndex;
            return GetPolyPointDescriptor(currentIterator);
        }

        #endregion

        #endregion

    }
}
