using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using TangramLector.OO;
using tud.mci.tangram.controller.observer;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    /// <summary>
    /// Class for manipulating OpenOffice Draw document elements 
    /// </summary>
    public partial class OpenOfficeDrawShapeManipulator
    {
        #region Last Selected Shape

        private OoShapeObserver _shape = null;
        private readonly Object _shapeLock = new Object();
        private readonly Object _shapePointLock = new Object();
        /// <summary>
        /// Gets or sets the last selected shape, which is should be modified.
        /// </summary>
        /// <value>The shape to modify.</value>
        public OoShapeObserver LastSelectedShape
        {
            get
            {
                bool fire = false;
                lock (_shapeLock)
                {
                    if (_shape != null && !_shape.Disposed && !_shape.IsValid(false))
                    {
                        unregisterFromEvents(_shape);
                        _shape = null;
                        _points = null;
                        Mode = ModificationMode.Unknown;
                        fire = true;
                        _shapeSelected = false;

                    }
                    if (fire) fire_SelectedShapeChanged(ChangeReson.Object);
                    return _shape;
                }
            }
            set
            {
                lock (_shapeLock)
                {
                    if (value != _shape)
                    {
                        if (value == null)
                        {
                            _shapeSelected = false;
                            unregisterFromEvents(_shape);
                        }
                        else
                        {
                            _shapeSelected = true;
                        }
                        _shape = value;
                        startValidationTimer();
                        registerForEvents(_shape);
                        _points = null;
                        if(LastSelectedShapePolygonPoints != null) fire_PolygonPointSelected_Reset();
                        Mode = ModificationMode.Unknown;
                    fire_SelectedShapeChanged(ChangeReson.Object);
                    }
                }
            }
        }

        #region selected shape properties handlers

        private void unregisterFromEvents(OoShapeObserver shape)
        {
            resetSelectedShapeProperties(true);
            if (shape != null && !shape.Disposed)
            {
                try
                {
                    shape.ObserverDisposing -= shape_ObserverDisposing;
                    shape.BoundRectChangeEventHandlers -= onViewOrZoomChange;
                    shape.Page.PagesObserver.ViewOrZoomChangeEventHandlers -= onViewOrZoomChange;
                }
                catch { }
            }
        }

        private void registerForEvents(OoShapeObserver shape)
        {
            if (shape != null)
            {
                try
                {
                    unregisterFromEvents(shape);
                    shape.BoundRectChangeEventHandlers += onViewOrZoomChange;
                    shape.Page.PagesObserver.ViewOrZoomChangeEventHandlers += onViewOrZoomChange;
                    shape.ObserverDisposing += shape_ObserverDisposing;
                }
                catch { }
            }
        }

        Object _selectedShapeAbsScreenBounds = null;
        /// <summary>
        /// Gets the absolute screen bounds of the selected shape in pixels.
        /// </summary>
        public Rectangle SelectedShapeAbsScreenBounds
        {
            get
            {
                if (IsShapeSelected)
                {
                    if ((_selectedShapeAbsScreenBounds == null || !(_selectedShapeAbsScreenBounds is Rectangle) || ((Rectangle)_selectedShapeAbsScreenBounds).IsEmpty) && LastSelectedShape != null)
                    {
                        _selectedShapeAbsScreenBounds = LastSelectedShape.GetAbsoluteScreenBoundsByDom();
                    }
                    return (Rectangle)_selectedShapeAbsScreenBounds;
                }
                return new Rectangle();
            }
            private set
            {
                _selectedShapeAbsScreenBounds = value;
            }
        }

        Object _selectedShapeRelScreenBounds = null;
        /// <summary>
        /// Gets the relative screen bounds of the selected shape in pixels.
        /// </summary>
        public Rectangle SelectedShapeRelScreenBounds
        {
            get
            {
                if (IsShapeSelected)
                {
                    if ((_selectedShapeRelScreenBounds == null || !(_selectedShapeRelScreenBounds is Rectangle) || ((Rectangle)_selectedShapeRelScreenBounds).IsEmpty) && LastSelectedShape != null)
                    {
                        _selectedShapeRelScreenBounds = LastSelectedShape.GetRelativeScreenBoundsByDom();
                    }
                    return (Rectangle)_selectedShapeRelScreenBounds;
                }
                return new Rectangle();
            }
            private set { _selectedShapeRelScreenBounds = value; }
        }

        byte[] _pngData = null;
        /// <summary>
        /// gets the slected shapes visual representation as png image
        /// </summary>
        public byte[] ShapePng
        {
            get
            {
                if (IsShapeSelected)
                {
                    if (_pngData == null && LastSelectedShape != null)
                    {
                        if (!(LastSelectedShape.GetShapeAsPng(out _pngData) > 0))
                        {
                            _pngData = null;
                        }
                        return _pngData;
                    }
                }
                return null;
            }
            private set { _pngData = value; }
        }

        bool _isSelcetedShapeVisible = false;
        /// <summary>
        /// Check if the currently selected shape is visible or not.
        /// </summary>
        public bool IsSelectedShapeVisible
        {
            get
            {
                if (IsShapeSelected)
                {
                    if (!_isSelcetedShapeVisible && LastSelectedShape != null)
                        _isSelcetedShapeVisible = LastSelectedShape.IsVisible();
                    return _isSelcetedShapeVisible;
                }
                return false;
            }
        }

        void shape_ObserverDisposing(object sender, EventArgs e)
        {
            if (sender == LastSelectedShape)
            {
                try
                {
                    string name = OoElementSpeaker.GetElementAudioText(LastSelectedShape);
                    play(LL.GetTrans("tangram.oomanipulation.shape.deleted", name), true);
                }
                catch (Exception ex)
                {
                }
                resetSelectedShapeProperties(true);
                LastSelectedShape = null;
            }
        }

        private void onViewOrZoomChange()
        {
            resetSelectedShapeProperties();
        }

        private void resetSelectedShapeProperties(bool silent = false)
        {
            _selectedShapeAbsScreenBounds = null;
            _selectedShapeRelScreenBounds = null;
            _pngData = null;
            _isSelcetedShapeVisible = false;
            if (IsShapeSelected && !silent) fire_SelectedShapeChanged(ChangeReson.Property);
        }

        bool _shapeSelected = false;
        /// <summary>
        /// Checked if a shape is selected for manipulation.
        /// </summary>
        public bool IsShapeSelected { get { return _shapeSelected; } private set { _shapeSelected = value; } }

        #endregion

        OoPolygonPointsObserver _points = null;
        /// <summary>
        /// Gets the  polygon points of the last selected shape.
        /// </summary>
        /// <value>
        /// The shape polygon point observer of the last selected shape if it is a freeform; otherwise <c>null</c>.
        /// </value>
        public OoPolygonPointsObserver LastSelectedShapePolygonPoints
        {
            get
            {
                lock (_shapePointLock)
                {
                    //if (_shape != null && (_points == null || _points.Shape != _shape || !_points.IsValid()))
                    //{
                    //    _points = _shape.GetPolygonPointsObserver();
                    //}
                    return _points;
                }
            }
            set
            {
                lock (_shapePointLock)
                {
                    _points = value;
                    if (_points != null)
                    {
                        Mode = ModificationMode.Move;
                    }
                    else
                    {
                        // reset mode
                        Mode = ModificationMode.Unknown;
                    }
                }
            }
        }

        #endregion

        #region validation check timer

        int _validationInterval = 10000;
        /// <summary>
        /// Gets or sets the validation interval for the last selected shape.
        /// </summary>
        /// <value>
        /// The validation interval.
        /// </value>
        public int ValidationInterval
        {
            get { return _validationInterval; }
            set { _validationInterval = value; }
        }
        Timer validationTimer = null;
        void startValidationTimer()
        {
            if (validationTimer != null) validationTimer.Dispose();
            if (_run) validationTimer = new Timer(validationTimerCallback, null, ValidationInterval, ValidationInterval);
        }

        void validationTimerCallback(object staus)
        {
            if (!_run || !IsShapeSelected ||
                LastSelectedShape == null ||
                !LastSelectedShape.IsValid(true))
            {
                validationTimer.Dispose();
            }
        }

        #endregion

    }
}
