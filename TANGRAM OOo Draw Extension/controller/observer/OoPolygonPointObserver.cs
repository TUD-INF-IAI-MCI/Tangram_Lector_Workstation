using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;

namespace tud.mci.tangram.controller.observer
{
    public class OoPolygonPointObserver
    {
        #region Member

        public const string POLYGON_POINTS_PROPERTY_NAME = "Polygon";

        System.Drawing.Point _p;
        /// <summary>
        /// Gets or sets the position of the point.
        /// </summary>
        /// <value>The point.</value>
        public System.Drawing.Point P
        {
            get { return _p; }
            private set { _p = value; bool success = setPosition(); }
        }
        /// <summary>
        /// Corresponding shape.
        /// </summary>
        /// <value>The shape.</value>
        public OoShapeObserver Shape { get; private set; }

        int index = 0;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OoPolygonPointObserver"/> class.
        /// </summary>
        /// <param name="shape">The coresponding shape.</param>
        /// <param name="p">The position of the point.</param>
        internal OoPolygonPointObserver(OoShapeObserver shape, unoidl.com.sun.star.awt.Point p, int index)
            : this(shape, new System.Drawing.Point(p.X, p.Y), index)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="OoPolygonPointObserver"/> class.
        /// </summary>
        /// <param name="shape">The coresponding shape.</param>
        /// <param name="p">The position of the point.</param>
        public OoPolygonPointObserver(OoShapeObserver shape, System.Drawing.Point p, int index)
        {
            Shape = shape;
            P = p;
            this.index = index;
        }

        #endregion

        #region Modification

        /// <summary>
        /// Sets the horizontal position.
        /// </summary>
        /// <param name="x">The horizontal position.</param>
        public void SetX(int x)
        {
            P = new System.Drawing.Point(x, P.Y);
        }

        /// <summary>
        /// Moves the point in the horizontal direction.
        /// </summary>
        /// <param name="xOffset">The horizontal movement added to the current horizontal (X) position.</param>
        public void MoveHorizontal(int xOffset)
        {
            SetX(P.X + xOffset);
        }

        /// <summary>
        /// Sets the vertical position.
        /// </summary>
        /// <param name="x">The vertical position.</param>
        public void SetY(int y)
        {
            P = new System.Drawing.Point(P.X, y);
        }

        /// <summary>
        /// Moves the point in the vertical direction.
        /// </summary>
        /// <param name="xOffset">The vertical movement added to the current vertical (Y) position.</param>
        public void MoveVertical(int yOffset)
        {
            SetY(P.Y + yOffset);
        }

        /// <summary>
        /// Moves the Point to the given position.
        /// </summary>
        /// <param name="p">The new Position.</param>
        public void MoveTo(System.Drawing.Point p)
        {
            P = p;
        }
        /// <summary>
        /// Moves the Point to the given position.
        /// </summary>
        /// <param name="x">The new horizontal position.</param>
        /// <param name="y">The new vertical position.</param>
        public void MoveTo(int x, int y)
        {
            MoveTo(new System.Drawing.Point(x, y));
        }

        bool setPosition()
        {
            if (Shape != null)
            {
                return Shape.SetPolygonPoint(this, index);
            }
            return false;
        }

        #endregion
    }
}
