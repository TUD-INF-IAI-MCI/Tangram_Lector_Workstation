using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.util;
using unoidl.com.sun.star.drawing;

namespace tud.mci.tangram.controller
{
    /// <summary>
    /// Factory class building specialized shape observers for different shape types.
    /// </summary>
    public static class OoShapeObserverFactory
    {
        /// <summary>
        /// Builds an Observer class for the specific shape type.
        /// </summary>
        /// <param name="s">The shape to observe.</param>
        /// <param name="page">The page the shape is placed on.</param>
        /// <returns>An shape observer class for the specific shape type.</returns>
        public static OoShapeObserver BuildShapeObserver(Object s, OoDrawPageObserver page)
        {
            return BuildShapeObserver(s as XShape, page, null);
        }

        /// <summary>
        /// Builds an Observer class for the specific shape type.
        /// </summary>
        /// <param name="s">The shape to observe.</param>
        /// <param name="page">The page the shape is placed on.</param>
        /// <param name="parent">The parent shape if the shape is part of a group.</param>
        /// <returns>
        /// An shape observer class for the specific shape type.
        /// </returns>
        public static OoShapeObserver BuildShapeObserver(Object s, OoDrawPageObserver page, OoShapeObserver parent)
        {
            return BuildShapeObserver(s as XShape, page, parent);
        }

        /// <summary>
        /// Builds an Observer class for the specific shape type.
        /// </summary>
        /// <param name="s">The shape to observe.</param>
        /// <param name="page">The page the shape is placed on.</param>
        /// <returns>An shape observer class for the specific shape type.</returns>
        internal static OoShapeObserver BuildShapeObserver(XShape s, OoDrawPageObserver page)
        {
            return BuildShapeObserver(s, page, null);
        }

        /// <summary>
        /// Builds an Observer class for the specific shape type.
        /// </summary>
        /// <param name="s">The shape to observe.</param>
        /// <param name="page">The page the shape is placed on.</param>
        /// <param name="parent">The parent shape if the shape is part of a group.</param>
        /// <returns>
        /// An shape observer class for the specific shape type.
        /// </returns>
        internal static OoShapeObserver BuildShapeObserver(XShape s, OoDrawPageObserver page, OoShapeObserver parent){
            
            if(s != null && page != null){

                if (OoUtils.ElementSupportsService(s, OO.Services.DRAW_SHAPE_CUSTOM))
                {
                    return new OoCustomShapeObserver(s, page, parent);
                }

                // normal shape
                return new OoShapeObserver(s, page, parent);
            }
            return null;
        }

    }
}
