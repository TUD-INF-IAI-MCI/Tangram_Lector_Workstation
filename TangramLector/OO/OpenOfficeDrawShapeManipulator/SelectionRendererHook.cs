using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using tud.mci.tangram.controller.observer;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    /// <summary>
    /// Rendering hook for marking the current GUI selection
    /// </summary>
    /// <seealso cref="tud.mci.tangram.TangramLector.SpecializedFunctionProxies.FocusRendererHook" />
    class SelectionRendererHook : FocusRendererHook
    {
        public SelectionRendererHook() : base(true){

        }


        /// <summary>
        /// Initializes a new instance of the <see cref="FocusRendererHook"/> class.
        /// </summary>
        /// <param name="drawDashedBox">if set to <c>true</c> the box is drawn as a dashed outline box.</param>
        public SelectionRendererHook(bool drawDashedBox = false) : base(drawDashedBox)
        {
           
        }

        private OoAccessibilitySelection _lastSelection = null;
        /// <summary>
        /// The current bounding box to display
        /// </summary>
        private Rectangle _currentBoundingBox = new Rectangle(-1, -1, 0, 0);
        /// <summary>
        /// The current bounding box to display
        /// </summary>
        override public Rectangle CurrentBoundingBox
        {
            get
            {
                if (_lastSelection != null)
                {
                    _lastSelection.RefreshBounds();
                    return _lastSelection.SelectionScreenBounds;
                }

                return _currentBoundingBox;
            }
            set { _currentBoundingBox = value; }
        }

        /// <summary>
        /// Sets the current selection.
        /// To clear, simply set the selection to <c>null</c>.
        /// </summary>
        /// <param name="selection">The selection.</param>
        public void SetSelection(OoAccessibilitySelection selection)
        {
            _lastSelection = selection;
        }

    }
}
