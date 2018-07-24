using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using tud.mci.tangram.TangramLector.OO;

namespace tud.mci.tangram.TangramLector.BrailleIO.Model
{
    /// <summary>
    /// A wrapping class for information about the DRAW application. 
    /// </summary>
    public class OoDrawModel
    {
        #region Members

        #region public
        ScreenObserver _obs = null;
        /// <summary>
        /// Gets or sets the screen observer for application windows or the desktop.
        /// </summary>
        /// <value>
        /// The screen observer.
        /// </value>
        public ScreenObserver ScreenObserver
        {
            get { return _obs; }
            protected set
            {
                if (value != _obs)
                {
                    unregisterFromScreenObserverEvents(_obs);
                    _obs = value;
                    registerToScreenObserverEvents(_obs);
                }
            }
        }

        Image _lstCapt = null;
        /// <summary>
        /// Gets or sets the last screen capturing result image.
        /// </summary>
        /// <value>
        /// The last screen capturing image.
        /// </value>
        public Image LastScreenCapturing
        {
            get { return _lstCapt.Clone() as Image; }
            set
            {
                if (_lstCapt != value)
                {
                    var tmp = _lstCapt;
                    _lstCapt = value;
                    tmp.Dispose();
                }
            }
        }

        /// <summary>
        /// Gets the global DRAW observer.
        /// </summary>
        /// <value>
        /// The DRAW observer.
        /// </value>
        public OoObserver OoObserver
        {
            get
            {
                if (OoConnector.Instance != null) return OoConnector.Instance.Observer;
                return null;
            }
        }

        #endregion

        #endregion

        #region Screen Observer

        private void registerToScreenObserverEvents(TangramLector.ScreenObserver _obs)
        {
            if (_obs != null)
            {
                _obs.Changed += _obs_Changed;
            }
        }

        private void unregisterFromScreenObserverEvents(TangramLector.ScreenObserver _obs)
        {
            try
            {
                if (_obs != null)
                {
                    _obs.Changed -= _obs_Changed;
                }
            }
            catch (Exception) { }
        }

        void _obs_Changed(object sender, CaptureChangedEventArgs e)
        {
            if (e != null) LastScreenCapturing = e.Img.Clone() as Image;
        }

        #endregion

    }
}
