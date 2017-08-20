using System;
using System.Collections.Generic;
using TangramLector.OO;
using tud.mci.tangram.controller.observer;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;
using System.Threading.Tasks;
using System.Drawing;
using System.Threading;

namespace tud.mci.tangram.TangramLector.SpecializedFunctionProxies
{
    /// <summary>
    /// Class for manipulating OpenOffice Draw document elements 
    /// </summary>
    public partial class OpenOfficeDrawShapeManipulator : AbstractSpecializedFunctionProxyBase, ILocalizable
    {

        #region Members

        #region private Member

        IOoDrawConnection OoConnection;
        IZoomProvider Zoomable;

        readonly InteractionManager interactionManager = InteractionManager.Instance;
        private SortedDictionary<string, string> PatterDic = new SortedDictionary<string, string>();

        private IFeedbackReceiver feedbackReciever; // a receiver to output text messages or audio messages

        private int fillStyleNum = -1; //save selected style name num

        private int lineStyleNum = 0; //save selected style name iterator
        private string[] linestyleNames = new string[] { "solid", "dashed_line", "dotted_line", "white_line" };

        static int _whiteColorInt = tud.mci.tangram.util.OoUtils.ConvertToColorInt(System.Drawing.Color.White);
        static int _blackColorInt = tud.mci.tangram.util.OoUtils.ConvertToColorInt(System.Drawing.Color.Black);

        /// <summary>
        /// translation to use for localization
        /// </summary>
        private static LL LL = new LL(Properties.Resources.Language);

        #endregion

        #region public Member

        #region last Shape

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
                    if (_shape != null && !_shape.IsValid(false))
                    {
                        unregisterFromEvents(_shape);
                        _shape = null;
                        _points = null;
                        Mode = ModificationMode.Unknown;
                        fire = true;
                        _shapeSelected = false;

                    }
                    if (fire) fire_SelectedShapeChanged();
                    return _shape;
                }
            }
            set
            {
                lock (_shapeLock)
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
                    registerForEvents(_shape);
                    _points = null;
                    fire_PolygonPointSelected_Reset();
                    Mode = ModificationMode.Unknown;
                }
                fire_SelectedShapeChanged();
            }
        }

        #region selected shape propertie handlers

        private void unregisterFromEvents(OoShapeObserver shape)
        {
            resetSelectedShapeProperties();
            if (shape != null)
            {
                try
                {
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
            private set { _selectedShapeAbsScreenBounds = value; }
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

        private void onViewOrZoomChange()
        {
            resetSelectedShapeProperties();
        }

        private void resetSelectedShapeProperties()
        {
            _selectedShapeAbsScreenBounds = null;
            _selectedShapeRelScreenBounds = null;
            _pngData = null;
            _isSelcetedShapeVisible = false;
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
                }
            }
        }

        #endregion

        private bool _active = false;
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="OpenOfficeDrawShapeManipulator"/> is active.
        /// Indicator if this <see cref="IInteractionContextProxy" /> instance decide to be responsible 
        /// for handling input events form the function proxy.
        /// </summary>
        /// <value><c>true</c> if active; otherwise, <c>false</c>.</value>
        public override bool Active
        {
            get { return _active; }
            set
            {
                if (value != _active)
                {
                    _active = value;
                    if (_active) { fireActivatedEvent(); }
                    else { fireDeactivatedEvent(); }
                }
            }
        }

        /// <summary>
        /// Mode for interpretation of modification requests
        /// </summary>
        public volatile ModificationMode Mode = ModificationMode.Unknown;

        #endregion

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenOfficeDrawShapeManipulator"/> class.
        /// </summary>
        /// <param name="drawConnection">A connection to the Draw application. ATTENTION: doen not be null.</param>
        /// <param name="fbReciever">The fb reciever.</param>
        public OpenOfficeDrawShapeManipulator(IOoDrawConnection drawConnection, IZoomProvider zommable, IFeedbackReceiver fbReciever = null)
            : base(10)
        {

            feedbackReciever = fbReciever;
            Zoomable = zommable;
            if (drawConnection == null)
            {
                throw new ArgumentNullException("drawConnection", "Without any connection to the Draw application no Manipulation is possible.");
            }
            OoConnection = drawConnection;
            initializePatters();
            loadConfiguration();
        }

        #endregion

        /// <summary>
        /// Sets the object to receive the feedback if some should be given.
        /// </summary>
        /// <param name="fbReceiver"></param>
        public void SetFeedbackReceiver(IFeedbackReceiver fbReceiver)
        {
            this.feedbackReciever = fbReceiver;
        }

        /// <summary>
        /// Send textual feedback to an receiver.
        /// </summary>
        /// <param name="text">the textual feedback to send.</param>
        /// <returns><c>true</c> if the feedback could be sent, otherwise <c>false</c>.</returns>
        private bool sentTextFeedback(string text)
        {
            if (feedbackReciever != null)
            {
                feedbackReciever.ReceiveTextualFeedback(text);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Send Quickinfo textual feedback to an receiver.
        /// </summary>
        /// <param name="text">the textual feedback to send.</param>
        /// <returns><c>true</c> if the feedback could be sent, otherwise <c>false</c>.</returns>
        private bool sentTextNotification(string text)
        {
            if (feedbackReciever != null)
            {
                feedbackReciever.ReceiveTextualNotification(text);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Send auditory feedback to an receiver.
        /// </summary>
        /// <param name="text">the text to send and interpret as audio feedback.</param>
        /// <returns><c>true</c> if the feedback could be sent, otherwise <c>false</c>.</returns>
        private bool sentAudioFeedback(string audio)
        {
            if (feedbackReciever != null)
            {
                feedbackReciever.ReceiveAuditoryFeedback(audio);
                return true;
            }
            return false;
        }

        void sendDetailInfo(string text)
        {
            sentTextNotification(text);
        }

        #region Audio

        /// <summary>
        /// Plays a sound indication an edit.
        /// </summary>
        private static void playEdit() { AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Info); }
        /// <summary>
        /// Plays a sound indicating an error has occurred.
        /// </summary>
        private static void playError() { AudioRenderer.Instance.PlayWaveImmediately(StandardSounds.Error); }
        /// <summary>
        /// Plays a specified text string.
        /// </summary>
        /// <param name="text">The text to play as text to speach.</param>
        private static void play(String text) { AudioRenderer.Instance.PlaySoundImmediately(text); }

        /// <summary>
        /// Brings the last selected shape to the audio output.
        /// </summary>
        private void sayLastSelectedShape()
        {
            if (LastSelectedShape != null)
            {
                OoElementSpeaker.PlayElementImmediately(LastSelectedShape, LL.GetTrans("tangram.oomanipulation.selected"));
            }
            else
            {
                play(LL.GetTrans("tangram.oomanipulation.no_element_selected"));
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs when the selected shape changed.
        /// </summary>
        public event EventHandler<EventArgs> SelectedShapeChanged;
        /// <summary>
        /// Occurs when a polygon point was selected.
        /// </summary>
        public event EventHandler<PolygonPointSelectedEventArgs> PolygonPointSelected;


        private void fire_SelectedShapeChanged()
        {
            if (SelectedShapeChanged != null)
            {
                Task t = new Task(new Action(() => { try { SelectedShapeChanged.Invoke(this, null); } catch { } }));
                Thread.Sleep(10);
                t.Start();
            }
        }

        private void fire_PolygonPointSelected(OoPolygonPointsObserver ppobs, util.PolyPointDescriptor point)
        {
            if (PolygonPointSelected != null)
            {
                Task t = new Task(new Action(() => { try { PolygonPointSelected.Invoke(this, new PolygonPointSelectedEventArgs(ppobs, point)); } catch { } }));
                t.Start();
            }
        }

        private void fire_PolygonPointSelected_Reset()
        {
            fire_PolygonPointSelected(null, new util.PolyPointDescriptor());
        }


        #endregion

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion


        #region Configuration

        /// <summary>
        /// Loads the app config configurations if set.
        /// </summary>
        internal void loadConfiguration()
        {
            try
            {
                var config = System.Configuration.ConfigurationManager.AppSettings;
                if (config != null && config.Count > 0)
                {
                    setMinimumShapeSizeFromConfig(config);
                }
            }
            catch { }
        }

        #region MinimumShapeSize

        /// <summary>
        /// The sound volume configuration key to search for in the app.config file.
        /// </summary>
        internal const String SHAPE_SIZE_FACTOR_CONFIG_KEY = "Tangram_MinimumShapeSize";
        /// <summary>
        /// Sets the MinimumShapeSize if the corresponding key <see cref="SHAPE_SIZE_FACTOR_CONFIG_KEY"/> 
        /// was set in the app.config of the current running application.
        /// </summary>
        /// <param name="config">The loaded configuration app settings.</param>
        internal void setMinimumShapeSizeFromConfig(System.Collections.Specialized.NameValueCollection config)
        {
            try
            {
                if (config != null && config.Count > 0)
                {
                    var value = config[SHAPE_SIZE_FACTOR_CONFIG_KEY];
                    if (value != null)
                    {
                        int shapesize = Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture);
                        MinimumShapeSize = shapesize;
                    }
                }
            }
            catch { }
        }

        #endregion

        #endregion

    }

    #region Enums

    /// <summary>
    /// Mode for interpretation of modification requests
    /// </summary>
    public enum ModificationMode
    {
        /// <summary>
        /// 
        /// </summary>
        Unknown,
        /// <summary>
        /// move elements
        /// </summary>
        Move,
        /// <summary>
        /// scale elements
        /// </summary>
        Scale,
        /// <summary>
        /// rotate elements
        /// </summary>
        Rotate,
        /// <summary>
        /// fill style
        /// </summary>
        Fill,
        /// <summary>
        /// line style
        /// </summary>
        Line

    }

    #endregion

    #region Event Args

    /// <summary>
    /// Event arguments for reporting that a polygon point is selected.
    /// </summary>
    /// <seealso cref="System.EventArgs" />
    public class PolygonPointSelectedEventArgs : EventArgs
    {
        /// <summary>
        /// The polygon points observer. Can be <c>null</c>
        /// </summary>
        public readonly OoPolygonPointsObserver PolygonPoints;
        /// <summary>
        /// The current selected point. Can be an empty one.
        /// </summary>
        public readonly tud.mci.tangram.util.PolyPointDescriptor Point;

        /// <summary>
        /// Initializes a new instance of the <see cref="PolygonPointSelectedEventArgs"/> class.
        /// </summary>
        /// <param name="ppobs">polygon points observer.</param>
        /// <param name="point">current selected point.</param>
        public PolygonPointSelectedEventArgs(OoPolygonPointsObserver ppobs, tud.mci.tangram.util.PolyPointDescriptor point)
        {
            PolygonPoints = ppobs;
            Point = point;
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="PolygonPointSelectedEventArgs"/> class with empty Members. 
        /// Should only be used to reset the listeners.
        /// </summary>
        public PolygonPointSelectedEventArgs()
        {
            PolygonPoints = null;
            Point = new util.PolyPointDescriptor();
        }
    }

    #endregion

}