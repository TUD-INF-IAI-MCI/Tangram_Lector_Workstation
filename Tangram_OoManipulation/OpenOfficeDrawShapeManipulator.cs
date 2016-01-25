using System;
using System.Collections.Generic;
using BrailleIO;
using BrailleIO.Interface;
using BrailleIO.Renderer;
using TangramLector.OO;
using tud.mci.tangram.controller.observer;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;

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
        private Dictionary<string, string> PatterDic = new Dictionary<string, string>();

        private IFeedbackReceiver feedbackReciever; // a receiver to output text messages or audio messages

        private int fillStyleNum = 0; //save selected style name num

        private int lineStyleNum = 0; //save selected style name iterator
        private string[] linestyleNames = new string[] { "solid", "dashed_line", "dotted_line" };

        /// <summary>
        /// translation to use for localization
        /// </summary>
        private static LL LL = new LL(Properties.Resources.Language);

        #endregion

        #region public Member

        #region last Shape

        private OoShapeObserver _shape = null;
        private readonly Object _shapeLock = new Object();
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
                    if (_shape != null && !_shape.IsValid())
                    {
                        _shape = null;
                        Mode = ModificationMode.Unknown;
                        fire = true;
                    }
                    if (fire) fire_SelectedShapeChanged();
                    return _shape;
                }
            }
            set
            {
                lock (_shapeLock)
                {
                    _shape = value;
                    Mode = ModificationMode.Unknown;
                }
                fire_SelectedShapeChanged();
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
        /// Send textual feedback to an receiver.
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
        /// Occurs when [selected shape changed].
        /// </summary>
        public event EventHandler<EventArgs> SelectedShapeChanged;

        private void fire_SelectedShapeChanged()
        {
            if (SelectedShapeChanged != null)
            {
                SelectedShapeChanged.DynamicInvoke(this, null);
            }
        }

        #endregion

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

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
}