using BrailleIO;
using BrailleIO.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using TangramLector.OO;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector.Control;
using tud.mci.tangram.TangramLector.OO;
using tud.mci.tangram.TangramLector.SpecializedFunctionProxies;

namespace tud.mci.tangram.TangramLector
{
    class ImageData : AbstractSpecializedFunctionProxyBase, ILocalizable
    {
        #region string member
        LL LL = new LL(Properties.Resources.Language);

        //public const string inputField_title = "ef Titel: ";
        //public const string inputField_desc = "ef Beschreibung: ";

        #endregion  string member

        public static ImageData Instance
        {
            get { return instance; }
        }

        private static readonly ImageData instance = new ImageData();
        private volatile bool isOpen = false;

        public readonly CaretRenderer CaretHook = new CaretRenderer();
        public UnderLiningRenderer UnderLiningHook = new UnderLiningRenderer();

        #region private
        private BrailleIOMediator io { get { return BrailleIOMediator.Instance; } }
        private WindowManager wManager { get { return WindowManager.Instance; } }
        private BrailleKeyboardInput bki { get { return InteractionManager.Instance.BKI; } }
        private Property activeProperty = Property.None;
        AudioRenderer audioRenderer = AudioRenderer.Instance;
        private Dictionary<string, BrailleIOViewRange> detailViewDic = new Dictionary<string, BrailleIOViewRange>();
        private BrailleIOScreen brailleIoScreen;
        private volatile SaveDialog saveMenu = SaveDialog.NonActive;
        private volatile Property selectedPropertyMenuItem = Property.None;

        private GUI_Menu gui_Menu;
        private string viewRangename = "imageDataView";
        private Dictionary<string, string> propertiesDic = new Dictionary<string, string>()
            {
                {"title",""},
                {"description",""}
            };

        /// <summary>
        /// Input string after last Braille input changed event.
        /// </summary>
        private string _lastInputString = String.Empty;
        /// <summary>
        /// Shape whose title and description are edited.
        /// </summary>
        private OoShapeObserver _shape = null;

        #endregion

        private ImageData()
            : base(100)
        {
            // events            
            bki.BrailleInputChanged += new EventHandler<InputChangedEventArgs>(brailleInput_Changed);
            bki.BrailleInputCaretMoved += new EventHandler<CaretMovedEvent>(bki_BrailleInputCaretMoved);
            ScriptFunctionProxy.Instance.BrailleKeyboardCommand += new EventHandler<BrailleKeyboardCommandEventArgs>(brailleKeybordCommandEvent);
            InteractionManager.Instance.GesturePerformed += new EventHandler<GestureEventArgs>(im_GesturePerformedEvent);
            CaretHook.BKI = bki;
        }

        #region Detail Region Set Up

        public void ShowImageDetailView(string detailViewName, int heightOffSet, int widthOffSet, Property property, bool loadFromShape = false)
        {
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] open title + description dialog");
            string speak = "";
            gui_Menu = new GUI_Menu(); //init GUI MENU  
            //create entry for dic
            createDetailView(detailViewName, 0, 48, 120, 12, true);
            detailViewDic[detailViewName].SetVisibility(true);
            //load data from selected Shape     
            if (loadFromShape)
            {
                loadImageData();
            }
            //as default title is selected and show
            activeProperty = property;
            selectedPropertyMenuItem = property;
            setContent(detailViewName, property);
            switch (property)
            {
                case Property.Title:
                    speak = LL.GetTrans("tangram.lector.image_data.title");
                    break;
                case Property.Description:
                    speak = LL.GetTrans("tangram.lector.image_data.desc");
                    break;
            }

            audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.image_data.generic_property_value", speak, propertiesDic[activeProperty.ToString().ToLower()]));
            isOpen = true;
        }

        public const string TITLE_DESC_VIEW_NAME = "imageData";
        public const string TITLE_DESC_SAVE_VIEW_NAME = "saveMenu";

        private void createDetailView(string detailViewName, int left, int top, int width, int height, bool scrollable)
        {
            try
            {
                if (!detailViewDic.ContainsKey(detailViewName))
                {
                    detailViewDic.Add(detailViewName, null);
                }
                brailleIoScreen = wManager.GetVisibleScreen();
                detailViewDic[detailViewName] = new BrailleIOViewRange(left, top, width, height);
                brailleIoScreen.AddViewRange(viewRangename, detailViewDic[detailViewName]);
                viewRangename = detailViewDic[detailViewName].Name;
                detailViewDic[detailViewName].SetVisibility(true);
                detailViewDic[detailViewName].SetZIndex(99);
                detailViewDic[detailViewName].SetBorder(1, 0, 0);
                detailViewDic[detailViewName].SetMargin(1, 0, 0);
                detailViewDic[detailViewName].SetPadding(1, 0, 0);
                detailViewDic[detailViewName].HasBorder = true;
                detailViewDic[detailViewName].ShowScrollbars = scrollable;
                detailViewDic[detailViewName].SetText("");
                var render = detailViewDic[detailViewName].ContentRender;
                if (render != null && render is IBrailleIOHookableRenderer)
                {
                    if (detailViewName == TITLE_DESC_VIEW_NAME)
                    {
                        ((IBrailleIOHookableRenderer)render).RegisterHook(CaretHook);
                    }
                    else if (detailViewName == TITLE_DESC_SAVE_VIEW_NAME)
                    {
                        ((IBrailleIOHookableRenderer)render).RegisterHook(UnderLiningHook);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, this, "[ERROR] while changing the height of the detail region", ex);
            }
        }

        /// <summary>
        /// shows Menu item in DetailView
        /// </summary>
        private void showMenuDetailView()
        {
            createDetailView(TITLE_DESC_SAVE_VIEW_NAME, 0, 53, 120, 7, false);
            detailViewDic[TITLE_DESC_SAVE_VIEW_NAME].SetMargin(0); // necessary to see the horizontal scrollbar slider
            setContent(TITLE_DESC_SAVE_VIEW_NAME, Property.None);
        }

        /// <summary>
        /// load existing images data, title and description, and store them
        /// to PropertyDic
        /// </summary>
        private void loadImageData()
        {
            _shape = getSelectedShape();
            if (_shape != null)
            {
                propertiesDic["title"] = _shape.Title;
                propertiesDic["description"] = _shape.Description;
            }
        }

        /// <summary>
        /// change the content of the detailView
        /// </summary>
        /// <param name="detailViewName">Name of the detail view.</param>
        /// <param name="property">The property.</param>
        private void setContent(string detailViewName, Property property)
        {
            string enteredText = ""; //load description or title of the image
            string content = ""; // text shown on device in detailView
            switch (detailViewName)
            {
                case TITLE_DESC_VIEW_NAME:
                    {
                        switch (property)
                        {
                            case Property.Title:
                                content = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.title") + ":";
                                enteredText = propertiesDic["title"];
                                break;

                            case Property.Description:
                                content = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.desc") + ":";
                                enteredText = propertiesDic["description"];
                                break;
                            default:
                                break;
                        }
                        bki.SetInput(enteredText);
                    }
                    break;
                case TITLE_DESC_SAVE_VIEW_NAME:
                    content = LL.GetTrans("tangram.lector.image_data.save")
                    + " | " + LL.GetTrans("tangram.lector.image_data.not_save")
                    + " | " + LL.GetTrans("tangram.lector.image_data.cancel");
                    break;
                default:
                    break;
            }

            BrailleIOViewRange vr = detailViewDic[detailViewName];
            setScrollableViewRangeTextContent(vr, content + enteredText);
        }

        #endregion

        #region Scrolling

        /// <summary>
        /// Set text value of a view range and scrolls to the bottom if necessary.
        /// </summary>
        /// <param name="vr">View range to fill with content.</param>
        /// <param name="text">Text content.</param>
        private void setScrollableViewRangeTextContent(BrailleIOViewRange vr, string text)
        {
            if (vr != null)
            {
                vr.SetText(text);
                if (text.Length <= 40) return; // TODO: only hack because at this point contentHeight is that of last shown content --> if the new content is very short, than the scrolling would show it out of view range

                int contentHeight = vr.ContentHeight - 1; // TODO: ignore last line in content renderer instead of "-1"
                if (vr.ShowScrollbars && (contentHeight > vr.ContentBox.Height)) // TODO: adaption also necessary if content is smaller than offset position
                {
                    int diff = vr.ContentBox.Height - contentHeight;
                    scrollViewRangeTo(vr, diff);
                }
            }
        }

        /// <summary>
        /// Set offset position of view range to given position.
        /// </summary>
        /// <param name="vr">View range to set new offset.</param>
        /// <param name="y_pos">new offset position.</param>
        private void scrollViewRangeTo(BrailleIOViewRange vr, int y_pos)
        {
            if (vr != null && y_pos <= 0)
            {
                if (y_pos >= (vr.ContentBox.Height - vr.ContentHeight)) vr.SetYOffset(y_pos);
                else vr.SetYOffset(0);
            }
        }

        /// <summary>
        /// Add the given scroll step to the current offset position of view range to realize a scrolling.
        /// </summary>
        /// <param name="vr">View range to scroll.</param>
        /// <param name="scrollStep">Steps to scroll the view range.</param>
        private void scrollViewRange(BrailleIOViewRange vr, int scrollStep)
        {
            if (vr != null)
            {
                int newPos = vr.OffsetPosition.Y + scrollStep;
                if (newPos <= 0 && newPos >= (vr.ContentBox.Height - vr.ContentHeight)) vr.SetYOffset(newPos);
            }
        }

        #endregion

        #region SaveDialog Carusel

        private readonly int _maxSaveMode = Enum.GetValues(typeof(SaveDialog)).Cast<int>().Max();

        private void rotateThroughSaveDialog(int direction)
        {
            object val = Convert.ChangeType(saveMenu, saveMenu.GetTypeCode());
            int m = Convert.ToInt32(val);

            // should not rotate to UNKOWN = 0
            if (direction == -1)
            {
                if (m == 1)
                {
                    m = _maxSaveMode;
                }
                else
                {
                    m--;
                }
            }
            else
            {
                m = Math.Max(1, (++m) % (_maxSaveMode + 1));
            }
            saveMenu = (SaveDialog)Enum.ToObject(typeof(SaveDialog), m);
            comunicateSaveDialogSwitch(saveMenu);
        }

        private void comunicateSaveDialogSwitch(SaveDialog saveDialog)
        {
            String command = "";
            switch (saveDialog)
            {
                case SaveDialog.Save:
                    command = LL.GetTrans("tangram.lector.image_data.save");
                    break;
                case SaveDialog.Abort:
                    command = LL.GetTrans("tangram.lector.image_data.cancel");
                    break;
                case SaveDialog.NotSave:
                    command = LL.GetTrans("tangram.lector.image_data.not_save");
                    break;
            }
            UnderLiningHook.selectedString = command;
            AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.selected", command));
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] save dialog switch to " + command);
        }

        #endregion

        # region Properties Carusel

        private readonly int _maxProperty = Enum.GetValues(typeof(Property)).Cast<int>().Max();

        private void rotateThroughProperties()
        {
            object val = Convert.ChangeType(selectedPropertyMenuItem, selectedPropertyMenuItem.GetTypeCode());
            int m = Convert.ToInt32(val);
            // should not rotate to None = 0
            m = Math.Max(1, (++m) % (_maxProperty + 1));
            var previousSelectedMenuItem = selectedPropertyMenuItem;
            selectedPropertyMenuItem = (Property)Enum.ToObject(typeof(Property), m);
            comunicatePropertiesSwitch(selectedPropertyMenuItem, previousSelectedMenuItem);
        }

        private void comunicatePropertiesSwitch(Property property, Property previousProperty)
        {
            String command = getAudioForProperty(property);
            activeProperty = property;
            string previousContent = removeInputStr(property, previousProperty); // content of detailView contain inputField
            SaveBrailleInput(previousProperty, previousContent);
            setContent(TITLE_DESC_VIEW_NAME, property);
            // Test, whether property has content, if yes speak it
            string content = propertiesDic[property.ToString().ToLower()];
            if (content.Length != 0)
            {
                //command += " ist " + content;
                command = LL.GetTrans("tangram.lector.image_data.generic_property_value", command, content);
            }
            else
            {
                //command += " eingeben";
                command = LL.GetTrans("tangram.lector.image_data.enter_property", command);
            }
            AudioRenderer.Instance.PlaySoundImmediately(command);
            Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] show " + property.ToString() + ": " + content);
        }

        #endregion

        /// <summary>
        /// Save and Set content of detailView depending on property
        /// </summary>
        /// <param name="property">The property.</param>
        /// <param name="content">The content.</param>
        private void SaveBrailleInput(Property property, string content)
        {
            propertiesDic[property.ToString().ToLower()] = content;
        }

        /// <summary>
        /// Gets the audio for property.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <returns><c>null</c> or String of Property</returns>
        private string getAudioForProperty(Property property)
        {
            switch (property)
            {
                case Property.Description:
                    return LL.GetTrans("tangram.lector.image_data.desc");
                case Property.Title:
                    return LL.GetTrans("tangram.lector.image_data.title");
                default:
                    return LL.GetTrans("tangram.lector.image_data.title");
            }

        }

        /// <summary>
        /// remove input label for property from detailView content
        /// </summary>
        /// <param name="property">The property.</param>
        /// <param name="previousProperty">The previous property.</param>
        /// <returns>string without input label</returns>
        private string removeInputStr(Property property, Property previousProperty)
        {
            string result = "";
            if (previousProperty != Property.None)
            {
                result = detailViewDic[TITLE_DESC_VIEW_NAME].GetText().Replace(getRemoveString(previousProperty), "");
            }
            else
            {
                result = detailViewDic[TITLE_DESC_VIEW_NAME].GetText().Replace(getRemoveString(property), "");
            }
            return result;
        }

        /// <summary>
        /// Gets the remove string.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <returns>string of property</returns>
        private string getRemoveString(Property property)
        {
            string removeStr = "";
            switch (property)
            {
                case Property.None:
                    removeStr = getRemoveString(selectedPropertyMenuItem);
                    break;
                case Property.Title:
                    removeStr = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.title") + ":";
                    break;
                case Property.Description:
                    removeStr = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.desc") + ":";
                    break;
                default:

                    break;
            }
            return removeStr;
        }

        /// <summary>
        /// close editing Dialog and save Changes if saveChanges is true
        /// </summary>
        /// <param name="saveChanges">if set to <c>true</c> [save changes].</param>
        private void CloseImageDialog(bool saveChanges)
        {
            string audioOutput = "";
            if (saveChanges)
            {
                var shape = getSelectedShape();

                if (shape != null)
                {
                    if (shape.Equals(_shape)) // original selected shape is still selected
                    {
                        SaveBrailleInput(selectedPropertyMenuItem, removeInputStr(selectedPropertyMenuItem, Property.None)); //save actual bki because Dictionary content is written to imageData
                        shape.Description = propertiesDic["description"];
                        shape.Title = propertiesDic["title"];
                        audioOutput = LL.GetTrans("tangram.lector.image_data.cahnges.saved");
                    }
                }
            }
            else
            {
                audioOutput += LL.GetTrans("tangram.lector.image_data.cahnges.not_saved");
            }

            _shape = null;
            saveMenu = SaveDialog.NonActive;
            gui_Menu.resetGuiMenu();
            audioOutput += " " + LL.GetTrans("tangram.lector.image_data.dialog_closed");
            audioRenderer.PlaySoundImmediately(audioOutput);
            isOpen = false;

            //change Interaction mode
            InteractionManager.Instance.ChangeMode(InteractionMode.Normal);
            OoElementSpeaker.PlayElement(getSelectedShape(), LL.GetTrans("tangram.lector.oo_observer.selected"));

            this.Active = false;
            foreach (var key in detailViewDic.Keys)
            {
                if (detailViewDic[key] != null)
                {
                    detailViewDic[key].SetVisibility(false);
                }
            }
            WindowManager.Instance.SetTopRegionContent(WindowManager.Instance.GetMainscreenTitle());


            //////////////////////////////////////////////////////////////////////////
            //TODO: start blinking again
            //////////////////////////////////////////////////////////////////////////

        }

        /// <summary>
        /// Get selected shape from OpenOffice Observer.
        /// </summary>
        /// <returns></returns>
        private OoShapeObserver getSelectedShape()
        {
            var con = OO.OoConnector.Instance;
            if (con.Observer != null)
            {
                return con.Observer.GetLastSelectedShape();
            }
            return null;
        }

        /// <summary>
        /// Handling of new selection while data of old selected shape is still loaded in title+desc dialog.
        /// </summary>
        /// <param name="newSelectedShape">new selected shape</param>
        public void NewSelectionHandling(object newSelectedShape)
        {
            audioRenderer.PlayWaveImmediately(StandardSounds.Error);
            if (_shape != null) audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.image_data.save_element_before", _shape.Name));
            else audioRenderer.PlaySoundImmediately(LL.GetTrans("tangram.lector.image_data.save_before"));

            // open save dialog
            if (saveMenu == SaveDialog.NonActive)
            {
                BrailleIODevice dev = null;
                if (io != null && io.AdapterManager != null && io.AdapterManager.ActiveAdapter != null)
                    dev = io.AdapterManager.ActiveAdapter.Device;

                // inject pressing "Enter" key
                ImageData.Instance.fire_ButtonCombinationReleased(null,
                        new ButtonReleasedEventArgs(
                            dev,
                            BrailleIO_DeviceButton.None,
                            BrailleIO_BrailleKeyboardButton.None,
                            null,
                            new List<string>(),
                            BrailleIO_DeviceButton.Enter,
                            BrailleIO_BrailleKeyboardButton.None,
                            null,
                            new List<string>() { "crc" })
                        );
            }
            // speak current selected option if save dialog is already open 
            else if (saveMenu == SaveDialog.Save) AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.selected", LL.GetTrans("tangram.lector.image_data.save")));
            else if (saveMenu == SaveDialog.NotSave) AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.selected", LL.GetTrans("tangram.lector.image_data.not_save")));
            else if (saveMenu == SaveDialog.Abort) AudioRenderer.Instance.PlaySoundImmediately(LL.GetTrans("tangram.lector.oo_observer.selected", LL.GetTrans("tangram.lector.image_data.cancel")));
        }

        #region Events

        void brailleInput_Changed(object sender, InputChangedEventArgs args)
        {
            if (isOpen)
            {
                // append braille input to shown text
                string label = "";
                switch (activeProperty)
                {
                    case Property.Title:
                        label = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.title") + ": ";
                        break;
                    case Property.Description:
                        label = LL.GetTrans("tangram.lector.image_data.ef") + " " + LL.GetTrans("tangram.lector.image_data.desc") + ": ";
                        break;
                    default:
                        break;
                }

                _lastInputString = args.Input;
                BrailleIOViewRange vr = detailViewDic[TITLE_DESC_VIEW_NAME];
                setScrollableViewRangeTextContent(vr, label + args.Input);
            }
        }

        void im_GesturePerformedEvent(object sender, GestureEventArgs e)
        {
            if (e != null && e.Gesture != null)
            {
                if (e.Gesture.Name.Equals("tap") && _shape != null)
                {
                    var shape = getSelectedShape();
                    if (!shape.Equals(_shape))
                    {
                        // after saving old data, new data has to be loaded
                        loadImageData();
                        setContent(TITLE_DESC_VIEW_NAME, Property.Title);
                    }
                }
            }
        }

        internal void fire_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        {
            this.im_ButtonCombinationReleased(sender, e);
        }

        protected override void im_ButtonCombinationReleased(object sender, ButtonReleasedEventArgs e)
        { }

        protected override void im_FunctionCall(object sender, FunctionCallInteractionEventArgs e)
        {
            if (Active && e != null && !String.IsNullOrEmpty(e.Function) && !e.AreButtonsPressed())
            {
                switch (e.Function)
                {
                    case "confirm":

                        switch (saveMenu)
                        {
                            case SaveDialog.NonActive:
                                //first call make save dialog visible
                                //saveDialog = SaveDialog.Active;
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] open save dialog");
                                saveMenu = SaveDialog.Save;
                                showMenuDetailView();
                                UnderLiningHook.selectedString = LL.GetTrans("tangram.lector.image_data.save");
                                SaveBrailleInput(activeProperty, removeInputStr(activeProperty, Property.None));
                                audioRenderer.PlaySoundImmediately(
                                    LL.GetTrans("tangram.lector.image_data.save_dialog.opend")
                                    + " " + LL.GetTrans("tangram.lector.image_data.save_dialog.entry_selected"
                                    , LL.GetTrans("tangram.lector.image_data.save")));
                                break;
                            case SaveDialog.Abort:
                                //close save dialog
                                saveMenu = SaveDialog.NonActive;
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] save menu aborted");

                                detailViewDic[TITLE_DESC_SAVE_VIEW_NAME].SetVisibility(false);
                                ShowImageDetailView(TITLE_DESC_VIEW_NAME, 15, 0, activeProperty);
                                break;
                            case SaveDialog.NotSave:
                                // quit dialog without saving changes
                                CloseImageDialog(false);
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] changes were not saved, dialog closed");
                                break;
                            case SaveDialog.Save:
                                // quit dialog and save changes
                                CloseImageDialog(true);
                                Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] changes were saved, dialog closed");
                                break;
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        break;


                    // scrolling through save, not saving and abort
                    case "changeRight":
                        //right scrolling through saveCommand
                        if (saveMenu != SaveDialog.NonActive)
                        {
                            rotateThroughSaveDialog(1);
                        }
                        else // editing mode --> move caret to right
                        {
                            bki.MoveCaret(1);
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] move caret to right");
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        break;

                    case "changeLeft":
                        if (saveMenu != SaveDialog.NonActive)
                        {
                            rotateThroughSaveDialog(-1);
                        }
                        else // editing mode --> move caret to left
                        {
                            bki.MoveCaret(-1);
                            Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] move caret to left");
                        }
                        e.Cancel = true;
                        e.Handled = true;
                        break;

                    case "changeUp":
                        rotateThroughProperties();
                        e.Cancel = true;
                        e.Handled = true;
                        break;

                    case "changeDown":
                        rotateThroughProperties();
                        e.Cancel = true;
                        e.Handled = true;
                        break;

                    default:
                        break;
                }
            }
        }


        void brailleKeybordCommandEvent(object sender, BrailleKeyboardCommandEventArgs e)
        {
            BrailleIOViewRange vr = detailViewDic["imageData"];
            switch (e.Command)
            {
                case BrailleCommand.CursorUp:
                    scrollViewRange(vr, 5);
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] cursor up");
                    break;
                case BrailleCommand.CursorDown:
                    scrollViewRange(vr, -5);
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] cursor down");
                    break;
                case BrailleCommand.Pos1:
                    scrollViewRangeTo(vr, 0);
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] cursor to pos 1");
                    break;
                case BrailleCommand.End:
                    int diff = vr.ContentBox.Height - vr.ContentHeight + 1; // TODO: "+1" has to be removed when content height is adapted
                    scrollViewRangeTo(vr, diff);
                    Logger.Instance.Log(LogPriority.MIDDLE, this, "[SHAPE EDIT DIALOG] cursor to end");
                    break;
                case BrailleCommand.Unkown:
                    if (e.Character.Equals("l+lr")) OoConnector.Instance.Observer.HighlightBrailleFocusOnScreen();
                    break;
                default:
                    break;
            }
        }

        void bki_BrailleInputCaretMoved(object sender, CaretMovedEvent e)
        {
            if (CaretHook != null)
            {
                CaretHook.MoveCaret(e.CaretPosition);
            }
        }

        #endregion

        private String Linebreaks(string str)
        {
            string breaks = "";

            if (str.Length > 40)
            {
                breaks = LL.GetTrans("tangram.lector.image_data.linebreak_necessary");
            }
            else
            {
                breaks = str;
            }
            return breaks;

        }

        #region ILocalizable

        void ILocalizable.SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

    }

    public class UnderLiningRenderer : IBailleIORendererHook
    {
        public string selectedString { get; set; }
        public void PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams) { }

        public void PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            if (view != null && content != null && result != null)
            {
                if (!String.IsNullOrEmpty(selectedString) && content.ToString().ToLower().Contains(selectedString.ToLower()))
                {
                    int start = content.ToString().ToLower().IndexOf(selectedString.ToLower());
                    //check
                    int end = start + selectedString.Length;
                    start = start * 3; // each letter has 3 dot * 5
                    end = end * 3;
                    int underlinePos = 4; // underline 4 & 8

                    for (int i = start; i < end; i++)
                    {
                        if (result.GetLength(0) > underlinePos)
                        {
                            if (result.GetLength(1) > i)
                            {
                                result[underlinePos, i] = true;
                            }
                        }
                    }
                }
            }
        }
    }

    public class CaretRenderer : IBailleIORendererHook
    {
        public BrailleKeyboardInput BKI = null;
        private int caretPosition = 0;
        private Boolean boolBlink = false;
        // not necessary till now
        public CaretRenderer()
        {
            var blinkTimer = WindowManager.blinkTimer;
            if (blinkTimer != null)
            {
                blinkTimer.HalfTick += new EventHandler<EventArgs>(blinkTimer_HalfTick);
            }
        }

        void blinkTimer_HalfTick(object sender, EventArgs e)
        {
            boolBlink = !boolBlink;
        }

        public void PreRenderHook(ref IViewBoxModel view, ref object content, params object[] additionalParams)
        {
            if (boolBlink && content != null && content is String)
            {
                string contentStr = (string)content;

                int caretPos = getCaretOff(contentStr) + BKI.Caret;
                string character = caretPos < contentStr.Length ? contentStr.Substring(caretPos, 1) : " ";
                var dotPattern = ScriptFunctionProxy.Instance.BrailleKeyboard.GetDotsForChar(character);
                if (dotPattern == null)
                {
                    dotPattern = "";
                }
                if (!dotPattern.Contains("8"))
                {
                    if (!dotPattern.Contains("7"))
                    {
                        dotPattern += "7";
                    }
                    //append 8 at the end of the string
                    dotPattern += "8";
                }
                else
                {
                    if (!dotPattern.Contains("7"))
                    {
                        dotPattern.Replace("8", "78");
                    }
                }
                string nc = ScriptFunctionProxy.Instance.BrailleKeyboard.GetCharFromDots(dotPattern); // char with underline dots 7 & 8

                if (caretPos >= contentStr.Length)
                {
                    contentStr += nc;
                }
                else
                {
                    contentStr = contentStr.Substring(0, caretPos) + nc + contentStr.Substring(caretPos + 1);
                }
                content = contentStr;
            }
        }

        public void PostRenderHook(IViewBoxModel view, object content, ref bool[,] result, params object[] additionalParams)
        {
            return;
        }

        public void MoveCaret(int posX)
        {
            this.caretPosition = posX;
        }

        private int getCaretOff(string text)
        {
            return text.Length - BKI.Input.Length;
        }
    }

    #region enum

    public enum SaveDialog
    {
        NonActive = 0,
        Save,
        NotSave,
        Abort
    }

    public enum Property
    {
        None = 0,
        Title,
        Description

    }

    #endregion
}


