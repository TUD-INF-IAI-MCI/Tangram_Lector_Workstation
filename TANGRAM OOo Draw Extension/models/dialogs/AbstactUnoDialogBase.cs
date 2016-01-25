using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.frame;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;

namespace tud.mci.tangram.models.dialogs
{
    public abstract class AbstactUnoDialogBase
    {
        #region Members

        private XComponentContext _mxContext = null;
        protected XComponentContext MXContext
        {
            get
            {
                return _mxContext;
            }
            set
            {
                _mxContext = value;
            }
        }

        private XMultiComponentFactory _mxMcf;
        protected XMultiComponentFactory MXMcf
        {
            get
            {
                return _mxMcf;
            }
            set
            {
                _mxMcf = value;
            }
        }

        private object _mxMsfDialogModel;
        /// <summary>
        /// Gets or sets the MX MSF dialog model.
        /// DEPRECATED: don't use this - because in libre office this is allways null
        /// </summary>
        /// <value>The MX MSF dialog model.</value>
        protected object MXMsfDialogModel
        {
            get
            {
                return _mxMsfDialogModel;
            }
            set
            {
                _mxMsfDialogModel = value;
            }
        }

        private XControlModel _mxModel;
        protected XControlModel MXModel
        {
            get
            {
                if (_mxModel == null && MXDialogControl != null)
                {
                    _mxModel = MXDialogControl.getModel();
                    //util.Debug.GetAllInterfacesOfObject(_mxModel);
                }
                return _mxModel;
            }
            set
            {
                _mxModel = value;
            }
        }

        private XNameContainer _mxDlgModelNameContainer;
        protected XNameContainer MXDlgModelNameContainer
        {
            get
            {
                return _mxDlgModelNameContainer;
            }
            set
            {
                _mxDlgModelNameContainer = value;
            }
        }

        private XControlContainer _mxDlgContainer;
        protected XControlContainer MXDlgContainer
        {
            get
            {
                return _mxDlgContainer;
            }
            set
            {
                _mxDlgContainer = value;
            }
        }

        private XControl _mxDialogControl;
        protected XControl MXDialogControl
        {
            get
            {
                return _mxDialogControl;
            }
            set
            {
                _mxDialogControl = value;
                MXDialog = _mxDialogControl as XDialog;
            }
        }

        private XDialog _xDialog;
        protected XDialog MXDialog
        {
            get
            {
                return _xDialog;
            }
            set
            {
                _xDialog = value;
            }
        }

        private XReschedule _mxReschedule;
        protected XReschedule MXReschedule
        {
            get
            {
                return _mxReschedule;
            }
            set
            {
                _mxReschedule = value;
            }
        }

        private XWindowPeer _mxWindowPeer = null;
        protected XWindowPeer MXWindowPeer
        {
            get
            {
                return _mxWindowPeer;
            }
            set
            {
                _mxWindowPeer = value;
            }
        }

        private XTopWindow _mxTopWindow = null;
        protected XTopWindow MXTopWindow
        {
            get
            {
                return _mxTopWindow;
            }
            set
            {
                _mxTopWindow = value;
            }
        }

        private XFrame _mxFrame = null;
        protected XFrame MXFrame
        {
            get
            {
                return _mxFrame;
            }
            set
            {
                _mxFrame = value;
            }
        }

        private XComponent _mxComponent = null;
        //private XComponentContext xComponentContext;
        //private XComponentContext xContext;
        //private XMultiComponentFactory xMultiComponentFactory;
        //private AbstactUnoDialogBase abstactUnoDialogBase;
        protected XComponent MXComponent
        {
            get
            {
                return _mxComponent;
            }
            set
            {
                _mxComponent = value;
            }
        }

        #endregion

        #region CTOR

        public AbstactUnoDialogBase() : this(OO.GetContext()) { }
        public AbstactUnoDialogBase(XComponentContext xContext) : this(xContext, OO.GetMultiComponentFactory(xContext)) { }
        public AbstactUnoDialogBase(XComponentContext xContext, XMultiComponentFactory xMcf)
        {
            MXContext = xContext;
            MXMcf = xMcf;
            CreateDialog();
        }

        #endregion

        #region Initalisation

        /// <summary>
        /// Creates the dialog model and container.
        /// Variables MXMsfDialogModel, MXDlgModelNameContainer, MXDialogControl, MXDlgContainer and MXTopWindow are created.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, BackgroundColor, Closeable, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Moveable, Sizeable, TextColor, TextLineColor, Title
        /// </summary>
        protected virtual void CreateDialog() { CreateDialog(MXMcf); }
        /// <summary>
        /// Creates the dialog model and container.
        /// Variables MXMsfDialogModel, MXDlgModelNameContainer, MXDialogControl, MXDlgContainer and MXTopWindow are created.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, BackgroundColor, Closeable, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Moveable, Sizeable, TextColor, TextLineColor, Title
        /// </summary>
        /// <param name="XMcf">The XMultiComponentFactory to use.</param>
        protected virtual void CreateDialog(XMultiComponentFactory XMcf)
        {
            try
            {
                Object oDialogModel = XMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_DIALOG_MODEL, MXContext);

                // The XMultiServiceFactory of the dialog model is needed to instantiate the controls...
                MXMsfDialogModel = (XMultiServiceFactory)oDialogModel;

                // The named container is used to insert the created controls into...
                MXDlgModelNameContainer = (XNameContainer)oDialogModel;

                // create the dialog...
                Object oUnoDialog = XMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_DIALOG, MXContext);
                MXDialogControl = (XControl)oUnoDialog;

                // The scope of the control container is public...
                MXDlgContainer = (XControlContainer)oUnoDialog;

                MXTopWindow = (XTopWindow)MXDlgContainer;

                // link the dialog and its model...
                XControlModel xControlModel = (XControlModel)oDialogModel;
                MXDialogControl.setModel(xControlModel);
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        /// <summary>
        /// Initializes the Dialog with basic properties.
        /// </summary>
        /// <param name="name">The name of the object.</param>
        /// <param name="title">The title of the dialog window.</param>
        /// <param name="width">The width of the dialog window.</param>
        /// <param name="height">The height of the dialog window.</param>
        public virtual void InitalizeDialog(string name, string title, int width, int height) { InitalizeDialog(name, title, width, height, 0, 0); }
        /// <summary>
        /// Initializes the Dialog with basic properties.
        /// </summary>
        /// <param name="name">The name of the object.</param>
        /// <param name="title">The title of the dialog window.</param>
        /// <param name="width">The width of the dialog window.</param>
        /// <param name="height">The height of the dialog window.</param>
        /// <param name="posX">The x posistion of the dialog window.</param>
        /// <param name="posY">The y posistion of the dialog window.</param>
        public virtual void InitalizeDialog(string name, string title, int width, int height, int posX, int posY) { InitalizeDialog(name, title, width, height, posX, posX, true, 0, 0); }
        /// <summary>
        /// Initializes the Dialog with basic properties.
        /// </summary>
        /// <param name="name">The name of the object.</param>
        /// <param name="title">The title of the dialog window.</param>
        /// <param name="width">The width of the dialog window.</param>
        /// <param name="height">The height of the dialog window.</param>
        /// <param name="posX">The x posistion of the dialog window.</param>
        /// <param name="posY">The y posistion of the dialog window.</param>
        /// <param name="moveable">if set to <c>true</c> the dialog window can be moved.</param>
        /// <param name="tabIndex">Tab index of the dialog window.</param>
        /// <param name="step">The step index of the dialog window.</param>
        public virtual void InitalizeDialog(string name, string title, int width, int height, int posX, int posY, bool moveable, int tabIndex, int step)
        {

            String[] PropertyNames = new String[] { "Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title", "Width" };
            Object[] PropertyValues = new Object[] { height, moveable, name, posX, posY, step, (short)tabIndex, title, width };

            try
            {
                XMultiPropertySet xMultiPropertySet = (XMultiPropertySet)MXDlgModelNameContainer;
                xMultiPropertySet.setPropertyValues(PropertyNames, Any.Get(PropertyValues));
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        #endregion

        #region Elements
        //TODO: document this

        #region Text / Label

        /// <summary>
        /// Inserts a fixed text label.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, MultiLine, Printable, TextColor, TextLineColor, VerticalAlign
        /// </summary>
        /// <param name="name">The base name of the element - will be set to a unique one by this function.</param>
        /// <param name="text">The text that should be insert.</param>
        /// <param name="_nPosX">The x position of the element.</param>
        /// <param name="_nPosY">The y posistion of the elemnt.</param>
        /// <param name="_nWidth">The width of the element.</param>
        /// <returns>XFixedText object</returns>
        public virtual XFixedText InsertFixedLabel(String text, int _nPosX, int _nPosY, int _nWidth, String sName = "") { return InsertFixedLabel(text, _nPosX, _nPosY, _nWidth, 8, 0, null, sName); }
        /// <summary>
        /// Inserts a fixed text label.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, MultiLine, Printable, TextColor, TextLineColor, VerticalAlign
        /// </summary>
        /// <param name="name">The base name of the element - will be set to a unique one by this function.</param>
        /// <param name="text">The text that should be insert.</param>
        /// <param name="_nPosX">The x position of the element.</param>
        /// <param name="_nPosY">The y posistion of the elemnt.</param>
        /// <param name="_nWidth">The width of the element.</param>
        /// <param name="_xMouseListener">A mouse listener.</param>
        /// <returns>XFixedText object</returns>
        public virtual XFixedText InsertFixedLabel(String text, int _nPosX, int _nPosY, int _nWidth, XMouseListener _xMouseListener, String sName = "") { return InsertFixedLabel(text, _nPosX, _nPosY, _nWidth, 8, 0, _xMouseListener, sName); }
        /// <summary>
        /// Inserts a fixed text label.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, MultiLine, Printable, TextColor, TextLineColor, VerticalAlign
        /// </summary>
        /// <param name="name">The base name of the element - will be set to a unique one by this function.</param>
        /// <param name="text">The text that should be insert.</param>
        /// <param name="_nPosX">The x position of the element.</param>
        /// <param name="_nPosY">The y posistion of the elemnt.</param>
        /// <param name="_nWidth">The width of the element.</param>
        /// <param name="_nHeight">The Height of the element.</param>
        /// <param name="_nStep">The step index.</param>
        /// <param name="_xMouseListener">A mouse listener.</param>
        /// <returns>XFixedText object</returns>
        public virtual XFixedText InsertFixedLabel(String text, int _nPosX, int _nPosY, int _nWidth, int _nHeight, int _nStep, XMouseListener _xMouseListener, String sName ="" )
        {
            XFixedText xFixedText = null;
            try
            {
                // create a unique name by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "TEXT_FIXED");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oFTModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_TEXT_FIXED_MODEL, MXContext);
                XMultiPropertySet xFTModelMPSet = (XMultiPropertySet)oFTModel;
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!

                String[] valueNames = new String[] { "Height", "Name", "PositionX", "PositionY", "Step", "Width", "Label" };
                var valueVals = Any.Get(new Object[] { _nHeight, sName, _nPosX, _nPosY, _nStep, _nWidth, text });

                xFTModelMPSet.setPropertyValues(
                        valueNames,valueVals
                       );
                // add the model to the NameContainer of the dialog model
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oFTModel));

                var element = MXDlgModelNameContainer.getByName(sName).Value;
                if (element != null)
                {

                    OoUtils.SetProperty(element, "PositionX", _nPosX);
                    OoUtils.SetProperty(element, "PositionY", _nPosY);
                    OoUtils.SetProperty(element, "Width", _nWidth);
                    //if (element is XControl)
                    //{
                    //    ((XControl)element).getModel();
                    //}
                    //if (element is XMultiPropertySet)
                    //{
                    //    ((XMultiPropertySet)element).setPropertyValues(valueNames, valueVals);
                    //}
                }


                // reference the control by the Name
                XControl xFTControl = GetControlByName(sName);
                xFixedText = (XFixedText)xFTControl;
                if (_xMouseListener != null)
                {
                    XWindow xWindow = (XWindow)xFTControl;
                    xWindow.addMouseListener(_xMouseListener);
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xFixedText;
        }


        public virtual XFixedText CreateFixedLabel(String text, int _nPosX, int _nPosY, int _nWidth, int _nHeight, int _nStep, XMouseListener _xMouseListener, String sName = "")
        {
            if (MXMcf == null) return null;
            XFixedText xFixedText = null;
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "TEXT_FIXED");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);

               
                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oFTModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_TEXT_FIXED_MODEL, MXContext);
                Object xFTControl = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_TEXT_FIXED, MXContext);
                
                XMultiPropertySet xFTModelMPSet = (XMultiPropertySet)oFTModel;
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!

                xFTModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "PositionX", "PositionY", "Step", "Width", "Label" },
                        Any.Get(new Object[] { _nHeight, sName, _nPosX, _nPosY, _nStep, _nWidth, text }));

                if (oFTModel != null && xFTControl != null && xFTControl is XControl)
                {
                    ((XControl)xFTControl).setModel(oFTModel as XControlModel);
                }

                xFixedText = xFTControl as XFixedText;
                if (_xMouseListener != null && xFTControl is XWindow)
                {
                    XWindow xWindow = (XWindow)xFTControl;
                    xWindow.addMouseListener(_xMouseListener);
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xFixedText;
        }


        /// <summary>
        /// Inserts a fixed one line text.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, MultiLine, Printable, TextColor, TextLineColor, VerticalAlign
        /// </summary>
        /// <param name="name">The base name of the element - will be set to a unique one by this function.</param>
        /// <param name="text">The text that should be insert.</param>
        /// <param name="_nPosX">The x position of the element.</param>
        /// <param name="_nPosY">The y posistion of the elemnt.</param>
        /// <param name="_nWidth">The width of the element.</param>
        /// <returns>XFixedText object</returns>
        public virtual XFixedText InsertFixedText(String text, int _nPosX, int _nPosY, int _nWidth, String sName = "") { return InsertFixedLabel(text, _nPosX, _nPosY, _nWidth, sName); }
        /// <summary>
        /// Inserts a fixed text.
        /// Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, MultiLine, Printable, TextColor, TextLineColor, VerticalAlign
        /// </summary>
        /// <param name="name">The base name of the element - will be set to a unique one by this function.</param>
        /// <param name="text">The text that should be insert.</param>
        /// <param name="_nPosX">The x position of the element.</param>
        /// <param name="_nPosY">The y posistion of the elemnt.</param>
        /// <param name="_nWidth">The width of the element.</param>
        /// <param name="_nHeight">The Height of the element.</param>
        /// <returns>XFixedText object</returns>
        public virtual XFixedText InsertFixedText(String text, int _nPosX, int _nPosY, int _nWidth, int _nHeight, String sName = "") { return InsertFixedLabel(text, _nPosX, _nPosY, _nWidth, _nHeight, 0, null, sName); }

        #endregion

        #region Currency Field
        //Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, BackgroundColor, Border, BorderColor#optional, CurrencySymbol, DecimalAccuracy, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, HideInactiveSelection#optional, PrependCurrencySymbol, Printable, ReadOnly, Repeat#optional, RepeatDelay#optional, ShowThousandsSeparator, Spin, StrictFormat, Tabstop, TextColor, TextLineColor, Value, ValueMax, ValueMin, ValueStep
        public XTextComponent InsertCurrencyField(int _nPositionX, int _nPositionY, int _nWidth, XTextListener _xTextListener, String sName = "") { return InsertCurrencyField(_nPositionX, _nPositionY, _nWidth, 12, String.Empty, 0.0, _xTextListener, sName); }
        public XTextComponent InsertCurrencyField(int _nPositionX, int _nPositionY, int _nWidth, double defaultValue, XTextListener _xTextListener, String sName = "") { return InsertCurrencyField( _nPositionX, _nPositionY, _nWidth, 12, String.Empty, defaultValue, _xTextListener, sName); }
        public XTextComponent InsertCurrencyField(int _nPositionX, int _nPositionY, int _nWidth, int _nHeight, string curencySymbol, double defaultValue, XTextListener _xTextListener, String sName = "")
        {
            XTextComponent xTextComponent = null;
            try
            {
                // create a unique name by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "CURRENCY_FIELD");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);


                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oCFModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_CURRENCY_FIELD_MODEL, MXContext);
                XMultiPropertySet xCFModelMPSet = (XMultiPropertySet)oCFModel;

                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xCFModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "PositionX", "PositionY", "Width" },
                        Any.Get(new Object[] { _nHeight, sName, _nPositionX, _nPositionY, _nWidth }));

                // The controlmodel is not really available until inserted to the Dialog container
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oCFModel));
                XPropertySet xCFModelPSet = (XPropertySet)oCFModel;

                // The following properties may also be set with XMultiPropertySet but we
                // use the XPropertySet interface merely for reasons of demonstration
                if (!curencySymbol.Equals(String.Empty))
                {
                    xCFModelPSet.setPropertyValue("PrependCurrencySymbol", Any.Get(true));
                    xCFModelPSet.setPropertyValue("CurrencySymbol", Any.Get(curencySymbol));
                }
                else
                {
                    xCFModelPSet.setPropertyValue("PrependCurrencySymbol", Any.Get(false));
                }
                xCFModelPSet.setPropertyValue("Value", Any.Get(defaultValue));

                if (_xTextListener != null)
                {
                    // add a textlistener that is notified on each change of the controlvalue...
                    Object oCFControl = GetControlByName(sName);
                    xTextComponent = (XTextComponent)oCFControl;
                    xTextComponent.addTextListener(_xTextListener);
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xTextComponent;
        }
        #endregion

        #region Progress Bar
        // Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, BackgroundColor, Border, BorderColor, Enabled, FillColor, HelpText, HelpURL, Printable, ProgressValue, ProgressValueMax, ProgressValueMin
        public XPropertySet InsertProgressBar(int _nPosX, int _nPosY, int _nWidth, int _nProgressMax, String sName = "") { return InsertProgressBar(_nPosX, _nPosY, _nWidth, _nProgressMax, 0, sName); }
        public XPropertySet InsertProgressBar(int _nPosX, int _nPosY, int _nWidth, int _nProgressMax, int _nProgress, String sName = "") { return InsertProgressBar(_nPosX, _nPosY, _nWidth, 8, _nProgressMax, _nProgress, sName); }
        public XPropertySet InsertProgressBar(int _nPosX, int _nPosY, int _nWidth, int _nHeight, int _nProgressMax, int _nProgress, String sName = "")
        {
            XPropertySet xPBModelPSet = null;
            try
            {
                // create a unique name by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "PROGRESS_BAR");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oPBModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_PROGRESS_BAR_MODEL, MXContext);

                XMultiPropertySet xPBModelMPSet = (XMultiPropertySet)oPBModel;
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xPBModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "PositionX", "PositionY", "Width" },
                        Any.Get(new Object[] { _nHeight, sName, _nPosX, _nPosY, _nWidth }));

                // The controlmodel is not really available until inserted to the Dialog container
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oPBModel));
                xPBModelPSet = (XPropertySet)oPBModel;

                // The following properties may also be set with XMultiPropertySet but we
                // use the XPropertySet interface merely for reasons of demonstration
                xPBModelPSet.setPropertyValue("ProgressValueMin", Any.Get(0));
                xPBModelPSet.setPropertyValue("ProgressValueMax", Any.Get(_nProgressMax));
                xPBModelPSet.setPropertyValue("ProgressValue", Any.Get(_nProgress));
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xPBModelPSet;
        }

        #endregion

        #region Horizontal Line

        // Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HelpText, HelpURL, Label, Orientation, Printable, TextColor, TextLineColor
        public Object InsertHorizontalFixedLine(int _nPosX, int _nPosY, int _nWidth, String sName = "") { return InsertHorizontalFixedLine(String.Empty, _nPosX, _nPosY, _nWidth, 8, sName); }
        public Object InsertHorizontalFixedLine(int _nPosX, int _nPosY, int _nWidth, int _nHeight, String sName = "") { return InsertHorizontalFixedLine(String.Empty, _nPosX, _nPosY, _nWidth, _nHeight, 0, sName); }
        public Object InsertHorizontalFixedLine(String _sLabel, int _nPosX, int _nPosY, int _nWidth, String sName = "") { return InsertHorizontalFixedLine(_sLabel, _nPosX, _nPosY, _nWidth, 8, sName); }
        public Object InsertHorizontalFixedLine(String _sLabel, int _nPosX, int _nPosY, int _nWidth, int _nHeight, String sName = "") { return InsertHorizontalFixedLine(_sLabel, _nPosX, _nPosY, _nWidth, _nHeight, 0, sName); }
        public Object InsertHorizontalFixedLine(String _sLabel, int _nPosX, int _nPosY, int _nWidth, int _nHeight, int orientation, String sName = "")
        {
            try
            {
                // create a unique name by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "FIXED_LINE");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oFLModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_FIXED_LINE_MODEL, MXContext);
                XMultiPropertySet xFLModelMPSet = (XMultiPropertySet)oFLModel;

                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xFLModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "Orientation", "PositionX", "PositionY", "Width" },
                        Any.Get(new Object[] { _nHeight, sName, orientation, _nPosX, _nPosY, _nWidth }));

                // The controlmodel is not really available until inserted to the Dialog container
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oFLModel));

                // The following property may also be set with XMultiPropertySet but we
                // use the XPropertySet interface merely for reasons of demonstration
                XPropertySet xFLPSet = (XPropertySet)oFLModel;
                xFLPSet.setPropertyValue("Label", Any.Get(_sLabel));

                return oFLModel;
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return null;
        }

        public XControl CreateHorizontalFixedLine(String _sLabel, int _nPosX, int _nPosY, int _nWidth, int _nHeight, int orientation, String sName = "")
        {
            if (MXMcf == null) return null;
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "FIXED_LINE");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oFLModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_FIXED_LINE_MODEL, MXContext);
                Object oFLControl = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_FIXED_LINE, MXContext);

                if (oFLModel != null && oFLControl != null && oFLControl is XControl)
                {
                    ((XControl)oFLControl).setModel(oFLModel as XControlModel);
                }


                XMultiPropertySet xFLModelMPSet = (XMultiPropertySet)oFLModel;

                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xFLModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "Orientation", "PositionX", "PositionY", "Width" },
                        Any.Get(new Object[] { _nHeight, sName, orientation, _nPosX, _nPosY, _nWidth }));

                XPropertySet xFLPSet = (XPropertySet)oFLModel;
                xFLPSet.setPropertyValue("Label", Any.Get(_sLabel));

                return oFLControl as XControl;
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return null;
        }



        #endregion

        #region Edit Field 
        // Properties: Height, Name, PositionX, PositionY, Step, TabIndex, Tag, Width, Align, BackgroundColor, Border, BorderColor, EchoChar, Enabled, FontDescriptor, FontEmphasisMark, FontRelief, HardLineBreaks, HelpText, HelpURL, HideInactiveSelection, HScroll, LineEndFormat, MaxTextLen, MultiLine, Printable, ReadOnly, Tabstop, Text, TextColor, TextLineColor, VScroll
        public XTextComponent InsertEditField(String defaultText, int _nPosX, int _nPosY, int _nWidth, XTextListener _xTextListener, XFocusListener _xFocusListener, XKeyListener _xKeyListener, String sName = "") { return InsertEditField(defaultText, _nPosX, _nPosY, _nWidth, 12, String.Empty, _xTextListener, _xFocusListener, _xKeyListener, sName); }
        public XTextComponent InsertEditField(String defaultText, int _nPosX, int _nPosY, int _nWidth, int _nHeight, XTextListener _xTextListener, XFocusListener _xFocusListener, XKeyListener _xKeyListener, String sName = "") { return InsertEditField(defaultText, _nPosX, _nPosY, _nWidth, 12, String.Empty, _xTextListener, _xFocusListener, _xKeyListener, sName); }
        public XTextComponent InsertEditField(String defaultText, int _nPosX, int _nPosY, int _nWidth, String echoChar, XTextListener _xTextListener, XFocusListener _xFocusListener, XKeyListener _xKeyListener, String sName = "") { return InsertEditField(defaultText, _nPosX, _nPosY, _nWidth, 12, echoChar, _xTextListener, _xFocusListener, _xKeyListener, sName); }
        public XTextComponent InsertEditField(String defaultText, int _nPosX, int _nPosY, int _nWidth, int _nHeight, String echoChar, XTextListener _xTextListener, XFocusListener _xFocusListener, XKeyListener _xKeyListener, String sName = "")
        {
            XTextComponent xTextComponent = null;
            try
            {
                // create a unique name by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = createUniqueName(MXDlgModelNameContainer, "EDIT");
                else sName = createUniqueName(MXDlgModelNameContainer, sName);
                

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oTFModel = MXMcf.createInstanceWithContext(OO.Services.AWT_CONTROL_EDIT_MODEL, MXContext);
                XMultiPropertySet xTFModelMPSet = (XMultiPropertySet)oTFModel;

                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xTFModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "PositionX", "PositionY", "Text", "Width" },
                       Any.Get(new Object[] { _nHeight, sName, _nPosX, _nPosY, defaultText, _nWidth }));

                // The controlmodel is not really available until inserted to the Dialog container
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oTFModel));

                if (!echoChar.Equals(String.Empty))
                {
                    XPropertySet xTFModelPSet = (XPropertySet)oTFModel;

                    // The following property may also be set with XMultiPropertySet but we
                    // use the XPropertySet interface merely for reasons of demonstration
                    xTFModelPSet.setPropertyValue("EchoChar", Any.Get((short)echoChar.ToCharArray(0, 1)[0]));
                }

                if (_xFocusListener != null || _xTextListener != null || _xKeyListener != null)
                {
                    XControl xTFControl = GetControlByName(sName);

                    // add a textlistener that is notified on each change of the controlvalue...
                    xTextComponent = (XTextComponent)xTFControl;
                    XWindow xTFWindow = (XWindow)xTFControl;
                    if (_xFocusListener != null) xTFWindow.addFocusListener(_xFocusListener);
                    if (_xTextListener != null) xTextComponent.addTextListener(_xTextListener);
                    if (_xKeyListener != null) xTFWindow.addKeyListener(_xKeyListener);
                }

            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xTextComponent;
        }

        #endregion

        #endregion

        #region Utils

        /// <summary>
        /// Set the properties at the model - keep in mind to pass the property names in alphabetical order!
        /// </summary>
        /// <param name="xmps">The MultiPropertySet.</param>
        /// <param name="keys">The property names.</param>
        /// <param name="values">The property values.</param>
        public static void setPropertyValues(XMultiPropertySet xmps, String[] keys, Object[] values)
        {
            if (xmps != null && keys.Length > 0 && values.Length > 0 && keys.Length == values.Length)
            {
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xmps.setPropertyValues(keys, Any.Get(values));
            }
        }

        /// <summary>
        /// Makes a String unique by appending a numerical suffix.
        /// </summary>
        /// <param name="_xElementContainer">the com.sun.star.container.XNameAccess container
        /// that the new Element is going to be inserted to</param>
        /// <param name="_sElementName"the StemName of the Element.</param>
        /// <returns>A String unique by appending a numerical suffix</returns>
        public static String createUniqueName(XNameAccess _xElementContainer, String _sElementName)
        {
            bool bElementexists = true;
            int i = 1;
            String BaseName = _sElementName;
            while (bElementexists)
            {
                bElementexists = _xElementContainer.hasByName(_sElementName);
                if (bElementexists)
                {
                    i += 1;
                    _sElementName = BaseName + i;
                }
            }
            return _sElementName;
        }
        
        /// <summary>
        /// Sets the special property 'step' if possible.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        /// <param name="step">The 'step' id.</param>
        /// <returns></returns>
        public bool SetStepProperty(object propertySet, int step)
        {
            return SetProperty(propertySet, "Step", step);
        }

        /// <summary>
        /// Sets a property if possible.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        /// <param name="name">The name of the property.</param>
        /// <param name="value">The value.</param>
        /// <returns><c>true</c> if the property could be set otherwise <c>false</c></returns>
        public bool SetProperty(object propertySet, string name, object value)
        {
            if (propertySet != null)
            {
                XPropertySet ps = null;
                if (propertySet is XPropertySet)
                {
                    ps = propertySet as XPropertySet;
                }
                else if (propertySet is XControl)
                { // for a XControll you have to get the model to change or get the properties
                    XControlModel model = ((XControl)propertySet).getModel();
                    if (model != null && model is XPropertySet)
                    {
                        ps = model as XPropertySet;
                    }
                }

                if (ps != null)
                {
                    var psInfo = ps.getPropertySetInfo();
                    bool exist = psInfo.hasPropertyByName(name);

                    //util.Debug.GetAllProperties(ps);

                    if (exist)
                    {
                        ps.setPropertyValue(name, Any.Get(value));
                        return true;
                    }
                    else { }
                }
            }
            return false;
        }

        /// <summary>
        /// Gets the property of a PropertySet or a XControl.
        /// </summary>
        /// <param name="propertySet">The property set or the control.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>the value of the Any Object returned from the PropertySet</returns>
        public static Object GetProperty(Object propertySet, String propName)
        {
            if (propertySet != null && !String.IsNullOrWhiteSpace(propName))
            {

                XPropertySet ps = null;
                if (propertySet is XPropertySet)
                {
                    ps = propertySet as XPropertySet;
                }
                else if (propertySet is XControl)
                { // for a XControll you have to get the model to change or get the properties
                    XControlModel model = ((XControl)propertySet).getModel();
                    if (model != null && model is XPropertySet)
                    {
                        ps = model as XPropertySet;
                    }
                }

                if (ps != null)
                {

                    return OoUtils.GetProperty(ps, propName);
                }
            }

            return null;
        }


        /// <summary>
        /// Try to get a XControl of the dialog by name.
        /// </summary>
        /// <param name="sName">Name of the XControll to search for.</param>
        /// <returns>The XControll with the searched name od <c>null</c></returns>
        public XControl GetControlByName(String sName) {
            if (MXDlgContainer != null) return MXDlgContainer.getControl(sName);
            return null;
        }

        /// <summary>
        /// Gets the name of a control.
        /// </summary>
        /// <param name="xCtr">The XControl.</param>
        /// <returns>the name or the empty String</returns>
        public static String GetNameOfControl(XControl xCtr)
        {

            if (xCtr != null)
            {
                Object nameObj = GetProperty(xCtr, "Name");
                if (nameObj != null && nameObj is String)
                    return nameObj.ToString();
            }

            return String.Empty;
        }

        public static XPropertySet GetPropertysetOfControl(Object xCtr) { return GetPropertysetOfControl(xCtr as XControl); }
        public static XPropertySet GetPropertysetOfControl(XControl xCtr){

            if (xCtr != null)
            {
                XControlModel model =xCtr.getModel();
                if (model != null && model is XPropertySet)
                {
                   return model as XPropertySet;
                }
            }
            
            return null;
        }


        #endregion
    }
}
