using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.awt;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.xml.dom.events;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.uno;

namespace tud.mci.tangram.models.dialogs
{

    /*
     *  Name: sName / SCROLL_CONTAINER
     *  XControl: OuterScrlContr
     *  XControlModel: outerScrlContrModel
     *  XMultiPropertySet: xoScrlContMPSet
     *  ╔═══════════════════════════════════════════════════════════╗┌─╖ 
     *  ║   ┌────────────────┐                                      ║│▲║
     *  ║   │  ██████████    │  Name: sName + "_S"                  ║│░║    XScrollBar: VerticalScrlBar
     *  ║   │                │  XControl: InnerScrlContr            ║│░║
     *  ║   │  ██████        │  XControlModel: innerScrlContrModel  ║│░║
     *  ║   │                │                                      ║│░║
     *  ║   │  ██████████    │                                      ║│░║
     *  ║   │                │                                      ║│▓║
     *  ║   │  ███████████   │                                      ║│▓║
     *  ║   │                │                                      ║│▓║
     *  ║   └────────────────┘                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│░║
     *  ║   ¦                ¦                                      ║│▼║
     *  ╚═══════════════════════════════════════════════════════════╝└─╜
     *  ┌───────────────────────────────────────────────────────────┐
     *  │◄░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░▓▓▓▓░░░░░░░░░░░░░░░░►│       XScrollBar: HorizontalScrlBar 
     *  ╘═══════════════════════════════════════════════════════════╛
     *      ¦               ¦
     *      ¦               ¦
     *      ¦               ¦
     *      ¦               ¦
     *      └---------------┘
     */

    public class ScrollableContainer : XAdjustmentListener, XMouseListener, IDisposable
    {
        #region Members

        #region Position and Size

        #region Outer Container

        volatile int _posX = 0;
        public int PosX
        {
            get { return _posX; }
            set
            {
                _posX = value;
                updateOuterSize();
            }
        }

        volatile int _posY = 0;
        public int PosY
        {
            get { return _posY; }
            set
            {
                _posY = value;
                updateOuterSize();
            }
        }

        volatile int _width = 0;
        public int Width
        {
            get { return _width; }
            set
            {
                _width = value;
                updateOuterSize();
            }
        }

        volatile int _height = 0;
        public int Height
        {
            get { return _height; }
            set
            {
                _height = value;
                updateOuterSize();
            }
        }

        #endregion


        #region Inner Container

        volatile int _innerPosX = 0;
        public int InnerPosX
        {
            get { return _innerPosX; }
            set
            {
                _innerPosX = value;
                updateInnerSize();
            }
        }

        volatile int _innerPosY = 0;
        public int InnerPosY
        {
            get { return _innerPosY; }
            set
            {
                _innerPosY = value;
                updateInnerSize();
            }
        }

        volatile int _innerWidth = 0;
        public int InnerWidth
        {
            get { return _innerWidth; }
            set
            {
                _innerWidth = value;
                updateInnerSize();
            }
        }

        volatile int _innerHeight = 0;
        public int InnerHeight
        {
            get { return _innerHeight; }
            set
            {
                _innerHeight = value;
                updateInnerSize();
            }
        }

        #endregion


        #endregion

        #region Basic Elements from the context

        XNameContainer MXDlgModelNameContainer;

        XComponentContext _mXContext = null;
        XComponentContext MXContext
        {
            get
            {
                if (_mXContext == null) _mXContext = OO.GetContext();
                return _mXContext;
            }
            set { _mXContext = value; }
        }

        XControlContainer parentCnt;

        XMultiComponentFactory _parentMCF = null;
        XMultiComponentFactory parentMCF
        {
            get
            {
                if (_parentMCF == null) _parentMCF = OO.GetMultiComponentFactory(MXContext);
                return _parentMCF;
            }
            set { _parentMCF = value; }
        }


        #endregion

        #region Basic Elements

        public XScrollBar HorizontalScrlBar { get; private set; }
        public XScrollBar VerticalScrlBar { get; private set; }

        public XControl OuterScrlContr { get; private set; }
        XControlModel outerScrlContrModel;

        public XControl InnerScrlContr { get; private set; }
        XControlModel innerScrlContrModel;

        #endregion


        #endregion


        // use MultiService factory

        public ScrollableContainer(XControlContainer parent, XNameContainer nameContainer, XComponentContext context = null, XMultiComponentFactory parentMCF = null)
        {
            MXDlgModelNameContainer = nameContainer;
            if (context != null) MXContext = context;
            parentCnt = parent;
            this.parentMCF = parentMCF;
        }

        #region Scrollable Container

        private const string AWT_UNO_CONTROL_CONTAINER = "com.sun.star.awt.UnoControlContainer";
        private const string AWT_UNO_CONTROL_CONTAINER_Model = "com.sun.star.awt.UnoControlContainerModel";

        private const string AWT_UNO_CONTROL_SCROLLBAR_MODEL = "com.sun.star.awt.UnoControlScrollBarModel";

        public void CreateScrollableContainer(int _nPosX, int _nPosY, int _nWidth, int _nHeight, String sName)
        {
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, "SCROLL_CONTAINER");
                else sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, sName);

                //XMultiServiceFactory _xMSF = OO.GetMultiServiceFactory(MXContext);
                XMultiComponentFactory _xMcf = OO.GetMultiComponentFactory(MXContext, true);

                #region Outer Container

                // create a UnoControlContainerModel. A thing which differs from other dialog-controls in many aspects
                // Position and size of the model have no effect, so we apply setPosSize() on it's view later.
                // Unlike a dialog-model the container-model can not have any control-models.
                // As an instance of a dialog-model it can have the property "Step"(among other things),
                // provided by service awt.UnoControlDialogElement, without actually supporting this service!

                // create a control model at the multi service factory of the dialog model...
                outerScrlContrModel = _xMcf.createInstanceWithContext(AWT_UNO_CONTROL_CONTAINER_Model, MXContext) as XControlModel;
                OuterScrlContr = _xMcf.createInstanceWithContext(AWT_UNO_CONTROL_CONTAINER, MXContext) as XControl;


                if (OuterScrlContr != null && outerScrlContrModel != null)
                {
                    OuterScrlContr.setModel(outerScrlContrModel);
                }

                XMultiPropertySet xoScrlContMPSet = outerScrlContrModel as XMultiPropertySet;

                if (xoScrlContMPSet != null)
                {
                    // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                    xoScrlContMPSet.setPropertyValues(
                            new String[] { "Height", "Name", "PositionX", "PositionY", "State", "Width" },
                            Any.Get(new Object[] { 0, sName, 0, 0, ((short)0), 0, 0 }));
                }

                //add to the dialog
                XControlContainer xCntrCont = parentCnt as XControlContainer;
                if (xCntrCont != null)
                {
                    xCntrCont.addControl(sName, OuterScrlContr as XControl);
                }

                #endregion

                #region Scroll bars

                //insert scroll bars
                HorizontalScrlBar = insertHorizontalScrollBar(this, _nPosX, _nPosY + _nHeight, _nWidth);
                VerticalScrlBar = insertVerticalScrollBar(this, _nPosX + _nWidth, _nPosY, _nHeight);

                #endregion

                #region Set Size of outer Container

                //make the outer container pos and size via the pos and size of the scroll bars
                if (HorizontalScrlBar is XWindow)
                {
                    Rectangle hSBPos = ((XWindow)HorizontalScrlBar).getPosSize();
                    _width = hSBPos.Width;
                    _posX = hSBPos.X;
                }
                if (VerticalScrlBar is XWindow)
                {
                    Rectangle vSBPos = ((XWindow)VerticalScrlBar).getPosSize();
                    _height = vSBPos.Height;
                    _posY = vSBPos.Y;
                }

                // Set the size of the surrounding container
                if (OuterScrlContr is XWindow)
                {
                    ((XWindow)OuterScrlContr).setPosSize(PosX, PosY, Width, Height, PosSize.POSSIZE);
                    ((XWindow)OuterScrlContr).addMouseListener(this);

                }
                #endregion

                #region inner Container


                // create a control model at the multi service factory of the dialog model...
                innerScrlContrModel = _xMcf.createInstanceWithContext(AWT_UNO_CONTROL_CONTAINER_Model, MXContext) as XControlModel;
                InnerScrlContr = _xMcf.createInstanceWithContext(AWT_UNO_CONTROL_CONTAINER, MXContext) as XControl;


                if (InnerScrlContr != null && innerScrlContrModel != null)
                {
                    InnerScrlContr.setModel(innerScrlContrModel);
                }

                XMultiPropertySet xinnerScrlContMPSet = innerScrlContrModel as XMultiPropertySet;

                if (xinnerScrlContMPSet != null)
                {
                    xinnerScrlContMPSet.setPropertyValues(
                            new String[] { "Name", "State" },
                            Any.Get(new Object[] { sName + "_S", ((short)0) }));
                }

                //// FIXME: only for fixing

                //util.Debug.GetAllServicesOfObject(parentCnt);

                //Object oDialogModel = OO.GetMultiComponentFactory().createInstanceWithContext(OO.Services.AWT_CONTROL_DIALOG_MODEL, MXContext);

                //// The named container is used to insert the created controls into...
                //MXDlgModelNameContainer = (XNameContainer)oDialogModel;

                //System.Diagnostics.Debug.WriteLine("_________");

                //util.Debug.GetAllServicesOfObject(MXDlgModelNameContainer);

                //// END

                //add inner container to the outer scroll container
                XControlContainer outerCntrCont = OuterScrlContr as XControlContainer;
                if (outerCntrCont != null)
                {
                    outerCntrCont.addControl(sName + "_S", InnerScrlContr as XControl);

                    InnerWidth = Width;
                    InnerHeight = Height;

                    // Set the size of the surrounding container
                    if (InnerScrlContr is XWindow)
                    {
                        ((XWindow)InnerScrlContr).setPosSize(0, 0, InnerWidth, InnerHeight, PosSize.POSSIZE);
                        ((XWindow)InnerScrlContr).addMouseListener(this);
                    }
                }

                #endregion

            }
            catch (System.Exception) { }
        }

        #region Add Elements

        public XControl AddElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin)
        {
            return AddElementToTheEndAndAdoptTheSize(cntrl, sName, posX, topMargin, Width - posX);
        }
        public XControl AddElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin, int width)
        {
            return AddElementToTheEndAndAdoptTheSize(cntrl, sName, posX, topMargin, width, 10);
        }
        public XControl AddElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin, int width, int height)
        {
            //get all elements in the container

            if (cntrl != null)
            {
                if (InnerScrlContr != null && InnerScrlContr is XControlContainer)
                {
                    XControl[] ctrls = ((XControlContainer)InnerScrlContr).getControls();
                    Rectangle p = new Rectangle();
                    p.X = posX;

                    if (ctrls != null && ctrls.Length > 0)
                    {
                        foreach (var item in ctrls)
                        {
                            if (item != null && item is XWindow)
                            {
                                Rectangle pos = ((XWindow)item).getPosSize();
                                p.Width = Math.Max(p.Width, pos.X + pos.Width);
                                p.Height = Math.Max(p.Height, pos.Y + pos.Height);
                            }
                        }
                    }

                    return AddElementAndAdoptTheSize(cntrl, sName, posX, p.Height + topMargin, width, height);

                }

            }
            return cntrl;
        }

        public XControl AddElementTryToKeeptheHeight(XControl cntrl, String sName, int posX, int posY, int height)
        {
            return AddElementAndAdoptTheSize(cntrl, sName, posX, posY, Width - posX, height);
        }

        public XControl AddElementTryToKeeptheWidth(XControl cntrl, String sName, int posX, int posY, int width)
        {
            return AddElementAndAdoptTheSize(cntrl, sName, posX, posY, width, 15);
        }

        public XControl AddElementAndAdoptTheSize(XControl cntrl, String sName, int posX, int posY)
        {
            return AddElementTryToKeeptheWidth(cntrl, sName, posX, posY, Width - posX);
        }

        public XControl AddElementAndAdoptTheSize(XControl cntrl, String sName, int posX, int posY, int width, int height)
        {
            Size s = new Size(width, height);
            if (cntrl is XLayoutConstrains)
            {
                s = ((XLayoutConstrains)cntrl).calcAdjustedSize(s);

                s.Width = Math.Max(width, s.Width);
                s.Height = Math.Max(height, s.Height);
            }
            return AddElement(cntrl, sName, posX, posY, s.Width, s.Height);
        }


        /// <summary>
        /// Adds a control element to the scrollable container.
        /// ATTENTION: The pos and size you have given to the control are getting lost!
        /// The position and the size have to be set after adding the control inside the container.
        /// The values of position and size are pixel based values. The position is relative to 
        /// the top left corner of the container.
        /// </summary>
        /// <param id="cntrl">The XControll you want to add.</param>
        /// <param id="sName">The supposed id of the controll - will be changed if it is not unique!.</param>
        /// <param id="posX">The x position - pixel based value relative to the top left corner of the container.</param>
        /// <param id="posY">The y position - pixel based value relative to the top left corner of the container..</param>
        /// <param id="width">The width in pixel.</param>
        /// <param id="height">The height in pixel.</param>
        public XControl AddElement(XControl cntrl, String sName, int posX, int posY, int width, int height)
        {
            // create a unique id by means of an own implementation...
            if (String.IsNullOrWhiteSpace(sName)) sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, "SCROLLABLE_CONTROL_ENTRY");
            else sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, sName);

            //System.Diagnostics.Debug.WriteLine("added control: ____________________________");
            //Debug.GetAllInterfacesOfObject(cntrl);
            //Debug.GetAllServicesOfObject(cntrl);
            //Debug.GetAllProperties(cntrl);

            if (cntrl != null && InnerScrlContr != null)
            {
                if (InnerScrlContr is XControlContainer)
                {
                    ((XControlContainer)InnerScrlContr).addControl(sName, cntrl);
                    OoUtils.SetProperty(cntrl, "Name", sName);
                }

                if (cntrl is XWindow)
                {
                    ((XWindow)cntrl).setPosSize(posX, posY, width, height, PosSize.POSSIZE);

                    Rectangle contrPos = ((XWindow)cntrl).getPosSize();

                    if (InnerWidth < (contrPos.X + contrPos.Width)) InnerWidth = contrPos.X + contrPos.Width;
                    if (InnerHeight < (contrPos.Y + contrPos.Height)) InnerHeight = contrPos.Y + contrPos.Height;

                }
            }

            return cntrl;
        }

        #endregion

        #region ScrollBars

        private const int defaultScrollBarwidth = 10;

        /// <summary>
        /// Inserts a vertical scroll bar.
        /// </summary>
        /// <param id="_nPosX">The X position.</param>
        /// <param id="_nPosY">The Y position.</param>
        /// <param id="_nHeight">Height of the Scrollbar.</param>
        /// <param id="sName">Name of the XControl - can be empty.</param>
        private XScrollBar insertVerticalScrollBar(XAdjustmentListener _xAdjustmentListener, int _nPosX, int _nPosY, int _nHeight, String sName = "")
        {
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, "VERTICAL_SCROLLBAR");
                else sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, sName);

                return insertScrollBar(_xAdjustmentListener, _nPosX, _nPosY, _nHeight, defaultScrollBarwidth, unoidl.com.sun.star.awt.ScrollBarOrientation.VERTICAL, sName);

            }
            catch { }
            return null;

        }

        /// <summary>
        /// Inserts s horizontal scroll bar.
        /// </summary>
        /// <param id="_nPosX">The X position.</param>
        /// <param id="_nPosY">The Y position.</param>
        /// <param id="_nWidth">Width of the Scrollbar.</param>
        /// <param id="sName">Name of the XControl - can be empty.</param>
        private XScrollBar insertHorizontalScrollBar(XAdjustmentListener _xAdjustmentListener, int _nPosX, int _nPosY, int _nWidth, String sName = "")
        {
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, "HORIZONTAL_SCROLLBAR");
                else sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, sName);
                return insertScrollBar(_xAdjustmentListener, _nPosX, _nPosY, defaultScrollBarwidth, _nWidth, unoidl.com.sun.star.awt.ScrollBarOrientation.HORIZONTAL, sName);
            }
            catch { }
            return null;
        }


        /// <summary>
        /// Inserts the scroll bar.
        /// </summary>
        /// <param id="_nPosX">The X position.</param>
        /// <param id="_nPosY">The Y position.</param>
        /// <param id="_nHeight">Height of the Scrollbar.</param>
        /// <param id="_nWidth">Width of the Scrollbar.</param>
        /// <param id="sName">Name of the XControl - can be empty.</param>
        private XScrollBar insertScrollBar(XAdjustmentListener _xAdjustmentListener, int _nPosX, int _nPosY, int _nHeight, int _nWidth, int orientation = unoidl.com.sun.star.awt.ScrollBarOrientation.VERTICAL, String sName = "", bool liveScroll = true)
        {
            try
            {
                // create a unique id by means of an own implementation...
                if (String.IsNullOrWhiteSpace(sName)) sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, "SCROLLBAR");
                else sName = AbstactUnoDialogBase.createUniqueName(MXDlgModelNameContainer, sName);

                // create a control model at the multiservicefactory of the dialog model...
                Object oSBModel = parentMCF.createInstanceWithContext(OO.Services.AWT_CONTROL_SCROLLBAR_MODEL, _mXContext);
                XMultiPropertySet xRBMPSet = (XMultiPropertySet)oSBModel;
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
                xRBMPSet.setPropertyValues(
                        new String[] { "BlockIncrement", "Height", "LineIncrement", "LiveScroll", "Name", "Orientation", "PositionX", "PositionY", "Width" },
                        Any.Get(new Object[] { 30, _nHeight, 10, liveScroll, sName, orientation, _nPosX, _nPosY, _nWidth }));
                // add the model to the NameContainer of the dialog model
                MXDlgModelNameContainer.insertByName(sName, Any.Get(oSBModel));

                XControl xSBControl = (parentCnt != null) ? parentCnt.getControl(sName) : null;

                if (xSBControl != null && xSBControl is XScrollBar && _xAdjustmentListener != null)
                {
                    ((XScrollBar)xSBControl).addAdjustmentListener(_xAdjustmentListener);
                }

                return xSBControl as XScrollBar;

            }
            catch { }

            return null;
        }
        #endregion


        #region Helper Functions

        public void Clear()
        {
            //TODO: clear the inner container
        }

        public bool SetVisible(bool visible)
        {
            //set container visible
            if (OuterScrlContr != null && OuterScrlContr is XWindow2)
            {
                ((XWindow2)OuterScrlContr).setVisible(visible);


                // set scrollbars visible
                if (VerticalScrlBar != null && VerticalScrlBar is XWindow2) ((XWindow2)VerticalScrlBar).setVisible(visible);
                if (HorizontalScrlBar != null && HorizontalScrlBar is XWindow2) ((XWindow2)HorizontalScrlBar).setVisible(visible);


                return ((XWindow2)OuterScrlContr).isVisible();

            }

            return false;
        }

        public bool IsVisible()
        {
            if (OuterScrlContr != null && OuterScrlContr is XWindow2)
            {
                return ((XWindow2)OuterScrlContr).isVisible();
            }
            return false;
        }


        #endregion

        #endregion

        #region util

        #region Position and size update

        /// <summary>
        /// Updates the size of the inner control container.
        /// </summary>
        void updateInnerSize()
        {
            if (InnerScrlContr != null && InnerScrlContr is XWindow)
            {
                ((XWindow)InnerScrlContr).setPosSize(InnerPosX, InnerPosY, InnerWidth, InnerHeight, PosSize.POSSIZE);
                updateScrollBars();
            }
        }

        /// <summary>
        /// Updates the size of the outer control container and the scroll bars.
        /// </summary>
        void updateOuterSize()
        {
            if (OuterScrlContr != null && OuterScrlContr is XWindow)
            {
                ((XWindow)OuterScrlContr).setPosSize(PosX, PosY, Width, Height, PosSize.POSSIZE);
                //TODO update the scroll bars
                updateScrollBars();
            }
        }

        void updateVerticalScrollBar()
        {
            try
            {
                if (VerticalScrlBar != null)
                {
                    if (VerticalScrlBar is XWindow)
                    {
                        //((XWindow)vScrlBar).setPosSize(PosX, PosY, Width, Height, PosSize.POSSIZE);
                    }
                    updateScrollBar(VerticalScrlBar, InnerHeight, Height, InnerPosY);
                }
            }
            catch (System.Exception) { }
        }

        void updateHorizontalScrollBar()
        {
            try
            {
                if (HorizontalScrlBar != null)
                {
                    updateScrollBar(HorizontalScrlBar, InnerWidth, Width, InnerPosX);
                }
            }
            catch (System.Exception) { }
        }


        bool updateScrollBar(XScrollBar xsb, int inner, int outer, int scrVal = 0)
        {
            try
            {
                int svm = Math.Max(0, inner - outer);

                //object scrVal = OoUtils.GetProperty(xsb, "ScrollValue");

                int val = Math.Min(svm, scrVal);

                OoUtils.SetProperty(xsb, "ScrollValueMax", svm);
                OoUtils.SetProperty(xsb, "ScrollValue", val);

                double ratio = (double)outer / (double)inner;
                double scroller = svm * ratio;

                // adapt the slider size
                OoUtils.SetProperty(xsb, "VisibleSize", (svm - (int)scroller));

                if (svm > 0)
                {
                    OoUtils.SetProperty(xsb, "EnableVisible", true);
                }
                else
                {
                    OoUtils.SetProperty(xsb, "EnableVisible", false);
                }

            }
            catch (System.Exception) { }

            return true;
        }

        void updateScrollBars()
        {
            updateVerticalScrollBar();
            updateHorizontalScrollBar();
        }

        void scrollVertical(int step)
        {
            InnerPosY = 0 - step;
        }

        void scrollHorizontal(int step)
        {
            InnerPosX = 0 - step;
        }


        #endregion

        #endregion


        #region XAdjustmentListener - makes it possible to receive adjustment events.

        /// <summary>
        /// is invoked when the adjustment has changed. 
        /// </summary>
        /// <param id="rEvent">The r event.</param>
        void XAdjustmentListener.adjustmentValueChanged(AdjustmentEvent rEvent)
        {
            if (rEvent != null && rEvent.Source is XControl)
            {
                XControlModel sbModel = ((XControl)rEvent.Source).getModel();
                var Orientation = OoUtils.GetProperty(rEvent.Source, "Orientation");
                int o = (Orientation is int) ? (int)Orientation : -1;

                // the value of a scroll bar is the topmost point of the 
                // scroller. That means the size of the scroller (VisibleSize)
                // is some kind of margin to the bottom of the scroll region.
                // Is is not possible to scroll down to the ScrollValueMax.
                // The maximum value that is reachable is 
                // ScrollValueMax - VisibleSize !
                // so the unreachable rest rest (VisibleSize) has to be 
                // distributed proportionally to the already traveled
                // scroll way (ScrollValue)
                int step = OoUtils.GetIntProperty(sbModel, "ScrollValue");
                int vissize = OoUtils.GetIntProperty(sbModel, "VisibleSize");
                int max = OoUtils.GetIntProperty(sbModel, "ScrollValueMax");
                int scrollable_size = max - vissize;
                double scrolledRatio = (double)step / (double)scrollable_size;
                double offset = (double)vissize * scrolledRatio;
                step += (int)offset;

                switch (o)
                {
                    case 0: // horizontal
                        scrollHorizontal(step);
                        break;
                    case 1: // vertical
                        scrollVertical(step);
                        break;
                    default:
                        break;
                }

                //switch (rEvent.Type)
                //{
                //    case AdjustmentType.ADJUST_ABS:
                //        System.Diagnostics.Debug.WriteLine("The event has been triggered by dragging the thumb...");
                //        break;
                //    case AdjustmentType.ADJUST_LINE:
                //        System.Diagnostics.Debug.WriteLine("The event has been triggered by a single line move..");
                //        break;
                //    case AdjustmentType.ADJUST_PAGE:
                //        System.Diagnostics.Debug.WriteLine("The event has been triggered by a block move...");
                //        break;
                //}
                //System.Diagnostics.Debug.WriteLine("The value of the scrollbar is: " + rEvent.Value);
            }
        }

        /// <summary>
        /// gets called when the broadcaster is about to be disposed. 
        /// All listeners and all other objects, which reference the broadcaster should 
        /// release the reference to the source. No method should be invoked anymore on 
        /// this object ( including XComponent::removeEventListener ).
        /// This method is called for every listener registration of derived listener 
        /// interfaced, not only for registrations at XComponent.
        /// </summary>
        /// <param id="Source">The source.</param>
        void unoidl.com.sun.star.lang.XEventListener.disposing(EventObject Source)
        {

        }

        #endregion

        #region XMouseListener

        void XMouseListener.mouseEntered(MouseEvent e)
        {
            //System.Diagnostics.Debug.WriteLine("#### mouse in");
        }

        void XMouseListener.mouseExited(MouseEvent e)
        {
            //System.Diagnostics.Debug.WriteLine("#### mouse out");
        }

        void XMouseListener.mousePressed(MouseEvent e)
        {
            //System.Diagnostics.Debug.WriteLine("#### mouse pressed");
        }

        void XMouseListener.mouseReleased(MouseEvent e)
        {
            //System.Diagnostics.Debug.WriteLine("#### mouse released");
        }

        #endregion

        public void Dispose()
        {
            //TODO: implement this
            XControlContainer cCont = OuterScrlContr as XControlContainer;
            //remove all element inside the outer container
            if (cCont != null)
            {
                var controls = cCont.getControls();
                foreach (var contr in controls)
                {
                    contr.dispose();
                }

                // FIXME: for fixing
                //var controls2 = cCont.getControls();

            }

            if (MXDlgModelNameContainer != null)
            {
                try
                {
                    OuterScrlContr.dispose();
                }
                catch (System.Exception) { }
                try
                {
                    (VerticalScrlBar as XControl).dispose();
                }
                catch (System.Exception) { }
                try
                {
                    (HorizontalScrlBar as XControl).dispose();
                }
                catch (System.Exception) { }
            }
        }
    }
}
