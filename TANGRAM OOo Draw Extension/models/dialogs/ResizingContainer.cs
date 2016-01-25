using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using unoidl.com.sun.star.awt;
using tud.mci.tangram.util;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.beans;

namespace tud.mci.tangram.models.dialogs
{
    public class ResizingContainer
    {

        #region Inner Container

        volatile int _innerPosX = 0;
        protected int innerPosX
        {
            get { return _innerPosX; }
            set
            {
                _innerPosX = value;
                updateContainerSize();
            }
        }

        volatile int _innerPosY = 0;
        protected int innerPosY
        {
            get { return _innerPosY; }
            set
            {
                _innerPosY = value;
                updateContainerSize();
            }
        }

        volatile int _innerWidth = 0;
        protected int innerWidth
        {
            get { return _innerWidth; }
            set
            {
                _innerWidth = value;
                updateContainerSize();
            }
        }

        volatile int _innerHeight = 0;
        protected virtual int innerHeight
        {
            get { return _innerHeight; }
            set
            {
                _innerHeight = value;
                updateContainerSize();
            }
        }


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

        //XControlContainer parentCnt;

        //XMultiServiceFactory _parentMSF = null;
        //XMultiServiceFactory parentMSF
        //{
        //    get
        //    {
        //        if (_parentMSF == null) _parentMSF = OO.GetMultiServiceFactory(MXContext);
        //        return _parentMSF;
        //    }
        //    set { _parentMSF = value; }
        //}

        public XControl innerScrlContr { get; private set; }
        XControlModel innerScrlContrModel;


        private const string AWT_UNO_CONTROL_CONTAINER = "com.sun.star.awt.UnoControlContainer";
        private const string AWT_UNO_CONTROL_CONTAINER_Model = "com.sun.star.awt.UnoControlContainerModel";



        #endregion

        public ResizingContainer(int width, int height, string sName)
        {

            #region inner Container

            // create a control model at the multi service factory of the dialog model...
            innerScrlContrModel = OO.GetMultiServiceFactory(MXContext).createInstance(AWT_UNO_CONTROL_CONTAINER_Model) as XControlModel;
            innerScrlContr = OO.GetMultiServiceFactory(MXContext).createInstance(AWT_UNO_CONTROL_CONTAINER) as XControl;

            if (innerScrlContr != null && innerScrlContrModel != null)
            {
                innerScrlContr.setModel(innerScrlContrModel);
            }

            XMultiPropertySet xinnerScrlContMPSet = innerScrlContrModel as XMultiPropertySet;

            if (xinnerScrlContMPSet != null)
            {
                xinnerScrlContMPSet.setPropertyValues(
                        new String[] { "Name", "State" },
                        Any.Get(new Object[] { sName , ((short)0) }));
            }

            innerWidth = width;
            innerHeight = height;

            // Set the size of the surrounding container
            if (innerScrlContr is XWindow)
            {
                ((XWindow)innerScrlContr).setPosSize(0, 0, innerWidth, innerHeight, PosSize.POSSIZE);
            }

            #endregion

        }

        #region Add Elements

        public XControl addElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin)
        {
            return addElementToTheEndAndAdoptTheSize(cntrl, sName, posX, topMargin, innerWidth - posX);
        }
        public XControl addElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin, int width)
        {
            return addElementToTheEndAndAdoptTheSize(cntrl, sName, posX, topMargin, width, 10);
        }
        public XControl addElementToTheEndAndAdoptTheSize(XControl cntrl, String sName, int posX, int topMargin, int width, int height)
        {
            //get all elements in the container

            if (cntrl != null)
            {
                if (innerScrlContr != null && innerScrlContr is XControlContainer)
                {
                    XControl[] ctrls = ((XControlContainer)innerScrlContr).getControls();
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

                    return addElementAndAdoptTheSize(cntrl, sName, posX, p.Height + topMargin, width, height);

                }

            }
            return cntrl;
        }

        public XControl addElementTryToKeeptheHeight(XControl cntrl, String sName, int posX, int posY, int height)
        {
            return addElementAndAdoptTheSize(cntrl, sName, posX, posY, innerWidth - posX, height);
        }

        public XControl addElementTryToKeeptheWidth(XControl cntrl, String sName, int posX, int posY, int width)
        {
            return addElementAndAdoptTheSize(cntrl, sName, posX, posY, width, 15);
        }

        public XControl addElementAndAdoptTheSize(XControl cntrl, String sName, int posX, int posY)
        {
            return addElementTryToKeeptheWidth(cntrl, sName, posX, posY, innerWidth - posX);
        }

        public XControl addElementAndAdoptTheSize(XControl cntrl, String sName, int posX, int posY, int width, int height)
        {
            Size s = new Size(width, height);
            if (cntrl is XLayoutConstrains)
            {
                s = ((XLayoutConstrains)cntrl).calcAdjustedSize(s);

                s.Width = Math.Max(width, s.Width);
                s.Height = Math.Max(height, s.Height);
            }
            return addElement(cntrl, sName, posX, posY, s.Width, s.Height);
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
        public XControl addElement(XControl cntrl, String sName, int posX, int posY, int width, int height)
        {
            //System.Diagnostics.Debug.WriteLine("added control: ____________________________");
            //Debug.GetAllInterfacesOfObject(cntrl);
            //Debug.GetAllServicesOfObject(cntrl);
            //Debug.GetAllProperties(cntrl);

            if (cntrl != null && innerScrlContr != null)
            {
                if (innerScrlContr is XControlContainer)
                {
                    ((XControlContainer)innerScrlContr).addControl(sName, cntrl);
                    OoUtils.SetProperty(cntrl, "Name", sName);
                }

                if (cntrl is XWindow)
                {
                    ((XWindow)cntrl).setPosSize(posX, posY, width, height, PosSize.POSSIZE);

                    Rectangle contrPos = ((XWindow)cntrl).getPosSize();

                    if (innerWidth < (contrPos.X + contrPos.Width)) innerWidth = contrPos.X + contrPos.Width;
                    if (innerHeight < (contrPos.Y + contrPos.Height)) innerHeight = contrPos.Y + contrPos.Height;

                }
            }

            return cntrl;
        }

        #endregion

        #region util

        #region Position and size update

        /// <summary>
        /// Updates the size of the inner control container.
        /// </summary>
        protected virtual void updateContainerSize()
        {
            if (innerScrlContr != null && innerScrlContr is XWindow)
            {
                ((XWindow)innerScrlContr).setPosSize(innerPosX, innerPosY, innerWidth, innerHeight, PosSize.POSSIZE);
            }
        }

        #endregion
        
        #endregion



    }
}
