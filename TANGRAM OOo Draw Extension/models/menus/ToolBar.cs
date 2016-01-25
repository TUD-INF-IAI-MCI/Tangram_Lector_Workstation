// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-21-2012
//
// Last Modified By : Admin
// Last Modified On : 10-11-2012
// ***********************************************************************
// <copyright file="ToolBar.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using System;
using tud.mci.tangram.util;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.graphic;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.ui;

namespace tud.mci.tangram.models.menus
{
    internal class ToolBar : AbstractStorableSettingsBase, XKeyListener
    {
        //private readonly XUIConfigurationManager _xCfgMgr;
        private readonly XModuleUIConfigurationManagerSupplier _xModCfgMgrSupplier;
        private XImageManager _xImgMgr;
        private XMultiComponentFactory _xMsf;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolBar"/> class.
        /// </summary>
        /// <param name="xMsf">the XMultiComponentFactory.</param>
        /// <param name="name">The UIName Parameter of the ToolBar.</param>
        /// <param name="documentModluleIdentifier">The document modlule identifier for the ToolBar. 
        /// E.g. "com.sun.star.text.TextDocument" to add the ToolBar to all text documents</param>
        public ToolBar(string name, string documentModluleIdentifier, XMultiComponentFactory xMsf = null)
        {
            if (name.Trim().Length < 1)
                throw new NullReferenceException("Name for ToolBar creation is empty");

            XMcf = xMsf ?? OO.GetMultiComponentFactory();
            Name = name;
            RecourceUrl = GetToolBarUrl(name);

            // Retrieve the module configuration manager supplier to  get
            // the configuration setting from the module
            _xModCfgMgrSupplier =
                (XModuleUIConfigurationManagerSupplier)
                XMcf.createInstanceWithContext(OO.Services.UI_MOD_UI_CONF_MGR_SUPPLIER, OO.GetContext());

            // Retrieve the module configuration manager
            XCfgMgr =
                _xModCfgMgrSupplier.getUIConfigurationManager(documentModluleIdentifier);

            if (XCfgMgr.hasSettings(RecourceUrl) == false)
            {
                // Creates a copy of the user interface configuration manager
                Settings = XCfgMgr.createSettings();

                // Defines the name of the toolbar 
                var xToolbarProperties = (XPropertySet)Settings;
                xToolbarProperties.setPropertyValue("UIName", Any.Get(name));

                // Makes the toolbar to the UI manager available
                try
                {
                    XCfgMgr.insertSettings(RecourceUrl, Settings);
                }
                catch (ElementExistException)
                {
                    Settings = XCfgMgr.getSettings(RecourceUrl, true) as XIndexContainer;
                }
            }
            else
            {
                Settings = XCfgMgr.getSettings(RecourceUrl, true) as XIndexContainer;
            }
        }

        private XMultiComponentFactory XMcf
        {
            get { return _xMsf ?? (_xMsf = OO.GetMultiComponentFactory()); }
            set { _xMsf = value; }
        }

        public XImageManager XImgMgr
        {
            get
            {
                if (_xImgMgr == null)
                {
                    if (XCfgMgr != null)
                    {
                        _xImgMgr = (XImageManager)XCfgMgr.getImageManager();
                    }
                    else
                    {
                        return null;
                    }
                }
                return _xImgMgr;
            }
        }

        /// <summary>
        /// Gets the UIName Parameter of the ToolBar.
        /// </summary>
        /// <value>The name.</value>
        public string Name { get; private set; }

        /// <summary>
        /// Returns a valid resource url for the tool bar.
        /// </summary>
        /// <param name="name">The name of the ToolBar.</param>
        /// <returns>A recource url for the ToolBar depending on the given name <code>private:resource/toolbar/custom_[NAME]</code></returns>
        private static string GetToolBarUrl(string name)
        {
            string url = "private:resource/toolbar/custom_" + name.Replace(" ", "");
            return url;
        }

        /// <summary>
        /// Stores the tool bar persistantly in the OOo-user-configuration.
        /// </summary>
        public void Store()
        {
            var xPersistence = XCfgMgr as XUIConfigurationPersistence;
            if (xPersistence != null)
                xPersistence.store();
        }

        /// <summary>
        /// Sets the item property.
        /// </summary>
        /// <param name="commandId">The command id.</param>
        /// <param name="label">The label.</param>
        /// <param name="piEItemType"> </param>
        /// <returns></returns>
        public static PropertyValue[] SetItemProperty(string commandId, string label,
                                                      OO.PropertyControlType piEItemType = 0)
        {
            var itemProperty = new PropertyValue[4];
            itemProperty[0] = new PropertyValue { Name = "CommandURL", Value = Any.Get(commandId) };
            itemProperty[1] = new PropertyValue { Name = "Label", Value = Any.Get(label) };
            itemProperty[2] = new PropertyValue { Name = "Type", Value = Any.Get((short)piEItemType) };
            itemProperty[3] = new PropertyValue { Name = "IsVisible", Value = Any.Get(true) };
            return itemProperty;
        }

        /// <summary>
        /// Adds an image button to the ToolBar.
        /// </summary>
        /// <param name="aCommandUrl">A command URL e.g. "macro:///Standard.Module1.myMacro()"</param>
        /// <param name="aGraphic">A graphic for the Button. Use the <code>XImgMgr</code> to retrieve or register the image.</param>
        /// <param name="itemLabel">The item label.</param>
        /// <param name="dublicate">Set <code>true</code> if you want to add this button even if the commandUrl allready exist in the ToolBar</param>
        /// <param name="index">Index for the position of the Button. Postitive value means normal index, negative value means position from the the end of all buttons, <code>-0</code> means at the and of the list. </param>
        /// <param name="buttonType"> </param>
        /// <param name="nImageType">Type of the image. Standard is 0</param>
        /// <returns></returns>
        public bool AddImgButton(string aCommandUrl, XGraphic aGraphic, string itemLabel, bool dublicate = false, int index = -0,
                                 OO.PropertyControlType buttonType = 0, short nImageType = 0)
        {
            return AddImgButton(new[] { aCommandUrl }, new[] { aGraphic }, new[] { itemLabel }, dublicate, new[] { index },
                                buttonType, nImageType);
        }

        /// <summary>
        /// Adds the img button.
        /// </summary>
        /// <param name="aCommandUrlSequence">A command URLs e.g. "macro:///Standard.Module1.myMacro()"</param>
        /// <param name="aGraphicSequence">A graphics for the Button. Use the <code>XImgMgr</code> to retrieve or register the image.</param>
        /// <param name="itemLabelSequence">The item labels.</param>
        /// <param name="dublicate">Set <code>true</code> if you want to add this button even if the commandUrl allready exist in the ToolBar</param>
        /// <param name="index"> </param>
        /// <param name="buttonType"> </param>
        /// <param name="nImageType">Type of the images. Standard is 0</param>
        /// <returns></returns>
        public bool AddImgButton(string[] aCommandUrlSequence, XGraphic[] aGraphicSequence, string[] itemLabelSequence, bool dublicate = false,
                                 int[] index = null, OO.PropertyControlType buttonType = 0, short nImageType = 0)
        {
            if (XImgMgr == null) return false;

            XImgMgr.insertImages(nImageType, aCommandUrlSequence, aGraphicSequence);

            for (int i = 0; i < aCommandUrlSequence.Length; i++)
            {
                if (!dublicate && ButtonAllreadyExists(aCommandUrlSequence[i]))
                    continue;

                int z = GetIndex(index, i);

                // Creates the PropertyValue set for the toolbar
                PropertyValue[] toolbarItemProperty = SetItemProperty(aCommandUrlSequence[i], itemLabelSequence[i],
                                                                      buttonType);

                // Adds the new toolbar to the toolbar container
                Settings.insertByIndex(z, Any.GetAsOne(toolbarItemProperty));
            }

            StoreSettings();

            return true;
        }

        private bool ButtonAllreadyExists(string commandUrl)
        {
            XIndexContainer mOoToolbarSettings = Settings;
            if (mOoToolbarSettings == null)
                return false;
            int iCount = mOoToolbarSettings.getCount();
            for (int i = 0; i < iCount; i++)
            {
                Array ooToolbarButton = mOoToolbarSettings.getByIndex(i).Value as Array ?? new Array[0];

                int jCount = ooToolbarButton.GetUpperBound(0);
                for (int j = 0; j < jCount; j++)
                {
                    var pv = ((PropertyValue)ooToolbarButton.GetValue(j));
                    if (pv.Name.Equals("CommandURL"))
                    {
                        if (pv.Value.Value.Equals(commandUrl))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private int GetIndex(int[] index, int i)
        {
            int z = 0;
            if (index != null)
                if (index.Length > i)
                    if (index[i] == -0)
                        z = Settings.getCount();
                    else if (index[i] < 0)
                        z = Settings.getCount() - 1 - index[i];
                    else
                        z = index[i];
            z = z < 0 ? 0 : z;
            return z;
        }

        /// <summary>
        /// Disables the button.
        /// </summary>
        /// <param name="commandUrl">The command URL.</param>
        public void DisableButton(String commandUrl)
        {
            SetButtonPropery(commandUrl, "IsVisible", false);
        }

        /// <summary>
        /// Enables the button.
        /// </summary>
        /// <param name="commandUrl">The command URL.</param>
        public void EnableButton(String commandUrl)
        {
            SetButtonPropery(commandUrl, "IsVisible", true);
        }

        /// <summary>
        /// Sets the button propery.
        /// </summary>
        /// <param name="commandUrl">The command URL.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="value">The value.</param>
        public void SetButtonPropery(String commandUrl, string propertyName, object value)
        {
            XIndexContainer mOoToolbarSettings = Settings;
            if (mOoToolbarSettings == null)
                return;

            bool bButtonFound = false;
            bool save = false;

            int iCount = mOoToolbarSettings.getCount();
            for (int i = 0; i < iCount; i++)
            {
                Array ooToolbarButton = mOoToolbarSettings.getByIndex(i).Value as Array ?? new Array[0];

                int jCount = ooToolbarButton.GetUpperBound(0);
                for (int j = 0; j < jCount; j++)
                {
                    var pv = ((PropertyValue)ooToolbarButton.GetValue(j));
                    if (pv.Name.Equals("CommandURL"))
                    {
                        if (pv.Value.Value.Equals(commandUrl))
                        {
                            bButtonFound = true;
                        }
                    }
                    else if (pv.Name.Equals(propertyName) && bButtonFound && (bool)pv.Value.Value)
                    {
                        pv.Value = Any.Get(value);
                        bButtonFound = false;
                        save = true;
                    }
                    else if (j + 1 >= jCount)
                    {
                        ooToolbarButton.SetValue(new PropertyValue { Name = propertyName, Value = Any.Get(value) }, jCount);
                        bButtonFound = false;
                        save = true;
                    }
                }
                if (save)
                    mOoToolbarSettings.replaceByIndex(i, Any.GetAsOne(ooToolbarButton));
            }
            if (save)
                Settings = mOoToolbarSettings;
        }




        public XWindow createItemWindow(XWindow xWindow)
        {
            // xMSF is set by initialize(Object[])
            try
            {
                // get XWindowPeer
                XWindowPeer xWinPeer = (XWindowPeer)xWindow;
                Object o = _xMsf.createInstanceWithContext("com.sun.star.awt.Toolkit", OO.GetContext());
                XToolkit xToolkit = (XToolkit)o;
                // create WindowDescriptor
                WindowDescriptor wd = new WindowDescriptor();
                wd.Type = WindowClass.SIMPLE;
                wd.Parent = xWinPeer;
                wd.Bounds = new Rectangle(0, 0, 20, 23);
                wd.ParentIndex = (short)-1;
                wd.WindowAttributes = WindowAttribute.SHOW;
                wd.WindowServiceName = "pushbutton";
                // create Button
                XWindowPeer cBox_xWinPeer = xToolkit.createWindow(wd);// null here
                var xButton = (XButton)cBox_xWinPeer;
                xButton.setLabel("My Texte");
                XWindow cBox_xWindow = (XWindow)cBox_xWinPeer;
                cBox_xWindow.addKeyListener(this);

                return cBox_xWindow;

            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("createItemWindow left not o.k.\n" + e);
                return null;
            }
        }






        public void keyPressed(KeyEvent e)
        {
            throw new NotImplementedException();
        }

        public void keyReleased(KeyEvent e)
        {
            throw new NotImplementedException();
        }

        public void disposing(EventObject Source)
        {
            throw new NotImplementedException();
        }
    }
}