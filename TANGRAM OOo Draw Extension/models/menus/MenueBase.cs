// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-17-2012
//
// Last Modified By : Admin
// Last Modified On : 09-17-2012
// ***********************************************************************
// <copyright file="MenueBase.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using System;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.awt;
using tud.mci.tangram.util;

namespace tud.mci.tangram.models.menus
{
    abstract class MenueBase
    {
        protected XComponentContext _xContext;
        protected XMultiComponentFactory _xMCF;
        protected XTopWindow _xWindow;
        protected XMenuListener _xMenuListener;
        protected XMenuBar _xMenuBar;

        protected MenueBase(XComponentContext xContext, XMultiComponentFactory xMcf, XTopWindow xWindow, XMenuListener xMListener, XMenuBar xMenuBar = null)
        {
            _xContext = xContext;
            _xMCF = xMcf;
            _xWindow = xWindow;
            _xMenuListener = xMListener;
            _xMenuBar = xMenuBar != null ? xMenuBar : AddMenuBar();
        }

        protected XMenuBar AddMenuBar()
        {
            // create a menubar at the global MultiComponentFactory...
            Object oMenuBar = _xMCF.createInstanceWithContext(OO.Services.VCLX_MENU_BAR, _xContext);
            // add the menu items...
            XMenuBar xMenuBar = oMenuBar as XMenuBar;
            if (xMenuBar != null)
            {
                xMenuBar.addMenuListener(_xMenuListener);
                _xWindow.setMenuBar(xMenuBar);
                return xMenuBar;
            }
            else
            {
                //TODO: handle if the cast fails
                throw new NullPointerException("Cast to XMenuBar fails:", oMenuBar);
            }

        }

        public bool AddPopupMenuToMenueEntry(short nId, XPopupMenu xPopupMenu)
        {
            if (_xMenuBar.getItemPos(nId) >= 0)
            {
                try
                {
                    _xMenuBar.setPopupMenu(nId, xPopupMenu);
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Exeption in addPopupMenuToMenueEntry: " + ex);
                    return false;
                }
                return true;
            }
            else
            {
                //throw new NullPointerException("node id " + nId + " does not exist in Menue: ", _xMenuBar);
                return false;
            }

        }


        public bool InsertItem(short id, string title, short style = unoidl.com.sun.star.awt.MenuItemStyle.AUTOCHECK, short pos = -1)
        {
            try
            {
                _xMenuBar.insertItem(id, title, style, pos < 0 ? _xMenuBar.getItemCount() : pos);
            }
            catch (System.Exception ex )
            {
                System.Diagnostics.Debug.WriteLine("Exeption in insertItem: " + ex);
                 return false;
            }
            return true;
        }



    }
}
