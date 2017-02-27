// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extension
// Author           : Admin
// Created          : 09-17-2012
//
// Last Modified By : Admin
// Last Modified On : 09-21-2012
// ***********************************************************************
// <copyright file="OOo.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using System;
using uno.util;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using Exception = System.Exception;
using System.Diagnostics;
using System.Threading;


namespace tud.mci.tangram.util
{
    /// <summary>
    /// Basic API wrapper functions and definitions
    /// </summary>
    public static class OO
    {
        /// <summary>
        /// The uno command URL path to the tangram AddOn
        /// </summary>
        public const string UNO_COMMAND_URL_PATH = "org.openoffice.Office.addon.tangram:";

        #region BASE COMPONENTS

        private static XComponentContext _context = null;

        private static readonly object _contextCreationLock = new Object();

        /// <summary>
        /// Gets the context.
        /// </summary>
        /// <returns></returns>
        public static XComponentContext GetContext(bool renew = false)
        {
            if (_context == null || renew)
            {
                lock (_contextCreationLock)
                {
                    // if the lock releases and a context exists
                    if (_context != null) { return _context; }
                    try
                    {
                        bool success = false;
                        //TODO: timeout
                        //success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(5000,
                        //     () =>
                        //     {
                                 try
                                 {
                                     _context = Bootstrap.bootstrap();
                                     addListener(_context);
                                     success = true;
                                 }
                                 catch (System.Threading.ThreadAbortException)
                                 {
                                     System.Diagnostics.Debug.WriteLine("[ERROR]  getting Bootstrap timed out");
                                 }
                                 catch (System.Exception ex)
                                 {
                                     System.Diagnostics.Debug.WriteLine("[FATAL ERROR]  cannot get Bootstrap: " + ex);
                                     Logger.Instance.Log(LogPriority.IMPORTANT, "OO", "[FATAL ERROR] Can not get connection to OpenOffec by bootstap: " + ex);
                                 }
                             //},
                             //"Bootstrap-ContextLoader"
                             //);

                        if (!success)
                        {
                            _context = null;
                        }                        
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Can't create context: " + ex);
                        return null;
                    }
                }
            }
            return _context;
        }

        static XMultiComponentFactory _xMcf;
        /// <summary>
        /// Gets the Multi Component Factory (ServiceManager).
        /// </summary>
        /// <param name="xContext">The context.</param>
        /// <returns>A related multi component factory for instantiating new objects by their service name.</returns>
        /// <remarks>This function is time limited to 300 ms.</remarks>
        public static XMultiComponentFactory GetMultiComponentFactory(XComponentContext xContext = null, bool renew = false)
        {
            if (_xMcf != null && !renew)
                return _xMcf;
            try
            {
                if (xContext == null)
                {
                    xContext = GetContext();
                }

                if (xContext != null)
                {
                    TimeLimitExecutor.WaitForExecuteWithTimeLimit(300, () =>
                    {
                        _xMcf = xContext.getServiceManager();
                        addListener(_xMcf);
                    }, "GetMultiComponentFactory");

                    return _xMcf;
                }
                else
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "OO","[ERROR] Could not create a XMultiComponentFactory because the XContext is NULL");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Can't create ServiceManager factory: " + ex);
            }
            return null;
        }

        static XDesktop _xDesktop = null;
        /// <summary>
        /// Gets the desktop.
        /// </summary>
        /// <param name="xMcf">The x MCF.</param>
        /// <param name="xContext">The x context.</param>
        /// <returns></returns>
        public static XDesktop GetDesktop(XMultiComponentFactory xMcf = null, XComponentContext xContext = null)
        {
            if (_xDesktop == null)
            {
                if (xContext == null)
                    xContext = GetContext();

                if (xMcf == null)
                    xMcf = GetMultiComponentFactory(xContext);
                if (xMcf != null)
                {
                    try
                    {
                        object oDesktop = xMcf.createInstanceWithContext(Services.FRAME_DESKTOP, xContext);
                        _xDesktop = oDesktop as XDesktop;
                        addListener(_xDesktop);

                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine("Can't create xDesktop: " + e);
                    }
                }
            }
            return _xDesktop;
        }

        static XExtendedToolkit _xToolkit;
        /// <summary>
        /// Gets the ExtendedToolkit.
        /// </summary>
        /// <param name="xMcf">The ServiceManager.</param>
        /// <param name="xContext">The context.</param>
        /// <returns></returns>
        public static XExtendedToolkit GetExtTooklkit(XMultiComponentFactory xMcf = null,
                                                      XComponentContext xContext = null)
        {
            if (_xToolkit == null)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "Oo.BasicObjects", "renew XExtendedToolkit");
                try
                {
                    if (xContext == null)
                    {
                        xContext = GetContext();
                    }

                    if (xMcf == null)
                    {
                        xMcf = GetMultiComponentFactory(xContext);
                    }

                    object toolkit = xMcf.createInstanceWithContext(Services.VCLX_EXT_TOOLKIT, xContext);
                    _xToolkit = toolkit as XExtendedToolkit;
                    addListener(_xToolkit);

                }
                catch (Exception ex)
                {
                    Logger.Instance.Log(LogPriority.ALWAYS, "Can't create ExtendedToolkit ", ex);
                    return null;
                }
            }
            return _xToolkit;
        }

        static XMultiServiceFactory _xMsf;
        /// <summary>
        /// Gets the multi service factory.
        /// DEPRECATED: for libre office this is always <c>NULL</c>. Use <see cref="XMultiComponentFactory"/> instead.
        /// </summary>
        /// <param name="xCompContext">The x comp context.</param>
        /// <param name="xMcf">The x MCF.</param>
        /// <returns></returns>
        public static XMultiServiceFactory GetMultiServiceFactory(XComponentContext xCompContext = null,
                                                                  XMultiComponentFactory xMcf = null)
        {
            if (_xMsf == null)
            {
                if (xCompContext == null)
                {
                    xCompContext = GetContext();
                }
                if (xMcf == null)
                {
                    xMcf = GetMultiComponentFactory(xCompContext);
                }

                if (_xMsf == null)
                {
                    if (xCompContext != null && xCompContext is XComponentContext)
                    {
                        _xMsf = ((XComponentContext)xCompContext).getServiceManager() as XMultiServiceFactory;
                    }
                }

                if (_xMsf == null) {
                    _xMsf = xMcf.createInstanceWithContext(Services.MULTI_SERVICE_FACTORY, xCompContext) as XMultiServiceFactory;
                }

                addListener(_xMsf);

            }
            return _xMsf;
        }

        private static readonly Object conCheckLock = new Object();
        /// <summary>
        /// Checks the current connection to Office application.
        /// resets the connection if necessary.
        /// </summary>
        /// <returns>always <c>false</c></returns>
        /// <remarks>This function is time limited to 50 ms.</remarks>
        public static bool CheckConnection()
        {
            lock (conCheckLock)
            {
                //how to check the connection
                Logger.Instance.Log(LogPriority.OFTEN, "OO", "Request for OpenOffice connection check");
                try
                {
                    if (GetDesktop() != null)
                    {

                        try
                        {
                            TimeLimitExecutor.WaitForExecuteWithTimeLimit(50,
                                () =>{
                                    var desktop = GetDesktop();
                                    desktop.getCurrentFrame();
                                }, "CheckDesktop");
                        }
                        catch
                        {
                            reset();
                        }
                    }
                    else
                    {
                        reset();
                    }
                }
                catch (System.Exception)
                {
                    reset();
                }

                return false;
            }
        }

        private static object syncLock = new Object();

        /// <summary>
        /// Tries to establish a connection to OpenOffice / LibreOffice
        /// </summary>
        /// <returns><c>true</c> if a connection could be established; otherwise <c>false</c>.</returns>
        public static bool ConnectToOO()
        {
            bool success = false;
            lock (syncLock)
            {
                var cont = OO.GetContext();
                XMultiComponentFactory xMcf = OO.GetMultiComponentFactory(cont);
                success = xMcf != null;

                //if (!success)
                //{
                //    cont = OO.GetContext(true);
                //} 
            }

            return success;
        }

        #endregion

        /// <summary>
        /// Adds a general event listener to the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        private static void addListener(object obj)
        {
            if (obj != null)
            {
                var dl = new DisposeListener(obj);
                if (obj is XComponent)
                    ((XComponent)obj).addEventListener(dl);
                if (obj is XDesktop)
                    ((XDesktop)obj).addTerminateListener(dl);
            }

        }

        #region Base Object Termination Listener and Event Throwing

        /// <summary>
        /// Occurs when [base object disposed]. Indicates that the last 
        /// open regular window of OpenOffice has been closed.
        /// This also indicates that the global listeners and the 
        /// context (bootstrap) will not longer work properly!
        /// </summary>
        public static event EventHandler<EventArgs> BaseObjectDisposed;

        static void fireBaseObjectDisposed(Object obj)
        {
            if (BaseObjectDisposed != null)
            {
                try
                {
                    BaseObjectDisposed.DynamicInvoke(obj, new EventArgs());
                }
                catch (System.Exception){}
            }
        }

        /// <summary>
        /// Resets this instance. And kills the OpenOffice process. 
        /// BE CAREFULL!!
        /// </summary>
        static void reset()
        {
            Logger.Instance.Log(LogPriority.OFTEN, "OO", "Request for OpenOffice connection reset");
            try
            {
                _context = null;
                kill();
                _xMcf = null;
                _xDesktop = null;
                _xToolkit = null;
               // _xMsf = null;
                _xDesktop = GetDesktop();

                addListener(GetContext());

            }
            catch (System.Exception){ }
        }

        private static void kill()
        {
            //FIXME: kill open office processes 
            // is a very bad hack but necessary!!!
            bool success = FindAndKillAllProcesses("soffice");
        }

        /// <summary>
        /// Finds the and kill all processes stating with the given name.
        /// </summary>
        /// <author>based on post by 'Vova Popov'</author>
        /// <param name="name">The name the process name hase to start.</param>
        /// <returns></returns>
        public static bool FindAndKillAllProcesses(string name)
        {

            bool success = false;

            //here we're going to get a list of all running processes on
            //the computer
            Process[] locals = Process.GetProcesses();
            foreach (Process clsProcess in locals)
            {
                if (clsProcess.ProcessName.StartsWith(name))
                {
                    try
                    {
                        string pName = clsProcess.ProcessName;
                        // Close process by sending a close message to its main window.
                        clsProcess.CloseMainWindow();
                        // Free resources associated with process.
                        clsProcess.Close();

                        //check if really closed, otherwise kill
                        Process[] localByName = Process.GetProcessesByName(pName);
                        foreach (var process in localByName)
                        {
                            try
                            {
                                process.Kill();
                                Thread.Sleep(50);
                            }
                            catch
                            {
                                success = false;
                            }
                        }
                    }
                    catch
                    {
                        success = false;
                    }
                    //process killed, return true
                    success = true;
                }
            }
            //process not found, return false
            return success;
        }

        /// <summary>
        /// Class that listens if one basic and important object of the 
        /// open office connection disposes and handle this bad news.
        /// </summary>
        private class DisposeListener : unoidl.com.sun.star.lang.XEventListener, XTerminateListener
        {
            Object obj;
            public DisposeListener(Object obj)
            {
                this.obj = obj;
            }

            void unoidl.com.sun.star.lang.XEventListener.disposing(unoidl.com.sun.star.lang.EventObject Source) { }

            void XTerminateListener.notifyTermination(unoidl.com.sun.star.lang.EventObject Event)
            {
                if (Event != null && Event.Source != null)
                {
                    //util.Debug.GetAllInterfacesOfObject(Event.Source);
                    if (Event.Source is XDesktop)
                    {
                        // if the desktop disposes - the whole system will fail to work. 
                        // so reset all important elements and kill the rest of the OpenOffice process.
                        // This is necessary because the bootstrap loader of the api can not register a 
                        // new namedpipe to the OO-process. This results in a not longer working 
                        // connection for listeners or api-functions!
                        reset();
                    }
                }

                fireBaseObjectDisposed(obj);
            }

            void XTerminateListener.queryTermination(unoidl.com.sun.star.lang.EventObject Event) { }
        }

        #endregion

        #region Const Classes & Enums

        /// <summary>
        /// known service strings for element creation via XMultiServiceFactory
        /// </summary>
        public static class Services
        {
            public const String VCLX_EXT_TOOLKIT = "stardiv.Toolkit.VCLXToolkit";
            public const String VCLX_MENU_BAR = "stardiv.Toolkit.VCLXMenuBar";

            /// <summary>
            /// Provides a collection of implementations of services. The factories for instantiating objects of implemetations are accessed via a service name.
            /// </summary>
            public const String MULTI_SERVICE_FACTORY = "com.sun.star.lang.MultiServiceFactory";

            /// <summary>
            /// Provides an easy way to dispatch an URL using one call instead of multiple ones.
            /// </summary>
            public const String DISPATCH_HELPER = "com.sun.star.frame.DispatchHelper";
            
            /// <summary>
            /// abstract service which specifies a storable and printable document
            /// </summary>
            public const String DOCUMENT = "com.sun.star.document.OfficeDocument";
            /// <summary>
            /// Specify the document service of the text module.
            /// </summary>
            public const String DOCUMENT_TEXT = "com.sun.star.text.TextDocument";
            /// <summary>
            /// deprecated -- Specify the document service of the web module.
            /// </summary>
            public const String DOCUMENT_WEB = "com.sun.star.text.WebDocument";            
            /// <summary>
            /// specifies a document which consists of multiple pages with drawings. Because its function is needed more then once, its defined as generic one.
            /// </summary>
            public const String DOCUMENT_DRAWING_GENERIC = "com.sun.star.drawing.GenericDrawingDocumen";
            /// <summary>
            /// deprecated -- Pleas use the factory interface of the service GenericDrawingDocument.
            /// </summary>
            public const String DOCUMENT_DRAWING_DOCUMENT_FACTORY = "com.sun.star.drawing.DrawingDocumentFactory";
            /// <summary>
            /// specifies a document which consists of multiple pages with drawings.
            /// </summary>
            public const String DOCUMENT_DRAWING = "com.sun.star.drawing.DrawingDocument";
            /// <summary>
            /// represents a model component which consists of some settings and one or more spreadsheets.
            /// </summary>
            public const String DOCUMENT_SPREADSHEET = "com.sun.star.sheet.SpreadsheetDocument";
            /// <summary>
            /// factory to create filter components.
            /// </summary>
            public const String DOCUMENT_FILTER_FACTORY = "com.sun.star.document.FilterFactory";

            /// <summary>
            /// describes a toolkit that creates windows on a screen.
            /// </summary>
            public const String AWT_TOOLKIT = "com.sun.star.awt.Toolkit";

            /// <summary>
            /// specifies the standard model of an UnoControlButton.
            /// </summary>
            public const String AWT_CONTROL_BUTTON_MODEL = "com.sun.star.awt.UnoControlButtonModel";
            /// <summary>
            /// specifies the standard model of an UnoControlCheckBox.
            /// </summary>
            public const String AWT_CONTROL_CHECKBOX_MODEL = "com.sun.star.awt.UnoControlCheckBoxModel";
            /// <summary>
            /// specifies the standard model of an UnoControlComboBox.
            /// </summary>
            public const String AWT_CONTROL_COMBOBOX_MODEL = "com.sun.star.awt.UnoControlComboBoxModel";
            /// <summary>
            /// specifies the standard model of an UnoControlCurrencyField.
            /// </summary>
            public const String AWT_CONTROL_CURRENCY_FIELD_MODEL = "com.sun.star.awt.UnoControlCurrencyFieldModel";
            /// <summary>
            /// specifies the standard model of an UnoControlDateField.
            /// </summary>
            public const String AWT_CONTROL_DATE_FIELD_MODEL = "com.sun.star.awt.UnoControlDateFieldModel";
            /// <summary>
            /// specifies a dialog control.
            /// </summary>
            public const String AWT_CONTROL_DIALOG = "com.sun.star.awt.UnoControlDialog";
            /// <summary>
            /// specifies the standard model of an UnoControlDialog.
            /// </summary>
            public const String AWT_CONTROL_DIALOG_MODEL = "com.sun.star.awt.UnoControlDialogModel";
            /// <summary>
            /// specifies the standard model of an UnoControlEdit.
            /// </summary>
            public const String AWT_CONTROL_EDIT_MODEL = "com.sun.star.awt.UnoControlEditModel";
            /// <summary>
            /// specifies the standard model of an UnoControlFileControl.
            /// </summary>
            public const String AWT_CONTROL_FILE_MODEL = "com.sun.star.awt.UnoControlFileControlModel";
            /// <summary>
            /// specifies the standard model of an UnoControlFixedLine.
            /// </summary>
            public const String AWT_CONTROL_FIXED_LINE_MODEL = "com.sun.star.awt.UnoControlFixedLineModel";
            /// <summary>
            /// specifies a fixed line control.
            /// </summary>
            public const String AWT_CONTROL_FIXED_LINE = "com.sun.star.awt.UnoControlFixedLine";
            /// <summary>
            /// specifies the standard model of an UnoControlFormattedField.
            /// </summary>
            public const String AWT_CONTROL_FORMATTED_FIELD_MODEL = "com.sun.star.awt.UnoControlFormattedFieldModel";
            /// <summary>
            /// specifies the standard model of a UnoControlGrid control.
            /// </summary>
            public const String AWT_CONTROL_GRID_MODEL = "com.sun.star.awt.UnoControlGridModel";
            /// <summary>
            /// specifies the standard model of an UnoControlGroupBox.
            /// </summary>
            public const String AWT_CONTROL_GROUP_BOX_MODEL = "com.sun.star.awt.UnoControlGroupBoxModel";
            /// <summary>
            /// specifies the standard model of an UnoControlImageControl.
            /// </summary>
            public const String AWT_CONTROL_IMAGE_MODEL = "com.sun.star.awt.UnoControlImageControlModel";
            /// <summary>
            /// specifies the standard model of an UnoControlListBox.
            /// </summary>
            public const String AWT_CONTROL_LISTBOX_MODEL = "com.sun.star.awt.UnoControlListBoxModel";
            /// <summary>
            /// specifies the standard model of an UnoControlNumericField.
            /// </summary>
            public const String AWT_CONTROL_NUMMERIC_FIELD_MODEL = "com.sun.star.awt.UnoControlNumericFieldModel";
            /// <summary>
            /// specifies the standard model of an UnoControlPatternField.
            /// </summary>
            public const String AWT_CONTROL_PATTERNFIELD_MODEL = "com.sun.star.awt.UnoControlPatternFieldModel";
            /// <summary>
            /// specifies the standard model of an UnoControlProgressBar.
            /// </summary>
            public const String AWT_CONTROL_PROGRESS_BAR_MODEL = "com.sun.star.awt.UnoControlProgressBarModel";
            /// <summary>
            /// specifies a progress bar control.
            /// </summary>
            public const String AWT_CONTROL_PROGRESS_BAR = "com.sun.star.awt.service UnoControlProgressBar";
            /// <summary>
            /// specifies the standard model of an UnoControlRadioButton.
            /// </summary>
            public const String AWT_CONTROL_RADIOBUTTON_MODEL = "com.sun.star.awt.UnoControlRadioButtonModel";
            /// <summary>
            /// specifies the standard model of an UnoControlContainer.
            /// </summary>
            public const String AWT_CONTROL_ROADMAP_MODEL = "com.sun.star.awt.UnoControlRoadmapModel";
            /// <summary>
            /// specifies the standard model of an UnoControlScrollBar.
            /// </summary>
            public const String AWT_CONTROL_SCROLLBAR_MODEL = "com.sun.star.awt.UnoControlScrollBarModel";
            /// <summary>
            /// specifies a model for a UnoControlTabPageContainer control.
            /// </summary>
            public const String AWT_CONTROL_TABPAGE_CONTAINER_MODEL = "com.sun.star.awt.UnoControlTabPageContainerModel";
            /// <summary>
            /// specifies the standard model of a XTabPageModel.
            /// </summary>
            public const String AWT_CONTROL_TABPAGE_MODEL = "com.sun.star.awt.UnoControlTabPageModel";
            /// <summary>
            /// specifies the standard model of an UnoControlFixedText.
            /// </summary>
            public const String AWT_CONTROL_TEXT_FIXED_MODEL = "com.sun.star.awt.UnoControlFixedTextModel";
            /// <summary>
            /// specifies a control for displaying fixed text.
            /// </summary>
            public const String AWT_CONTROL_TEXT_FIXED = "com.sun.star.awt.UnoControlFixedText";
            /// <summary>
            /// specifies the standard model of an UnoControlTimeField.
            /// </summary>
            public const String AWT_CONTROL_TIME_FIELD_MODEL = "com.sun.star.awt.UnoControlTimeFieldModel";
            /// <summary>
            /// specifies the standard model of a TreeControl.
            /// </summary>
            public const String AWT_CONTROL_TREE_MODEL = "com.sun.star.awt.TreeControlModel";
            /// <summary>
            /// If you do not want to implement the XTreeDataModel yourself, use this service. This implementation uses MutableTreeNode for its nodes.
            /// </summary>
            public const String AWT_CONTROL_MUTABLETREE_MODEL = "com.sun.star.awt.tree.MutableTreeDataModel";

            /// <summary>
            /// Central service of the Graphic API that gives access to graphics of any kind. This service allows to load graphics from and to store graphics to any location. 
            /// </summary>
            public const String GRAPHIC_GRAPHICPROVIDER = "com.sun.star.graphic.GraphicProvider";

            /// <summary>
            /// specifies accessibility support for a menu separator.
            /// </summary>
            public const String AWT_ACCESSIBILITY_MENU_SEPERATOR = "com.sun.star.awt.AccessibleMenuSeparator";
            /// <summary>
            /// specifies accessibility support for a menu.
            /// </summary>
            public const String AWT_ACCESSIBILITY_MENU = "com.sun.star.awt.AccessibleMenu";
            /// <summary>
            /// specifies accessibility support for a menu bar.
            /// </summary>
            public const String AWT_ACCESSIBILITY_MENUBAR = "com.sun.star.awt.AccessibleMenuBar";
            /// <summary>
            /// specifies accessibility support for a menu item.
            /// </summary>
            public const String AWT_ACCESSIBILITY_MENU_ITEM = "com.sun.star.awt.AccessibleMenuItem";
            /// <summary>
            /// specifies accessibility support for a window.
            /// </summary>
            public const String AWT_ACCESSIBILITY_WINDOW = "com.sun.star.awt.AccessibleWindow";
            /// <summary>
            /// specifies accessibility support for a popup menu.
            /// </summary>
            public const String AWT_ACCESSIBILITY_POPUP_MENU = "com.sun.star.awt.AccessiblePopupMenu";
            /// <summary>
            /// specifies accessibility support for a fixed text.
            /// </summary>
            public const String AWT_ACCESSIBILITY_FIXEDTEXT = "com.sun.star.awt.AccessibleFixedText";
            /// <summary>
            /// specifies accessibility support for an edit.
            /// </summary>
            public const String AWT_ACCESSIBILITY_EDIT = "com.sun.star.awt.AccessibleEdit";
            /// <summary>
            /// specifies accessibility support for a button.
            /// </summary>
            public const String AWT_ACCESSIBILITY_BUTTON = "com.sun.star.awt.AccessibleButton";
            /// <summary>
            /// specifies accessibility support for a scroll bar.
            /// </summary>
            public const String AWT_ACCESSIBILITY_SCROLLBAR = "com.sun.star.awt.AccessibleScrollBar";
            /// <summary>
            /// specifies accessibility support for a tab page.
            /// </summary>
            public const String AWT_ACCESSIBILITY_TAB_PAGE = "com.sun.star.awt.AccessibleTabPage";
            /// <summary>
            /// specifies accessibility support for a tab control.
            /// </summary>
            public const String AWT_ACCESSIBILITY_TAB_CONTROL = "com.sun.star.awt.AccessibleTabControl";

            public const String AWT_POINTER = "com.sun.star.awt.Pointer";

            /// <summary>
            /// provides an supplier of number formats
            /// </summary>
            public const String UTIL_FORMAT_NUMBER = "com.sun.star.util.NumberFormatsSupplier";
            /// <summary>
            /// Supports read/write access and listener for the paths properties that the Office uses.
            /// The property names of the Office paths/directories are an exactly match to the configuration entries found in the file (org/openoffice/Office/Common.xml).
            /// This service supports the usage of path variables to define paths that a relative to other office or system directories.
            /// </summary>
            public const String UTIL_PATH_SETTINGS = "com.sun.star.util.PathSettings";
            /// <summary>
            /// The File Content Provider (FCP) implements a ContentProvider for the UniversalContentBroker (UCB).
            /// The served contents enable access to the local file system.
            /// The FCP is able to restrict access to the local file system to a number of directories shown to the client under configurable aliasnames.
            /// </summary>
            public const String UTIL_FILE_CONTENT_PRVIDER = "com.sun.star.ucb.FileContentProvider";
            /// <summary>
            /// helps to split up a string containing a URL into its structural parts and assembles the parts into a single string.
            /// </summary>
            public const String UTIL_URL_TRANSFORMER = "com.sun.star.util.URLTransformer";

            /// <summary>
            /// specifies a central user interface configuration provider which gives access to module based user interface configuration managers.
            /// </summary>
            public const String UI_MOD_UI_CONF_MGR_SUPPLIER = "com.sun.star.ui.ModuleUIConfigurationManagerSupplier";

            /// <summary>
            /// This abstract service specifies the general characteristics of all Shapes.
            /// </summary>
            public const String DRAW_SHAPE = "com.sun.star.drawing.Shape";
            /// <summary>
            /// This service is for a rectangle Shape.
            /// </summary>
            public const String DRAW_SHAPE_RECT = "com.sun.star.drawing.RectangleShape";
            /// <summary>
            /// This service is for an ellipse or circle shape.
            /// </summary>
            public const String DRAW_SHAPE_ELLIPSE = "com.sun.star.drawing.EllipseShape";
            /// <summary>
            /// This service is for a simple Shape with lines.
            /// </summary>
            public const String DRAW_SHAPE_LINE = "com.sun.star.drawing.LineShape";
            /// <summary>
            /// This service is for a text shape.
            /// </summary>
            public const String DRAW_SHAPE_TEXT = "com.sun.star.drawing.TextShape";
            /// <summary>
            /// This service is for a CustomShape
            /// </summary>
            public const String DRAW_SHAPE_CUSTOM = "com.sun.star.drawing.CustomShape";

            /// <summary>
            /// This service is for a closed bezier shape.
            /// </summary>
            public const String DRAW_SHAPE_BEZIER_CLOSED = "com.sun.star.drawing.ClosedBezierShape";
            /// <summary>
            /// This service is for an open bezier shape.
            /// </summary>
            public const String DRAW_SHAPE_BEZIER_OPEN = "com.sun.star.drawing.OpenBezierShape";
            /// <summary>
            /// This service is for a polyline shape.
            /// </summary>
            public const String DRAW_SHAPE_POLYLINE = "com.sun.star.drawing.PolyLineShape";
            /// <summary>
            /// This service is for a polygon shape.
            /// </summary>
            public const String DRAW_SHAPE_POLYPOLYGON = "com.sun.star.drawing.PolyPolygonShape";

            /// <summary>
            /// This service describes a polypolygon.
            /// A polypolygon consists of multiple polygons combined in one.
            /// </summary>
            public const String DRAW_POLY_POLYGON_DESCRIPTOR = "com.sun.star.drawing.PolyPolygonDescriptor";
            /// <summary>
            /// This service describes a polypolygonbezier.
            /// A polypolygonbezier consists of multiple bezier polygons combined in one.
            /// </summary>
            public const String DRAW_POLY_POLYGON_BEZIER_DESCRIPTOR = "com.sun.star.drawing.PolyPolygonBezierDescriptor";

            /// <summary>
            /// represents the environment for a desktop component.
            /// Frames are the anchors for the office components and they are the components' link to the outside world. 
            /// They create a skeleton for the whole office api infrastructure by building frame hierarchys. These hierarchies contains all currently loaded documents and make it possible to walk during these trees.
            /// </summary>
            public const String FRAME_FRAME = "com.sun.star.frame.Frame";
            /// <summary>
            /// is the environment for components which can instantiate within frames.
            /// A desktop environment contains tasks with one or more frames in which components can be loaded. 
            /// The term "task" or naming a frame as a "task frame" is not in any way related to any additional implemented interfaces, it's just because these frames use task windows.
            /// </summary>
            public const String FRAME_DESKTOP = "com.sun.star.frame.Desktop";
            /// <summary>
            /// deprecated -- represents a top level frame in the frame hierarchy with the desktop as root.
            /// Please use the service Frame instead of this deprecated Task. If it's method XFrame.isTop() returns true, it's the same as a check for the Task service.
            /// </summary>
            public const String FRAME_TASK = "com.sun.star.frame.Task";

            /// <summary>
            /// This service specifies a collection of Bookmarks.
            /// </summary>
            public const String TEXT_BOOKMARK = "com.sun.star.text.Bookmark";
            /// <summary>
            /// is a table of text cells which is anchored to a surrounding text.
            /// Note: The anchor of the actual implementation for text tables does not have a position in the text.
            /// </summary>
            public const String TEXT_TEXT_TABLE = "com.sun.star.text.TextTable";
            /// <summary>
            /// specifies a rectangular shape which contains a Text object and is attached to a piece of surrounding Text.
            /// </summary>
            public const String TEXT_FRAME = "com.sun.star.text.TextFrame";

            /// <summary>
            /// The accessible view of a paragraph fragment.
            /// </summary>
            public const String TEXT_ACCESSIBLE_PARAGRAPH = "com.sun.star.text.AccessibleParagraphView";
            /// <summary>
            /// Every class has to support this service in order to be accessible.
            /// It provides the means to derive a XAccessibleContext object--which may but usually is not the same object as the object that 
            /// supports the XAccessible interface--that provides the actual information that is needed to make it accessible.
            /// Service Accessible is just a wrapper for the interface XAccessible. See the interface's documentation for more information.
            /// </summary>
            public const String ACCESSIBILITY_ACCESSIBLE = "com.sun.star.accessibility.Accessible";
            /// <summary>
            /// Central service of the Accessibility API that gives access to various facets of an object's content.
            /// This service has to be implemented by every class that represents the actual accessibility information of another UNO service. 
            /// It exposes two kinds of information: A tree structure in which all accessible objects are organized can be navigated in freely. 
            /// It typically represents spatial relationship of one object containing a set of children like a dialog box contains a set of buttons. 
            /// Additionally the XAccessibleContext interface of this service exposes methods that provide access to the actual object's content. 
            /// This can be the object's role, name, description, and so on.
            /// </summary>
            public const String ACCESSIBILITY_CONTEXT = "com.sun.star.accessibility.AccessibleContext";

            /// <summary>
            /// The AccessibleShape service is implemented by UNO shapes to provide accessibility information that describe the shapes' features. 
            /// A UNO shape is any object that implements the XShape interface.
            /// The content of a draw page is modeled as tree of accessible shapes and accessible text paragraphs. 
            /// The root of this (sub-)tree is the accessible draw document view. An accessible shape implements either this service or one of the 
            /// 'derived' service
            /// </summary>
            public const String DRAWING_ACCESSIBLE_SHAPE = "com.sun.star.drawing.AccessibleShape";
            /// <summary>
            /// The AccessibleDrawDocumentView service is implemented by views of Draw and Impress documents.
            /// An object that implements the AccessibleDrawDocumentView service provides information about the view of a 
            /// Draw or Impress document in one of the various view modes. With its children it gives access to the current 
            /// page and the shapes on that page.
            /// This service gives a simplified view on the underlying document. It tries both to keep the structure of 
            /// the accessibility representation tree as simple as possible and provide as much relevant information as possible. 
            /// </summary>
            public const String DRAWING_ACCESSIBLE_DOC = "com.sun.star.drawing.AccessibleDrawDocumentView";

            /// <summary>
            /// This abstract service specifies the general characteristics of an optional rotation and shearing for a Shape. 
            /// This service is deprecated, instead please use the Transformation property of the service Shape.
            /// </summary>
            public const String DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR = "com.sun.star.drawing.RotationDescriptor";
            /// <summary>
            /// This is a set of properties to describe the style for rendering an area.
            /// </summary>
            public const String DRAWING_PROPERTIES_FILL = "com.sun.star.drawing.FillProperties";
            /// <summary>
            /// The drawing properties custom
            /// </summary>
            public const String DRAWING_PROPERTIES_CUSTOM = "com.sun.star.drawing.CustomShapeProperties";
            /// <summary>
            /// This is a set of properties to describe the style for rendering a Line.
            /// The properties for line ends and line starts are only supported by shapes with open line ends.
            /// </summary>
            public const String DRAWING_PROPERTIES_LINE = "com.sun.star.drawing.LineProperties";
            /// <summary>
            /// This abstract service specifies the general characteristics of an optional text inside a Shape.
            /// </summary>
            public const String DRAWING_TEXT = "com.sun.star.drawing.Text";
            /// <summary>
            /// This is a set of properties to describe the style for rendering the text area inside a shape.
            /// </summary>
            public const String DRAWING_PROPERTIES_TEXT = "com.sun.star.drawing.TextProperties";
            /// <summary>
            /// describes the style of paragraphs.
            /// </summary>
            public const String DRAWING_PROPERTIES_PARAGRAPH = "com.sun.star.style.ParagraphProperties";
            /// <summary>
            /// contains settings for the style of paragraphs with complex text layout.
            /// </summary>
            public const String DRAWING_PROPERTIES_PARAGRAPH_COMPLEX = "com.sun.star.style.ParagraphPropertiesComplex";

            /// <summary>
            /// this service is supported from all shapes inside a PresentationDocument.
            /// This usually enhances objects of type ::com::sun::star::drawing::Shape with presentation properties.
            /// </summary>
            public const String PRESENTATION_SHAPE = "com.sun.star.presentation.Shape";

        }

        /// <summary>
        /// known doc types and components
        /// </summary>
        public static class DocTypes
        {
            public const String DOC_TYPE_BASE = "private:factory/";
            public const String DOC_TYPE_SDRAW_FULL = "private:factory/sdraw";
            public const String DOC_TYPE_SWRITER_FULL = "private:factory/swriter";
            public const String DOC_TYPE_SCALC_FULL = "private:factory/scalc";
            public const String DOC_TYPE_SIMPRESS_FULL = "private:factory/simpress";

            public const String COMP_DB_QUERY = ".component:DB/QueryDesign";
            public const String COMP_DB_TABLE = ".component:DB/TableDesign";
            public const String COMP_DB_RELATION = ".component:DB/RelationDesign";
            public const String COMP_DB_BROWSER = ".component:DB/DataSourceBrowser";
            public const String COMP_DB_FORM_GRID_VIEW = ".component:DB/FormGridView";
            public const String COMP_BIBLIOGRAPHY = " 	.component:Bibliography/View1";

            public const String SDRAW = "sdraw";
            public const String SWRITER = "swriter";
            public const String SCALC = "scalc";
            public const String SIMPRESS = "simpress";
            //TODO: extend that
        }

        /// <summary>
        /// All standard available tool bars
        /// </summary>
        public static class ToolBars
        {
            public const string Algignment = "private:resource/toolbar/alignmentbar";
            public const string ArrowShapes = "private:resource/toolbar/arrowshapes";
            public const string BasicShapes = "private:resource/toolbar/basicshapes";
            public const string CalloutShapes = "private:resource/toolbar/calloutshapes";
            public const string ColorBar = "private:resource/toolbar/colorbar";
            public const string DrawBar = "private:resource/toolbar/drawbar";
            public const string DrawObjectBar = "private:resource/toolbar/drawobjectbar";
            public const string ExtrusionObjectBar = "private:resource/toolbar/extrusionobjectbar";
            public const string FontWorkObjectBar = "private:resource/toolbar/fontworkobjectbar";
            public const string FontworkShapeTypes = "private:resource/toolbar/fontworkshapetypes";
            public const string FormatObjectBar = "private:resource/toolbar/formatobjectbar";
            public const string FormControlBar = "private:resource/toolbar/formcontrols";
            public const string FormdesignBar = "private:resource/toolbar/formdesign";
            public const string FormsFilterBar = "private:resource/toolbar/formsfilterbar";
            public const string FormsNavigationBar = "private:resource/toolbar/formsnavigationbar";
            public const string FormsObjectBar = "private:resource/toolbar/formsobjectbar";
            public const string FormtextBar = "private:resource/toolbar/formtextobjectbar";
            public const string FullScreenbar = "private:resource/toolbar/fullscreenbar";
            public const string GraphicObjectBar = "private:resource/toolbar/graphicobjectbar";
            public const string InsertBar = "private:resource/toolbar/insertbar";
            public const string InsertCellsBar = "private:resource/toolbar/insertcellsbar";
            public const string InsertObjectBar = "private:resource/toolbar/insertobjectbar";
            public const string MediaObjectBar = "private:resource/toolbar/mediaobjectbar";
            public const string MoreFormControlBar = "private:resource/toolbar/moreformcontrols";
            public const string PreviewBar = "private:resource/toolbar/previewbar";
            public const string StandardBar = "private:resource/toolbar/standardbar";
            public const string StarShapesBar = "private:resource/toolbar/starshapes";
            public const string SymbolShapesBar = "private:resource/toolbar/symbolshapes";
            public const string TextObjectBar = "private:resource/toolbar/textobjectbar";
            public const string ToolBar = "private:resource/toolbar/toolbar";
            public const string ViewerBar = "private:resource/toolbar/viewerbar";
            public const string MainMenu = "private:resource/menubar/menubar";
        }

        /// <summary>
        /// describes pre-defined possible control types to be used to display and enter property values within a ObjectInspector.
        /// </summary>
        public enum PropertyControlType : short
        {
            BUTTON = 0,
            /// <summary>
            /// denotes a control which allows the user to choose from a list of possible property values  
            /// </summary>
            LIST_BOX = 1,
            /// <summary>
            /// 	denotes a control which allows the user to choose from a list of possible property values, combined with the possibility to enter a new property value.  
            /// </summary>
            COMBO_BOX = 2,
            /// <summary>
            /// denotes a control which allows the user to enter property values consisting of a single line of text
            /// </summary>
            TEXT_FIELD = 3,
            /// <summary>
            /// denotes a control which allows the user to enter pure text, including line breaks 
            /// </summary>
            MULTI_LINE_TEXT_FIELD = 4,
            /// <summary>
            /// denotes a control which allows the user to enter a single character  
            /// </summary>
            CHARACTER_FIELD = 5,
            /// <summary>
            /// denotes a control which allows the user to enter a list of single-line strings  
            /// </summary>
            STRING_LIST_FIELD = 6,
            /// <summary>
            /// denotes a control which allows the user to choose from a list of colors.  
            /// </summary>
            COLOR_LIST_BOX = 7,
            /// <summary>
            /// denotes a control which allows the user to enter a numerical value  
            /// </summary>
            NUMERIC_FIELD = 8,
            /// <summary>
            /// denotes a control which allows the user to enter a date value  
            /// </summary>
            DATE_FIELD = 9,
            /// <summary>
            /// denotes a control which allows the user to enter a time value  
            /// </summary>
            TIME_FIELD = 10,
            /// <summary>
            /// denotes a control which allows the user to enter a combined date/time value  
            /// </summary>
            DATE_TIME_FIELD = 11,
            /// <summary>
            /// denotes a control which displays a string in a hyper-link-like appearance  
            /// </summary>
            HYPERLINK_FIELD = 12,
            /// <summary>
            /// denotes a non-standard property control, which is usually provided by an XPropertyHandler  
            /// </summary>
            UNKNOWN = 13
        }

        /// <summary>
        /// Element types received e.g. by XLayoutManager.getElement(...).Type
        /// </summary>
        public enum UiElementType : short
        {
            /// <summary>
            /// unknown user interface element type, which can be used 
            /// as a wild card to specify all types. 
            /// </summary>
            UNKNOWN = 0,

            /// <summary>
            /// specifies a menu bar.  
            /// </summary>
            MENUBAR = 1,

            /// <summary>
            /// specifies a popup menu. 
            /// </summary>
            POPUPMENU = 2,

            /// <summary>
            /// specifies a toolbar. 
            /// </summary>
            TOOLBAR = 3,

            /// <summary>
            /// specifies a statusbar. 
            /// </summary>
            STATUSBAR = 4,

            /// <summary>
            /// specifies a floating window, which can also be docked.
            /// </summary>
            FLOATINGWINDOW = 5,

            /// <summary>
            /// specifies a floating window, which can also be docked. 
            /// </summary>
            PROGRESSBAR = 6,

            /// <summary>
            /// specifies a tool panel 
            /// </summary>
            TOOLPANEL = 7,

            /// <summary>
            /// specifies a window that can be docked. 
            /// </summary>
            DOCKINGWINDOW = 7,

            /// <summary>
            /// specifies the number of constants.
            /// </summary>
            COUNT = 8
        }

        /// <summary>
        /// These values are used to specify the behavior of a Property.
        /// </summary>
        [Flags]
        public enum PropertyAttribute : short
        {
            UNKNOWN = 0,
            /// <summary>
            /// indicates that a property value can be void. It does not mean that the type of the property is void!
            /// </summary>
            MAYBEVOID = 1,
            /// <summary>
            /// indicates that a PropertyChangeEvent will be fired to all registered XPropertyChangeListeners whenever the value of this property changes.
            /// </summary>
            BOUND = 2,
            /// <summary>
            /// indicates that a PropertyChangeEvent will be fired to all registered XVetoableChangeListeners whenever the value of this property changes. This always implies that the property is bound, too.
            /// </summary>
            CONSTRAINED = 4,
            /// <summary>
            /// indicates that the value of the property is not persistent.
            /// </summary>
            TRANSIENT = 8,
            /// <summary>
            /// indicates that the value of the property is read-only.
            /// </summary>
            READONLY = 16,
            /// <summary>
            /// indicates that the value of the property can be ambiguous.
            /// </summary>
            MAYBEAMBIGUOUS = 32,
            /// <summary>
            /// indicates that the property can be set to default.
            /// </summary>
            MAYBEDEFAULT = 64,
            /// <summary>
            /// indicates that the property can be removed (i.e., by calling XPropertyContainer::removeProperty).
            /// </summary>
            REMOVEABLE = 128,
            /// <summary>
            /// indicates that a property is optional. This attribute is not of interest for concrete property implementations. It's needed for property specifications inside service specifications in UNOIDL.
            /// </summary>
            OPTIONAL = 256
        }

        /// <summary>
        /// Gets the orientation of a page
        /// </summary>
        public enum PaperOrientation
        {
            /// <summary>
            /// Portrait mode
            /// </summary>
            PORTRAIT = 0,
            /// <summary>
            /// landscape mode
            /// </summary>
            LANDSCAPE = 1,
        }


        #endregion

    }
}