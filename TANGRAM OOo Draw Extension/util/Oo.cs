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
using tud.mci.tangram.models;
using uno.util;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using Exception = System.Exception;
using System.Diagnostics;
using System.Threading;

namespace tud.mci.tangram.util
{
    public static class OO
    {
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
                        success = TimeLimitExecutor.WaitForExecuteWithTimeLimit(5000,
                             () =>
                             {
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
                             },
                             "Bootstrap-ContextLoader"
                             );

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
        /// <returns></returns>
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

                _xMsf = xMcf.createInstanceWithContext(Services.MULTI_SERVICE_FACTORY, xCompContext) as XMultiServiceFactory;

                String[] bla = xMcf.getAvailableServiceNames();
                System.Diagnostics.Debug.WriteLine("Services:\n" + String.Join("\n", bla));

                if (_xMsf == null)
                {
                    xCompContext = GetContext(true);
                    xMcf = GetMultiComponentFactory(xCompContext, true);

                    _xMsf = xMcf.createInstanceWithContext(Services.MULTI_SERVICE_FACTORY, xCompContext) as XMultiServiceFactory;

                }
                addListener(_xMsf);

            }
            return _xMsf;
        }

        private static readonly Object conCheckLock = new Object();
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
                catch (System.Exception ex)
                {
                    reset();
                }

                return false;
            }
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
                catch (System.Exception ex)
                {
                }
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
            catch (System.Exception ex)
            { }
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

            public const String MULTI_SERVICE_FACTORY = "com.sun.star.lang.MultiServiceFactory";

            public const String DOCUMENT = "com.sun.star.document.OfficeDocument";
            public const String DOCUMENT_TEXT = "com.sun.star.text.TextDocument";
            public const String DOCUMENT_WEB = "com.sun.star.text.WebDocument";
            public const String DOCUMENT_DRAWING_GENERIC = "com.sun.star.drawing.GenericDrawingDocumen";
            public const String DOCUMENT_DRAWING_DOCUMENT_FACTORY = "com.sun.star.drawing.DrawingDocumentFactory";
            public const String DOCUMENT_DRAWING = "com.sun.star.drawing.DrawingDocument";
            public const String DOCUMENT_SPREADSHEET = "com.sun.star.sheet.SpreadsheetDocument";
            public const String DOCUMENT_FILTER_FACTORY = "com.sun.star.document.FilterFactory";

            public const String AWT_TOOLKIT = "com.sun.star.awt.Toolkit";

            public const String AWT_CONTROL_BUTTON_MODEL = "com.sun.star.awt.UnoControlButtonModel";
            public const String AWT_CONTROL_CHECKBOX_MODEL = "com.sun.star.awt.UnoControlCheckBoxModel";
            public const String AWT_CONTROL_COMBOBOX_MODEL = "com.sun.star.awt.UnoControlComboBoxModel";
            public const String AWT_CONTROL_CURRENCY_FIELD_MODEL = "com.sun.star.awt.UnoControlCurrencyFieldModel";
            public const String AWT_CONTROL_DATE_FIELD_MODEL = "com.sun.star.awt.UnoControlDateFieldModel";
            public const String AWT_CONTROL_DIALOG = "com.sun.star.awt.UnoControlDialog";
            public const String AWT_CONTROL_DIALOG_MODEL = "com.sun.star.awt.UnoControlDialogModel";
            public const String AWT_CONTROL_EDIT_MODEL = "com.sun.star.awt.UnoControlEditModel";
            public const String AWT_CONTROL_FILE_MODEL = "com.sun.star.awt.UnoControlFileControlModel";
            public const String AWT_CONTROL_FIXED_LINE_MODEL = "com.sun.star.awt.UnoControlFixedLineModel";
            public const String AWT_CONTROL_FIXED_LINE = "com.sun.star.awt.UnoControlFixedLine";
            public const String AWT_CONTROL_FORMATTED_FIELD_MODEL = "com.sun.star.awt.UnoControlFormattedFieldModel";
            public const String AWT_CONTROL_GRID_MODEL = "com.sun.star.awt.UnoControlGridModel";
            public const String AWT_CONTROL_GROUP_BOX_MODEL = "com.sun.star.awt.UnoControlGroupBoxModel";
            public const String AWT_CONTROL_IMAGE_MODEL = "com.sun.star.awt.UnoControlImageControlModel";
            public const String AWT_CONTROL_LISTBOX_MODEL = "com.sun.star.awt.UnoControlListBoxModel";
            public const String AWT_CONTROL_NUMMERIC_FIELD_MODEL = "com.sun.star.awt.UnoControlNumericFieldModel";
            public const String AWT_CONTROL_PATTERNFIELD_MODEL = "com.sun.star.awt.UnoControlPatternFieldModel";
            public const String AWT_CONTROL_PROGRESS_BAR_MODEL = "com.sun.star.awt.UnoControlProgressBarModel";
            public const String AWT_CONTROL_PROGRESS_BAR = "com.sun.star.awt.service UnoControlProgressBar";
            public const String AWT_CONTROL_RADIOBUTTON_MODEL = "com.sun.star.awt.UnoControlRadioButtonModel";
            public const String AWT_CONTROL_ROADMAP_MODEL = "com.sun.star.awt.UnoControlRoadmapModel";
            public const String AWT_CONTROL_SCROLLBAR_MODEL = "com.sun.star.awt.UnoControlScrollBarModel";
            public const String AWT_CONTROL_TABPAGE_CONTAINER_MODEL = "com.sun.star.awt.UnoControlTabPageContainerModel";
            public const String AWT_CONTROL_TABPAGE_MODEL = "com.sun.star.awt.UnoControlTabPageModel";
            public const String AWT_CONTROL_TEXT_FIXED_MODEL = "com.sun.star.awt.UnoControlFixedTextModel";
            public const String AWT_CONTROL_TEXT_FIXED = "com.sun.star.awt.UnoControlFixedText";
            public const String AWT_CONTROL_TIME_FIELD_MODEL = "com.sun.star.awt.UnoControlTimeFieldModel";
            public const String AWT_CONTROL_TREE_MODEL = "com.sun.star.awt.TreeControlModel";
            public const String AWT_CONTROL_MUTABLETREE_MODEL = "com.sun.star.awt.tree.MutableTreeDataModel";

            public const String GRAPHIC_GRAPHICPROVIDER = "com.sun.star.graphic.GraphicProvider";

            public const String AWT_ACCESSIBILITY_MENU_SEPERATOR = "com.sun.star.awt.AccessibleMenuSeparator";
            public const String AWT_ACCESSIBILITY_MENU = "com.sun.star.awt.AccessibleMenu";
            public const String AWT_ACCESSIBILITY_MENUBAR = "com.sun.star.awt.AccessibleMenuBar";
            public const String AWT_ACCESSIBILITY_MENU_ITEM = "com.sun.star.awt.AccessibleMenuItem";
            public const String AWT_ACCESSIBILITY_WINDOW = "com.sun.star.awt.AccessibleWindow";
            public const String AWT_ACCESSIBILITY_POPUP_MENU = "com.sun.star.awt.AccessiblePopupMenu";
            public const String AWT_ACCESSIBILITY_FIXEDTEXT = "com.sun.star.awt.AccessibleFixedText";
            public const String AWT_ACCESSIBILITY_EDIT = "com.sun.star.awt.AccessibleEdit";
            public const String AWT_ACCESSIBILITY_BUTTON = "com.sun.star.awt.AccessibleButton";
            public const String AWT_ACCESSIBILITY_SCROLLBAR = "com.sun.star.awt.AccessibleScrollBar";
            public const String AWT_ACCESSIBILITY_TAB_PAGE = "com.sun.star.awt.AccessibleTabPage";
            public const String AWT_ACCESSIBILITY_TAB_CONTROL = "com.sun.star.awt.AccessibleTabControl";

            public const String AWT_POINTER = "com.sun.star.awt.Pointer";

            public const String UTIL_FORMAT_NUMBER = "com.sun.star.util.NumberFormatsSupplier";
            public const String UTIL_PATH_SETTINGS = "com.sun.star.util.PathSettings";
            public const String UTIL_FILE_CONTENT_PRVIDER = "com.sun.star.ucb.FileContentProvider";
            public const String UTIL_URL_TRANSFORMER = "com.sun.star.util.URLTransformer";

            public const String UI_MOD_UI_CONF_MGR_SUPPLIER = "com.sun.star.ui.ModuleUIConfigurationManagerSupplier";

            public const String DRAW_SHAPE = "com.sun.star.drawing.Shape";
            public const String DRAW_SHAPE_RECT = "com.sun.star.drawing.RectangleShape";
            public const String DRAW_SHAPE_ELLIPSE = "com.sun.star.drawing.EllipseShape";
            public const String DRAW_SHAPE_LINE = "com.sun.star.drawing.LineShape";
            public const String DRAW_SHAPE_TEXT = "com.sun.star.drawing.TextShape";
            public const String DRAW_SHAPE_CUSTOM = "com.sun.star.drawing.CustomShape";


            public const String FRAME_FRAME = "com.sun.star.frame.Frame";
            public const String FRAME_DESKTOP = "com.sun.star.frame.Desktop";
            public const String FRAME_TASK = "com.sun.star.frame.Task";

            public const String TEXT_BOOKMARK = "com.sun.star.text.Bookmark";
            public const String TEXT_TEXT_TABLE = "com.sun.star.text.TextTable";
            public const String TEXT_FRAME = "com.sun.star.text.TextFrame";

            public const String TEXT_ACCESSIBLE_PARAGRAPH = "com.sun.star.text.AccessibleParagraphView";

            public const String ACCESSIBILITY_ACCESSIBLE = "com.sun.star.accessibility.Accessible";
            public const String ACCESSIBILITY_CONTEXT = "com.sun.star.accessibility.AccessibleContext";

            public const String DRAWING_ACCESSIBLE_SHAPE = "com.sun.star.drawing.AccessibleShape";
            public const String DRAWING_ACCESSIBLE_DOC = "com.sun.star.drawing.AccessibleDrawDocumentView";

            public const String DRAWING_PROPERTIES_ROTATE_AND_SHERE_DESCRIPTOR = "com.sun.star.drawing.RotationDescriptor";
            public const String DRAWING_PROPERTIES_FILL = "com.sun.star.drawing.FillProperties";
            public const String DRAWING_PROPERTIES_CUSTOM = "com.sun.star.drawing.CustomShapeProperties";
            public const String DRAWING_PROPERTIES_LINE = "com.sun.star.drawing.LineProperties";
            public const String DRAWING_TEXT = "com.sun.star.drawing.Text";
            public const String DRAWING_PROPERTIES_TEXT = "com.sun.star.drawing.TextProperties";
            public const String DRAWING_PROPERTIES_PARAGRAPH = "com.sun.star.style.ParagraphProperties";
            public const String DRAWING_PROPERTIES_PARAGRAPH_COMPLEX = "com.sun.star.style.ParagraphPropertiesComplex";

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

        public enum PropertyControlType : short
        {
            BUTTON = 0,
            LIST_BOX = 1,
            COMBO_BOX = 2,
            TEXT_FIELD = 3,
            MULTI_LINE_TEXT_FIELD = 4,
            CHARACTER_FIELD = 5,
            STRING_LIST_FIELD = 6,
            COLOR_LIST_BOX = 7,
            NUMERIC_FIELD = 8,
            DATE_FIELD = 9,
            TIME_FIELD = 10,
            DATE_TIME_FIELD = 11,
            HYPERLINK_FIELD = 12,
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

        #endregion

    }

}