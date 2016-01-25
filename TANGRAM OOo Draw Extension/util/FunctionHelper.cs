// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extension
// Author           : Admin
// Created          : 09-19-2012
//
// Last Modified By : Admin
// Last Modified On : 09-19-2012
// ***********************************************************************
// <copyright file="FunctionHelper.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.container;
using tud.mci.tangram.models;

namespace tud.mci.tangram.util
{
    /**************************************************************
     * 
     * Licensed to the Apache Software Foundation (ASF) under one
     * or more contributor license agreements.  See the NOTICE file
     * distributed with this work for additional information
     * regarding copyright ownership.  The ASF licenses this file
     * to you under the Apache License, Version 2.0 (the
     * "License"); you may not use this file except in compliance
     * with the License.  You may obtain a copy of the License at
     * 
     *   http://www.apache.org/licenses/LICENSE-2.0
     * 
     * Unless required by applicable law or agreed to in writing,
     * software distributed under the License is distributed on an
     * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
     * KIND, either express or implied.  See the License for the
     * specific language governing permissions and limitations
     * under the License.
     * 
     *************************************************************/

    /// <summary>
    /// Is a collection of basic features.
    /// This helper shows different functionality of framework api
    /// in an example manner. You can use the follow ones:
    ///      (1) parse URL's
    ///      (2) create frames (inside/outside a java application)
    ///      (3) dispatches (with[out] notifications)
    ///      (4) loading/saving documents
    ///      (5) convert documents to HTML (if possible)
    ///      (6) close documents (and her frames) correctly
    ///
    /// There exist some other helper functionality too, which
    /// doesn't use or demonstrate the office api:
    ///      (a) getting file names by using a file chooser
    /// </summary>
    public static class FunctionHelper
    {
        /// <summary>
        /// This convert an URL (formated as a string) to a struct com.sun.star.util.URL.
        /// It use a special service to do that: the URLTransformer.
        /// Because some API calls need it and it's not allowed to set "Complete"
        /// part of the util struct only. The URL must be parsed.
        /// </summary>
        /// <param name="sUrl">URL for parsing in string notation</param>
        /// <returns>URL in UNO struct notation </returns>
        public static URL ParseUrl(String sUrl)
        {
            URL aUrl;

            if (sUrl == null || sUrl.Equals("", StringComparison.CurrentCulture))
            {
                System.Diagnostics.Debug.WriteLine("wrong using of URL parser");
                return null;
            }

            try
            {
                XComponentContext xOfficeCtx = OO.GetContext();

                // Create special service for parsing of given URL.
                var xParser = (XURLTransformer)xOfficeCtx.getServiceManager().createInstanceWithContext(
                            OO.Services.UTIL_URL_TRANSFORMER, xOfficeCtx);

                // Because it's an in/out parameter we must use an array of URL objects.
                var aParseUrl = new URL[1];
                aParseUrl[0] = new URL {Complete = sUrl};

                // Parse the URL
                xParser.parseStrict(ref aParseUrl[0]);

                aUrl = aParseUrl[0];
            }
            catch (RuntimeException)
            {
                // Any UNO method of this scope can throw this exception.
                // Reset the return value only.
                aUrl = null;
            }
            catch (unoidl.com.sun.star.uno.Exception)
            {
                // "createInstance()" method of used service manager can throw it.
                // Then it wasn't possible to get the URL transformer.
                // Return default instead of realy parsed URL.
                aUrl = null;
            }

            return aUrl;
        }


        /// <summary>
        /// Dispatch an URL to given frame.
        /// Caller can register himself for following status events for dispatched
        /// URL too. But nobody guarantee that such notifications will occur.
        /// (see dispatchWithNotification() if you interest on that)
        /// The returned dispatch object should be hold alive by caller
        /// till he doesn't need it any longer. Otherwise the dispatcher can(!)
        /// die by decreasing his refcount.
        /// (Note: Deregistration is part of this listener himself!)
        /// </summary>
        /// <param name="xFrame">frame wich should be the target of this dispatch</param>
        /// <param name="aUrl">full parsed and converted office URL for dispatch</param>
        /// <param name="lProperties">optional arguments for dispatch</param>
        /// <param name="xListener">optional listener which is registered automaticly for status events</param>
        /// <returns>
        /// It's the used dispatch object and can be used for deregistration of an optional listener.
        /// Otherwise caller can ignore it.
        /// </returns>
        public static XDispatch Execute(XFrame xFrame, URL aUrl, PropertyValue[] lProperties, XStatusListener xListener)
        {
            XDispatch xDispatcher;

            try
            {
                // Query the frame for right interface which provides access to all available dispatch objects.
                XDispatchProvider xProvider = (XDispatchProvider)xFrame;

                // Ask him for right dispatch object for given URL.
                // Force given frame as target for following dispatch by using "".
                // It means the same like "_self".
                xDispatcher = xProvider.queryDispatch(aUrl, "", 0);

                // Dispatch the URL into the frame.
                if (xDispatcher != null)
                {
                    if (xListener != null)
                        xDispatcher.addStatusListener(xListener, aUrl);

                    xDispatcher.dispatch(aUrl, lProperties);
                }
            }
            catch (RuntimeException exUno)
            {
                // Any UNO method of this scope can throw this exception.
                // But there will be nothing to do then - because
                // we haven't changed anything inside the remote objects
                // except method "addStatusListener().
                // But in this case the source of this exception has to
                // rollback all his operations. There is no chance to
                // make anything right then.
                // Reset the return value to a default - that's it.
                System.Diagnostics.Debug.WriteLine(exUno);
                xDispatcher = null;
            }

            return xDispatcher;
        }

        /// <summary>
        /// Dispatch an URL to given frame.
        /// Caller can register himself for following result events for dispatched
        /// URL too. Notifications are guaranteed (instead of dispatch())
        /// Returning of the dispatch object isn't necessary.
        /// Nobody must hold it alive longer the dispatch needs.
        /// </summary>
        /// <param name="xFrame">frame wich should be the target of this dispatch</param>
        /// <param name="aUrl">full parsed and converted office URL for dispatch</param>
        /// <param name="lProperties">optional arguments for dispatch</param>
        /// <param name="xListener">
        /// optional listener which is registered automatically for status events
        /// (Note: Deregistration is not supported. Dispatcher does it automatically.)
        /// </param>
        public static void ExecuteWithNotification(XFrame xFrame, URL aUrl, PropertyValue[] lProperties, XDispatchResultListener xListener)
        {
            try
            {
                // Query the frame for right interface which provides access to all available dispatch objects.
                XDispatchProvider xProvider = (XDispatchProvider)xFrame;

                // Ask him for right dispatch object for given URL.
                // Force THIS frame as target for following dispatch.
                // Attention: The interface XNotifyingDispatch is an optional one!
                XDispatch xDispatcher = xProvider.queryDispatch(aUrl, "", 0);
                XNotifyingDispatch xNotifyingDispatcher = (XNotifyingDispatch)xDispatcher;

                // Dispatch the URL.
                if (xNotifyingDispatcher != null)
                    xNotifyingDispatcher.dispatchWithNotification(aUrl, lProperties, xListener);
            }
            catch (RuntimeException exUno)
            {
                // Any UNO method of this scope can throw this exception.
                // But there is nothing we can do then.
                System.Diagnostics.Debug.WriteLine(exUno);
            }
        }

        /// <summary>
        /// Load document specified by an URL into given frame synchronously.
        /// The result of this operation will be the loaded document for success
        /// or null if loading failed.
        /// </summary>
        /// <param name="xFrame">frame wich should be the target of this load call</param>
        /// <param name="sUrl">unparsed URL for loading</param>
        /// <param name="lProperties">optional arguments</param>
        /// <returns>
        /// the loaded document for success or null if it's failed
        /// </returns>
        public static XComponent LoadDocument(XFrame xFrame, String sUrl, PropertyValue[] lProperties)
        {
            XComponent xDocument;
            String sOldName = null;

            try
            {
                XComponentContext xCtx = OO.GetContext();

                // First prepare frame for loading
                // We must address it inside the frame tree without any complications.
                // So we set an unambiguous (we hope it) name and use it later.
                // Don't forget to reset original name after that.
                sOldName = xFrame.getName();
                String sTarget = "odk_officedev_desk";
                xFrame.setName(sTarget);

                // Get access to the global component loader of the office
                // for synchronous loading the document.
                var xLoader =
                    (XComponentLoader)xCtx.getServiceManager().createInstanceWithContext(OO.Services.FRAME_DESKTOP, xCtx);

                // Load the document into the target frame by using his name and
                // special search flags.
                xDocument = xLoader.loadComponentFromURL(
                    sUrl,
                    sTarget,
                    FrameSearchFlag.CHILDREN,
                    lProperties);

                // don't forget to restore old frame name ...
                xFrame.setName(sOldName);
            }
            catch (unoidl.com.sun.star.io.IOException exIo)
            {
                // Can be thrown by "loadComponentFromURL()" call.
                // The only thing we should do then is to reset changed frame name!
                System.Diagnostics.Debug.WriteLine(exIo);
                xDocument = null;
                if (sOldName != null)
                    xFrame.setName(sOldName);
            }
            catch (IllegalArgumentException exIllegal)
            {
                // Can be thrown by "loadComponentFromURL()" call.
                // The only thing we should do then is to reset changed frame name!
                System.Diagnostics.Debug.WriteLine(exIllegal);
                xDocument = null;
                if (sOldName != null)
                    xFrame.setName(sOldName);
            }
            catch (RuntimeException exRuntime)
            {
                // Any UNO method of this scope can throw this exception.
                // The only thing we can try(!) is to reset changed frame name.
                System.Diagnostics.Debug.WriteLine(exRuntime);
                xDocument = null;
                if (sOldName != null)
                    xFrame.setName(sOldName);
            }
            catch (unoidl.com.sun.star.uno.Exception exUno)
            {
                // "createInstance()" method of used service manager can throw it.
                // The only thing we should do then is to reset changed frame name!
                System.Diagnostics.Debug.WriteLine(exUno);
                xDocument = null;
                if (sOldName != null)
                    xFrame.setName(sOldName);
            }

            return xDocument;
        }

        /// <summary>
        /// Save currently loaded document of given frame.
        /// </summary>
        /// <param name="xDocument">document for saving changes</param>
        public static void SaveDocument(XComponent xDocument)
        {
            try
            {
                // Check for supported model functionality.
                // Normally the application documents (text, spreadsheet ...) do so
                // but some other ones (e.g. db components) doesn't do that.
                // They can't be save then.
                var xModel = (XModel)xDocument;
                if (xModel != null)
                {
                    // Check for modifications => break save process if there is nothing to do.
                    XModifiable xModified = (XModifiable)xModel;
                    if (xModified.isModified() == true)
                    {
                        XStorable xStore = (XStorable)xModel;
                        xStore.store();
                    }
                }
            }
            catch (unoidl.com.sun.star.io.IOException exIo)
            {
                // Can be thrown by "store()" call.
                // But there is nothing we can do then.
                System.Diagnostics.Debug.WriteLine(exIo);
            }
            catch (RuntimeException exUno)
            {
                // Any UNO method of this scope can throw this exception.
                // But there is nothing we can do then.
                System.Diagnostics.Debug.WriteLine(exUno);
            }
        }

        /// <summary>
        /// It try to export given document in HTML format.
        /// Current document will be converted to HTML and moved to new place on disk.
        /// A "new" file will be created by given URL (may be overwritten
        /// if it already exist). Right filter will be used automatically if factory of
        /// this document support it. If no valid filter can be found for export,
        /// nothing will be done here.
        /// </summary>
        /// <param name="xDocument">document which should be exported</param>
        /// <param name="sUrl">target URL for converted document</param>
        public static void SaveAsHtml(XComponent xDocument, String sUrl)
        {
            try
            {
                // First detect factory of this document.
                // Ask for the supported service name of this document.
                // If information is available it can be used to find out which
                // filter exist for HTML export. Normally this filter should be searched
                // inside the filter configuration but this little demo doesn't do so.
                // (see service com.sun.star.document.FilterFactory for further
                // informations too)
                // Well known filter names are used directly. They must exist in current
                // office installation. Otherwise this code will fail. But to prevent
                // this code against missing filters it check for existing state of it.
                XServiceInfo xInfo = (XServiceInfo)xDocument;

                if (xInfo != null)
                {
                    // Find out possible filter name.
                    String sFilter = null;
                    if (xInfo.supportsService(OO.Services.DOCUMENT_TEXT))
                        sFilter = "HTML (StarWriter)";
                    else
                        if (xInfo.supportsService(OO.Services.DOCUMENT_WEB))
                            sFilter = "HTML";
                        else
                            if (xInfo.supportsService(OO.Services.DOCUMENT_SPREADSHEET))
                                sFilter = "HTML (StarCalc)";

                    // Check for existing state of this filter.
                    if (sFilter != null)
                    {
                        XComponentContext xCtx = OO.GetContext();

                        var xFilterContainer = (XNameAccess)xCtx.getServiceManager().createInstanceWithContext(
                                    OO.Services.DOCUMENT_FILTER_FACTORY, xCtx);

                        if (xFilterContainer.hasByName(sFilter) == false)
                            sFilter = null;
                    }

                    // Use this filter for export.
                    if (sFilter != null)
                    {
                        // Export can be forced by saving the document and using a
                        // special filter name which can write needed format. Build
                        // necessary argument list now.
                        // Use special flag "Overwrite" too, to prevent operation
                        // against possible exceptions, if file already exist.
                        var lProperties = new PropertyValue[2];
                        lProperties[0] = new PropertyValue {Name = "FilterName", Value = Any.Get(sFilter)};
                        lProperties[1] = new PropertyValue {Name = "Overwrite", Value = Any.Get(true)};

                        XStorable xStore = (XStorable)xDocument;

                        xStore.storeAsURL(sUrl, lProperties);
                    }
                }
            }
            catch (unoidl.com.sun.star.io.IOException exIo)
            {
                // Can be thrown by "store()" call.
                // Do nothing then. Saving failed - that's it.
                System.Diagnostics.Debug.WriteLine(exIo);
            }
            catch (RuntimeException exRuntime)
            {
                // Can be thrown by any uno call.
                // Do nothing here. Saving failed - that's it.
                System.Diagnostics.Debug.WriteLine(exRuntime);
            }
            catch (unoidl.com.sun.star.uno.Exception exUno)
            {
                // Can be thrown by "createInstance()" call of service manager.
                // Do nothing here. Saving failed - that's it.
                System.Diagnostics.Debug.WriteLine(exUno);
            }
        }


        /// <summary>
        /// Try to close the document without any saving of modifications.
        /// We can try it only! Controller and/or model of this document
        /// can disagree with that. But mostly they doesn't do so.
        /// </summary>
        /// <param name="xDocument">document which should be clcosed</param>
        public static void CloseDocument(XComponent xDocument)
        {
            try
            {
                // Check supported functionality of the document (model or controller).
                XModel xModel = (XModel)xDocument;

                if (xModel != null)
                {
                    // It's a full featured office document.
                    // Reset the modify state of it and close it.
                    // Note: Model can disagree by throwing a veto exception.
                    XModifiable xModify = (XModifiable)xModel;

                    xModify.setModified(false);
                    xDocument.dispose();
                }
                else
                {
                    // It's a document which supports a controller .. or may by a pure
                    // window only. If it's at least a controller - we can try to
                    // suspend him. But - he can disagree with that!
                    XController xController = (XController)xDocument;

                    if (xController != null)
                    {
                        if (xController.suspend(true))
                        {
                            // Note: Don't dispose the controller - destroy the frame
                            // to make it right!
                            XFrame xFrame = xController.getFrame();
                            xFrame.dispose();
                        }
                    }
                }
            }
            catch (PropertyVetoException exVeto)
            {
                // Can be thrown by "setModified()" call on model.
                // He disagree with our request.
                // But there is nothing to do then. Following "dispose()" call wasn't
                // never called (because we catch it before). Closing failed -that's it.
                System.Diagnostics.Debug.WriteLine(exVeto);
            }
            catch (DisposedException exDisposed)
            {
                // If an UNO object was already disposed before - he throw this special
                // runtime exception. Of course every UNO call must be look for that -
                // but it's a question of error handling.
                // For demonstration this exception is handled here.
                System.Diagnostics.Debug.WriteLine(exDisposed);
            }
            catch (RuntimeException exRuntime)
            {
                // Every uno call can throw that.
                // Do nothing - closing failed - that's it.
                System.Diagnostics.Debug.WriteLine(exRuntime);
            }
        }

        /// <summary>
        /// Try to close the frame instead of the document.
        /// It shows the possible interface to do so.
        /// </summary>
        /// <param name="xFrame">frame which should be clcosed</param>
        /// <returns>    
        /// <TRUE/>
        ///  in case frame could be closed
        /// <FALSE/>
        ///  otherwise
        /// </returns>
        public static bool CloseFrame(XFrame xFrame)
        {
            bool bClosed = false;

            try
            {
                // first try the new way: use new interface XCloseable
                // It replace the deprecated XTask::close() and should be preferred ...
                // if it can be queried.
                XCloseable xCloseable = (XCloseable)xFrame;
                if (xCloseable != null)
                {
                    // We deliver the owner ship of this frame not to the (possible)
                    // source which throw a CloseVetoException. We whish to have it
                    // under our own control.
                    try
                    {
                        xCloseable.close(false);
                        bClosed = true;
                    }
                    catch (CloseVetoException)
                    {
                        bClosed = false;
                    }
                }
                else
                {
                    // OK: the new way isn't possible. Try the old one.
                    XTask xTask = (XTask)xFrame;
                    if (xTask != null)
                    {
                        // return value doesn't interest here. Because
                        // we forget this task ...
                        bClosed = xTask.close();
                    }
                }
            }
            catch (DisposedException)
            {
                // Of course - this task can be already dead - means disposed.
                // But for us it's not important. Because we tried to close it too.
                // And "already disposed" or "closed" should be the same ...
                bClosed = true;
            }

            return bClosed;
        }


        private const String BASEFRAMENAME = "Desk View ";
        /// <summary>
        /// Try to find an unique frame name, which isn't currently used inside
        /// remote office instance. Because we create top level frames
        /// only, it's enough to check the names of existing child frames on the
        /// desktop only.
        /// </summary>
        /// <returns>
        /// should represent an unique frame name, which currently isn't
        /// used inside the remote office frame tree
        /// (Couldn't guaranteed for a real multi threaded environment.
        /// But we try it ...)
        /// </returns>
        public static String GetUniqueFrameName()
        {
            String sName = null;

            XComponentContext xCtx = OO.GetContext();

            try
            {
                var xSupplier = (XFramesSupplier)xCtx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", xCtx);

                var xContainer = (XIndexAccess)xSupplier.getFrames();

                int nCount = xContainer.getCount();
                for (int i = 0; i < nCount; ++i)
                {
                    XFrame xFrame = (XFrame)xContainer.getByIndex(i).Value;
                    sName = BASEFRAMENAME + _mnViewCount;
                    while (String.Compare(sName, xFrame.getName(), System.StringComparison.Ordinal) == 0)
                    {
                        ++_mnViewCount;
                        sName = BASEFRAMENAME + _mnViewCount;
                    }
                }
            }
            catch (unoidl.com.sun.star.uno.Exception)
            {
                sName = BASEFRAMENAME;
            }

            if (sName == null)
            {
                System.Diagnostics.Debug.WriteLine("invalid name!");
                sName = BASEFRAMENAME;
            }

            return sName;
        }


        /**
         * mnViewCount we try to set unique names on every frame we create (that's why we must count it)
         */
        private static int _mnViewCount/* = 0*/;
    }
}
