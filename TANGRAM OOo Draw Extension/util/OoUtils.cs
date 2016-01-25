using System;
using System.Collections.Generic;
using unoidl.com.sun.star.document;
using unoidl.com.sun.star.awt;
using System.Diagnostics;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.container;
using tud.mci.tangram.models;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.graphic;

namespace tud.mci.tangram.util
{
    /*
     * OOoUtils.java
     *
     * Created on February 22, 2003, 12:10 PM
     *
     * Copyright 2003 Danny Brewer
     * Copyright 2014 Jens Bornschein
     * Anyone may run this code.
     * If you wish to modify or distribute this code, then
     *  you are granted a license to do so only under the terms
     *  of the Gnu Lesser General Public License.
     * See:  http://www.gnu.org/licenses/lgpl.html
     */

    /**
     *
     * @author  danny brewer and Jens Bornschein
     */
    public class OoUtils
    {
        //----------------------------------------------------------------------
        //  Conveniences to make some structures
        //----------------------------------------------------------------------

        /// <summary>
        /// Create a basic property value object.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="value">The value.</param>
        /// <returns>the PropertyValue element</returns>
        public static PropertyValue MakePropertyValue(String propertyName, Object value)
        {
            var propValue = new PropertyValue { Name = propertyName, Value = Any.Get(value) };
            return propValue;
        }

        /// <summary>
        /// Creates a basic unoidl.com.sun.star.awt.Point Object.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <returns>a basic unoidl.com.sun.star.awt.Point Object</returns>
        public static Point MakePoint(int x, int y)
        {
            var point = new Point { X = x, Y = y };
            return point;
        }

        /// <summary>
        /// Creates a basic unoidl.com.sun.star.awt.Size Object.
        /// </summary>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns>a basic unoidl.com.sun.star.awt.Size Object</returns>
        public static Size MakeSize(int width, int height)
        {
            var size = new Size { Width = width, Height = height };
            return size;
        }

        /// <summary>
        /// Creates a basic unoidl.com.sun.star.awt.Rectangle Object.
        /// </summary>
        /// <param name="x">The x.</param>
        /// <param name="y">The y.</param>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns>a basic unoidl.com.sun.star.awt.Rectangle Object</returns>
        public static Rectangle MakeRectangle(int x, int y, int width, int height)
        {
            var rect = new Rectangle { X = x, Y = y, Width = width, Height = height };
            return rect;
        }

        #region Property

        #region SET

        #region Normal

        /// <summary>
        /// Gets a String property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value.</param>
        /// <returns><c>true</c> if the property could been set without an error, otherwise <c>false</c></returns>
        public static bool SetProperty(Object obj, String propName, Object value)
        {
            if (obj is XControl)
            {
                obj = ((XControl)obj).getModel();
            }

            // We need the XPropertySet interface.
            XPropertySet objXPropertySet;
            if (obj is XPropertySet)
            {
                // If the right interface was passed in, just typecaset it.
                objXPropertySet = (XPropertySet)obj;
            }
            else
            {
                // Get a different interface to the drawDoc.
                // The parameter passed in to us is the wrong interface to the object.
                objXPropertySet = obj as XPropertySet;
            }

            // Now just call our sibling using the correct interface.
            return SetProperty(objXPropertySet, propName, value);
        }

        /// <summary>
        /// Sets a property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <returns><c>true</c> if the property could been set without an error, otherwise <c>false</c></returns>
        public static bool SetProperty(XPropertySet obj, String propName, Object value)
        {
            if (obj != null)
            {
                //var proInfo = obj.getPropertySetInfo();
                //var properties = proInfo.getProperties();
                //List<String> props = new List<String>();
                //foreach (var prop in properties)
                //{
                //    props.Add(prop.Name);
                //}                

                if (obj.getPropertySetInfo() != null && obj.getPropertySetInfo().hasPropertyByName(propName))
                {
                    try
                    {
                        obj.setPropertyValue(propName, Any.Get(value));
                        return true;
                    }
                    catch (unoidl.com.sun.star.beans.PropertyVetoException)
                    {
                        System.Diagnostics.Debug.WriteLine("Property value is unacceptable");
                    }
                    catch (unoidl.com.sun.star.uno.RuntimeException)
                    {
                        System.Diagnostics.Debug.WriteLine("Property value set cause internal error");
                        Logger.Instance.Log(LogPriority.IMPORTANT, "OoUtils", "[FATAL ERROR] Cannot set property '" + propName + "' to value: '" + value.ToString() + "' for the current object");
                    }
                }
                else
                {
                    XPropertyContainer pc = obj as XPropertyContainer;
                    if (pc != null)
                    {
                        try
                        {
                            pc.addProperty(propName, (short)2, Any.Get(value));
                            return true;
                        }
                        catch { }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Sets an Integer property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value.</param>
        public static void SetIntProperty(Object obj, String propName, int value)
        {
            SetProperty(obj, propName, value);
        }

        /// <summary>
        /// Sets a Boolean property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value.</param>
        public static void SetBooleanProperty(Object obj, String propName, bool value)
        {
            SetProperty(obj, propName, value);
        }

        /// <summary>
        /// Sets a String property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value.</param>
        public static bool SetStringProperty(Object obj, String propName, String value)
        {
            return SetProperty(obj, propName, value);
        }

        #endregion

        #region Undo / Redo

        //TODO:

        /// <summary>
        /// Sets a property and add the change in the uno manager.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <param name="doc">The undo manager.</param>
        public static bool SetPropertyUndoable(XPropertySet obj, String propName, Object value, XUndoManagerSupplier doc)
        {
            bool success = false;
            XUndoManager undoManager = null;
            if (doc != null)
            {
                undoManager = ((XUndoManagerSupplier)doc).getUndoManager() as XUndoManager;
                if (undoManager != null)
                {
                    var is_lock = undoManager.isLocked();
                    undoManager.enterUndoContext("Change Property : " + propName);
                }
            }

            var oldVal = GetProperty(obj, propName);
            success = SetProperty(obj, propName, value);

            if (doc != null)
            {
                if (undoManager != null)
                {
                    undoManager.addUndoAction(new ParameterUndo("Change a Prop", obj, propName, oldVal, value));
                    undoManager.leaveUndoContext();
                }
            }
            return success;
        }

        /// <summary>
        /// Sets a property and add the change in the uno manager.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <param name="doc">The undo manager.</param>
        public static void SetPropertyUndoable(Object obj, String propName, Object value, XUndoManagerSupplier doc)
        {
            if (obj is XControl)
            {
                obj = ((XControl)obj).getModel();
            }

            // We need the XPropertySet interface.
            XPropertySet objXPropertySet;
            if (obj is XPropertySet)
            {
                // If the right interface was passed in, just typecaset it.
                objXPropertySet = (XPropertySet)obj;
            }
            else
            {
                // Get a different interface to the drawDoc.
                // The parameter passed in to us is the wrong interface to the object.
                objXPropertySet = obj as XPropertySet;
            }

            // Now just call our sibling using the correct interface.
            SetPropertyUndoable(objXPropertySet, propName, value, doc);
        }

        #endregion

        #endregion

        #region GET

        /// <summary>
        /// Gets a property value.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>
        /// the property form the property set or null
        /// </returns>
        public static Object GetProperty(Object obj, String propName)
        {
            // We need the XPropertySet interface.
            XPropertySet objXPropertySet;

            try
            {
                if (obj is XControl) { obj = ((XControl)obj).getModel(); }
            }
            catch { }

            // Get a different interface to the drawDoc.
            // The parameter passed in to us is the wrong interface to the object.
            objXPropertySet = obj as XPropertySet;

            // Now just call our sibling using the correct interface.
            return GetProperty(objXPropertySet, propName);
        }

        /// <summary>
        /// Gets a property value.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>
        /// the property form the property set or null
        /// </returns>
        public static Object GetProperty(XPropertySet obj, String propName)
        {
            try
            {
                if (obj != null)
                {
                    XPropertySetInfo propInfo = obj.getPropertySetInfo();
                    if (propInfo != null && propInfo.hasPropertyByName(propName))
                    {
                        return obj.getPropertyValue(propName).Value;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Gets an Integer property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>the value of the property</returns>
        public static int GetIntProperty(Object obj, String propName)
        {
            object value = GetProperty(obj, propName);
            if (value != null)
            {
                if (value is int || value is Int32 || value is Int16)
                {
                    var intValue = (int)Convert.ToInt32(value);
                    return intValue;
                }
            }
            return 0;
        }

        /// <summary>
        /// Gets a Boolean property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>the value of the property</returns>
        public static bool GetBooleanProperty(Object obj, String propName)
        {
            object value = GetProperty(obj, propName);
            if (value != null)
            {
                if (value is bool)
                {
                    var booleanValue = (bool)value;
                    return booleanValue;
                }
            }
            return false;
        }

        /// <summary>
        /// Gets a String property.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>the value of the property</returns>
        public static String GetStringProperty(Object obj, String propName)
        {
            Object value = GetProperty(obj, propName);
            if (value != null)
            {
                var stringValue = value as string;
                if (stringValue != null)
                {
                    return stringValue;
                }
            }
            return "";
        }

        #endregion

        #endregion

        #region Implementation ID

        /// <summary>
        /// Try to get the unique implementation id of this element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns>the runtime unique implementation id byte array for this OpenOffice object</returns>
        public static byte[] GetImplementationId(Object element)
        {
            return GetImplementationId(element as XTypeProvider);
        }
        /// <summary>
        /// Try to get the unique implementation id of this element.
        /// This is NOT a hash value of the object. It only gives a hint if 
        /// objects are of the same object type!! Similar object types should 
        /// result in the same implementation id.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns>the runtime unique implementation id byte array for this OpenOffice object</returns>
        public static byte[] GetImplementationId(XTypeProvider element)
        {
            byte[] id = new byte[0];
            if (element != null)
            {
                id = element.getImplementationId();
            }
            return id;
        }

        /// <summary>
        /// Try to get the unique implementation id of this element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns>the runtime unique implementation id byte array as a string for this OpenOffice object</returns>
        public static String GetImplementationIdString(Object element)
        {
            return getStringFromByte(GetImplementationId(element));
        }

        #endregion

        /// <summary>
        /// Try to compare two OpenOffice objects to be equal.
        /// Therefore at least the implementation id is compared. 
        /// </summary>
        /// <param name="a">Object a.</param>
        /// <param name="b">Object b.</param>
        /// <returns></returns>
        public static bool AreOoObjectsEqual(Object a, Object b)
        {
            if (a.Equals(b)) return true;
            else
            {
                if (a != null && b != null)
                {
                    byte[] idA = GetImplementationId(a);
                    byte[] idB = GetImplementationId(b);

                    if (!idA.Equals(new byte[0]) && !idB.Equals(new byte[0]))
                    {
                        return idA.Equals(idB);
                    }
                }
            }
            return false;
        }


        //----------------------------------------------------------------------
        //  Colors.
        //----------------------------------------------------------------------

        /// <summary>
        /// Gets a color number corresponding to the color representation of OO.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="g">The g.</param>
        /// <param name="b">The b.</param>
        /// <returns></returns>
        public static int ConvertToColorInt(byte r, byte g, byte b)
        {
            /*
                * Color 
                   Defining Type
                       long

                   Description
                       describes an RGB color value with an optional alpha channel. 

                       The byte order is from high to low:

                           alpha channel
                           red
                           green
                           blue
                */
            //TODO: the alpha value is not practicable - have to be 0
            try
            {
                var cb = new byte[] { b, g, r, 0 };
                return BitConverter.ToInt32(cb, 0);
            }
            catch (System.Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// Gets a color number corresponding to the color representation of OO.
        /// </summary>
        /// <param name="c">The c.</param>
        /// <returns></returns>
        public static int ConvertToColorInt(System.Drawing.Color c)
        {
            return ConvertToColorInt(BitConverter.GetBytes(c.R)[0], BitConverter.GetBytes(c.G)[0], BitConverter.GetBytes(c.B)[0]);
        }

        //----------------------------------------------------------------------
        //  Sugar coating for com.sun.star.container.XNamed interface.
        //----------------------------------------------------------------------

        /// <summary>
        /// Gets a named property of an XNameAccess object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>The property or null</returns>
        public static Object GetNamedProperty(Object obj, String propName) { return GetNamedProperty(obj as XNameAccess, propName); }

        /// <summary>
        /// Gets a named property of an XNameAccess object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns>The property or null</returns>
        public static Object GetNamedProperty(XNameAccess obj, String propName)
        {
            Object prop = null;
            if (obj != null)
            {
                if (obj.hasByName(propName))
                {
                    prop = obj.getByName(propName);
                }
            }
            return prop;
        }

        /// <summary>
        /// Sets a named property of an XNameAccess object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The new value.</param>
        public static void SetNamedProperty(Object obj, String propName, Object value) { SetNamedProperty(obj as XNameContainer, propName, value); }

        /// <summary>
        /// Sets a named property of an XNameAccess object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The new value.</param>
        public static void SetNamedProperty(XNameContainer obj, String propName, Object value)
        {
            if (obj != null)
            {
                if (obj.hasByName(propName))
                {
                    obj.replaceByName(propName, Any.Get(value));
                }
                else
                {
                    obj.insertByName(propName, new uno.Any(value.ToString()));
                }
            }
        }

        /// <summary>
        /// Sets the name of an object that has XNamed interface.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="name">The new name of the object.</param>
        public static void XNamedSetName(Object obj, String name)
        {
            var xNamed = obj as XNamed;
            if (xNamed != null)
            {
                xNamed.setName(name);
            }
        }

        /// <summary>
        /// Gets the name of an object that has XNamed interface.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>The name string or empty string</returns>
        public static String XNamedGetName(Object obj)
        {
            var xNamed = obj as XNamed;
            return xNamed != null ? xNamed.getName() : String.Empty;
        }

        //----------------------------------------------------------------------
        //  Service Testing
        //----------------------------------------------------------------------
        /// <summary>
        /// Determines if the given object supports the requested service.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="service">The service.</param>
        /// <returns><c>true</c> if the elemnt support the requested service, otherwise <c>false</c></returns>
        public static bool ElementSupportsService(Object element, String service)
        {
            if (element != null && element is XServiceInfo)
            {
                return ((XServiceInfo)element).supportsService(service);
            }
            return false;
        }

        //----------------------------------------------------------------------
        //  Basic tool
        //----------------------------------------------------------------------

        /// <summary>
        /// Gets the Oo process ID.
        /// </summary>
        /// <returns>The integer id of the process otherwise -1</returns>
        public static int GetOoProcessID()
        {
            Process[] localByName = Process.GetProcessesByName("soffice");
            return (localByName != null && localByName.Length > 0) ? localByName[0].Id : -1;
        }

        public static int GetOoBinProcessID()
        {
            Process[] localByName = Process.GetProcessesByName("soffice.bin");
            return (localByName != null && localByName.Length > 0) ? localByName[0].Id : -1;
        }

        /// <summary>
        /// Try to get the window handle of the window peer for WIN_32 Systems.
        /// </summary>
        /// <param name="windowPeer">The window peer.</param>
        /// <returns>the window handle or IntPtr.Zero</returns>
        public static IntPtr GetWindowHandle(XSystemDependentWindowPeer windowPeer)
        {
            IntPtr wHnd = IntPtr.Zero;
            if (windowPeer != null)
            {
                int pId = OoUtils.GetOoProcessID();
                if (pId > 0)
                {
                    var whnd = windowPeer.getWindowHandle(BitConverter.GetBytes(pId), (short)1);
                    if (whnd.hasValue())
                    {
                        try
                        {
                            int val = (int)whnd.Value;
                            wHnd = new IntPtr(val);
                        }
                        catch { }
                    }
                }
            }
            return wHnd;
        }

        /// <summary>
        /// Converts a String into its corresponding byte array.
        /// </summary>
        /// <param name="str">The String.</param>
        /// <returns>the byte array that corresponds to the string</returns>
        static byte[] getBytesFromString(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        /// <summary>
        /// Converts a byte array into a String.
        /// </summary>
        /// <param name="bytes">The byte array.</param>
        /// <returns>A string that corresponds to the byte array</returns>
        static string getStringFromByte(byte[] bytes)
        {
            char[] chars = new char[bytes.Length / sizeof(char)];
            System.Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length);
            return new string(chars);
        }

        /// <summary>
        /// Return the currently active top-level window, i.e. which has currently the input focus.
        /// </summary>
        /// <returns>The returned reference may be empty when no top-level window is active.</returns>
        public static object GetActiveTopWindow()
        {
            var etk = OO.GetExtTooklkit();
            if (etk != null)
            {
                return etk.getActiveTopWindow();
            }
            return null;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        public static Object GetGraphicFromUrl(string url)
        {
            try
            {
                var msf = OO.GetMultiComponentFactory();
                if (msf != null)
                {
                    var graphicProvider = msf.createInstanceWithContext(OO.Services.GRAPHIC_GRAPHICPROVIDER, OO.GetContext());
                    if (graphicProvider != null && graphicProvider is XGraphicProvider)
                    {
                        PropertyValue val = new PropertyValue();
                        val.Name = "URL";

                        val.Value = Any.Get(CreateBitmapUrl(url));
                        PropertyValue val2 = new PropertyValue();
                        val2.Name = "MimeType";
                        val2.Value = Any.Get(@"image/png");
                        var graphic = ((XGraphicProvider)graphicProvider).queryGraphic(new PropertyValue[] { val });
                        return graphic;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "oOUtils", "Cannot get Graphic provider: " + ex);
            } return null;

        }

        public static string CreateBitmapUrl(string path)
        {
            return @"file:///" + path.Replace("\\", "/");
        }


        #region Undo / Redo

        /// <summary>
        /// Adds an action to an undo manager.
        /// </summary>
        /// <param name="undoManager">The undo manager supplier (normally the document).</param>
        /// <param name="undo">The undo-action to add.</param>
        public static void AddActionToUndoManager(XUndoManagerSupplier undoManager, XUndoAction undo)
        {
            AddActionToUndoManager(undoManager.getUndoManager(), undo);
        }

        /// <summary>
        /// Adds an action to an undo manager.
        /// </summary>
        /// <param name="undoManager">The undo manager.</param>
        /// <param name="undo">The undo-action to add.</param>
        public static void AddActionToUndoManager(XUndoManager undoManager, XUndoAction undo)
        {
            if (undoManager != null && undo != null)
            {
                undoManager.enterUndoContext(undo.Title);
                undoManager.addUndoAction(undo);
                undoManager.leaveUndoContext();
            }
        }

        #endregion
    }

    /// <summary>
    /// Class implementing the XUndoAction interface. This can be added to an
    /// Undo manager to support the undo redo of an property change.
    /// </summary>
    public class ParameterUndo : XUndoAction
    {

        private readonly Object Target;
        readonly Dictionary<String, Object> OldParameters;
        readonly Dictionary<String, Object> NewParameters;

        public ParameterUndo(string title, Object target, Dictionary<String, Object> oldParameters, Dictionary<String, Object> newParameters)
        {
            Title = title;
            OldParameters = oldParameters;
            NewParameters = newParameters;
            Target = target;
        }

        public ParameterUndo(string title, Object target, String name, Object oldValue, Object newValue)
        {
            Dictionary<String, Object> oldParams = new Dictionary<String, Object>();
            oldParams.Add(name, oldValue);
            Dictionary<String, Object> newParams = new Dictionary<String, Object>();
            newParams.Add(name, newValue);
            Target = target;
            OldParameters = oldParams;
            NewParameters = newParams;
            Title = title;

        }

        public string Title
        {
            get;
            private set;
        }

        public void redo()
        {
            if (NewParameters != null)
            {
                foreach (var item in NewParameters.Keys)
                {
                    try
                    {
                        util.OoUtils.SetProperty(Target, item.ToString(), NewParameters[item]);
                    }
                    catch { }
                }
            }
        }

        public void undo()
        {
            if (OldParameters != null)
            {
                foreach (var item in OldParameters.Keys)
                {
                    try
                    {
                        util.OoUtils.SetProperty(Target, item.ToString(), OldParameters[item]);
                    }
                    catch { }
                }
            }
        }
    }

    /// <summary>
    /// Implementation for an undo- redo-action. An instance of 
    /// this class can be added to an undo-redo-manager of open office.
    /// </summary>
    public class ActionUndo : XUndoAction
    {

        Action _redo;
        Action _undo;

        public string Title
        {
            get;
            private set;
        }

        public ActionUndo(String title, Action undo, Action redo)
        {
            Title = title;
            _undo = undo;
            _redo = redo;

        }

        public void redo()
        {
            try { _redo.Invoke(); }
            catch { }
        }

        public void undo()
        {
            try { _undo.Invoke(); }
            catch { }
        }
    }

}
