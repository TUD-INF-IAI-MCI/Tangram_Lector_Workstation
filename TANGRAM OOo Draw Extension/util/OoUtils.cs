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
using unoidl.com.sun.star.datatransfer;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.view;

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
    /// <summary>
    /// Basic static helper functions for handling OpenOffice/LibreOffice   
    /// </summary>
    public class OoUtils
    {
        #region Structures creation

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

        #endregion

        #region Property

        #region Properties SET

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
        internal static bool SetProperty(XPropertySet obj, String propName, Object value)
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
        public static bool SetIntProperty(Object obj, String propName, int value)
        {
            return SetProperty(obj, propName, value);
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

        /// <summary>
        /// Sets a property and add the change in the uno manager.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <param name="doc">The undo manager.</param>
        internal static bool SetPropertyUndoable(XPropertySet obj, String propName, Object value, XUndoManagerSupplier doc)
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
        /// <param name="obj">The obj to change its property [XControl], [XPropertySet].</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <param name="doc">The undo manager.</param>
        /// <returns><c>true</c> if the property could be changed successfully</returns>
        internal static bool SetPropertyUndoable(Object obj, String propName, Object value, XUndoManagerSupplier doc)
        {
            if (obj is XControl)
            {
                obj = ((XControl)obj).getModel();
            }

            // We need the XPropertySet interface.
            XPropertySet objXPropertySet;
            if (obj is XPropertySet)
            {
                // If the right interface was passed in, just typecast it.
                objXPropertySet = (XPropertySet)obj;
            }
            else
            {
                // Get a different interface to the drawDoc.
                // The parameter passed in to us is the wrong interface to the object.
                objXPropertySet = obj as XPropertySet;
            }

            // Now just call our sibling using the correct interface.
            return SetPropertyUndoable(objXPropertySet, propName, value, doc);
        }


        /// <summary>
        /// Sets a property and add the change in the uno manager.
        /// </summary>
        /// <param name="obj">The obj to change its property [XControl], [XPropertySet].</param>
        /// <param name="propName">Name of the property.</param>
        /// <param name="value">The value of the property.</param>
        /// <param name="undoManager">The undo manager [XUndoManagerSupplier] - normally this is 
        /// the document (DrawPagesSupplier: SERVICE com.sun.star.document.OfficeDocument).</param>
        /// <returns>
        ///   <c>true</c> if the property could be changed successfully
        /// </returns>
        public static bool SetPropertyUndoable(Object obj, String propName, Object value, Object undoManager)
        { 
            return SetPropertyUndoable(obj, propName, value, undoManager as XUndoManagerSupplier); 
        }


        #endregion

        #endregion

        #region Properties GET

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
        internal static Object GetProperty(XPropertySet obj, String propName)
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
        /// <remarks>ATTENTION: This function is time limited to 100 ms!</remarks>
        public static String GetStringProperty(Object obj, String propName)
        {
            String result = String.Empty;
            bool succ = TimeLimitExecutor.WaitForExecuteWithTimeLimit(100,
                 new Action(() =>
                 {
                     Object value = GetProperty(obj, propName);
                     if (value != null)
                     {
                         var stringValue = value as string;
                         if (stringValue != null)
                         {
                             result = stringValue;
                         }
                     }
                 }));
            if (!succ) result = String.Empty;
            return result;
        }

        internal static Dictionary<String, uno.Any> GetAllProperties(Object obj)
        {
            // Debug.GetAllInterfacesOfObject(obj);
            Dictionary<String, uno.Any> props = new Dictionary<string, uno.Any>();

            if (obj != null && obj is XPropertySet)
            {
                XPropertySetInfo bla = ((XPropertySet)obj).getPropertySetInfo();
                if (bla != null)
                {
                    Property[] properties = bla.getProperties();
                    List<String> storable = new List<string>();

                    foreach (Property item in properties)
                    {
                        if (item != null)
                        {
                            OO.PropertyAttribute att = ((OO.PropertyAttribute)item.Attributes);
                            //if (item.Attributes == 0) { storable.Add(item.Name); }
                            if (!att.HasFlag(OO.PropertyAttribute.READONLY))
                            {
                                //if (att.HasFlag(OO.PropertyAttribute.MAYBEDEFAULT))
                                //{
                                //}
                                storable.Add(item.Name);
                            }
                        }
                    }

                    try
                    {
                        if (obj is XMultiPropertySet)
                        {
                            var values = ((XMultiPropertySet)obj).getPropertyValues(storable.ToArray());

                            if (values.Length == storable.Count)
                            {
                                for (int i = 0; i < values.Length; i++)
                                {
                                    props.Add(storable[i], values[i]);
                                }
                            }
                            else
                            {
                                // does not get all values
                            }

                            return props;
                        }
                    }
                    catch (Exception ex) { }
                }
            }
            return null;
        }

        internal static bool SetAllProperties(Object obj, Dictionary<String, uno.Any> props)
        {
            try
            {
                if (obj is XMultiPropertySet)
                {

                    String[] keys = new String[props.Count];
                    props.Keys.CopyTo(keys, 0);
                    uno.Any[] vals = new uno.Any[props.Count];
                    props.Values.CopyTo(vals, 0);

                    ((XMultiPropertySet)obj).setPropertyValues(keys, vals);
                }
            }
            catch (Exception ex) { }

            return false;
        }


        internal static Object CreateObjectFromService(Object _msf, String[] services)
        {
            try
            {
                // get MSF
                XMultiServiceFactory msf = _msf as XMultiServiceFactory;

                if (msf == null)
                    msf = OO.GetMultiServiceFactory();
                if (msf != null && services != null && services.Length > 0)
                {

                    string[] serv = msf.getAvailableServiceNames();
                    System.Diagnostics.Debug.WriteLine("Available Service Names: " + String.Join("\n\t", serv));

                    //object component = msf.createInstance(services[0]);
                    // object component = msf.createInstance("com.sun.star.document.ExportGraphicObjectResolver");
                    object component = msf.createInstance("com.sun.star.document.ImportEmbeddedObjectResolver");

                    //Debug.GetAllServicesOfObject(component);
                    Debug.GetAllInterfacesOfObject(component);

                    var n = ((XNameAccess)component).getElementNames();


                    return component;
                }

            }
            catch (Exception ex)
            {
            }

            return null;
        }


        #endregion


        #region ProprtyValue Struct handling

        /// <summary>
        /// Gets a dictionary of the properties with the Name as key and the PropertyValue as key.
        /// </summary>
        /// <param name="props">The props.</param>
        /// <returns>Dictionary of indexed Names.</returns>
        internal static Dictionary<String, PropertyValue> GetPropertyvalueDictionary(PropertyValue[] props)
        {
            Dictionary<String, PropertyValue> propDict = new Dictionary<string, PropertyValue>();

            if (props != null && props.Length > 0)
            {
                foreach (PropertyValue prop in props)
                {
                    propDict.Add(prop.Name, prop);
                }
            }

            return propDict;
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
            return GetStringFromByte(GetImplementationId(element));
        }

        #endregion

        #region Colors

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

        /// <summary>
        /// Gets a randomized color.
        /// </summary>
        /// <param name="minDarkness">The minimum darkness.</param>
        /// <returns></returns>
        public static int GetRandomizedColor(int minDarkness = 0)
        {
            var rand = new Random(DateTime.Now.Millisecond);

            return ConvertToColorInt(
                BitConverter.GetBytes(rand.Next(minDarkness, 255))[0],
                BitConverter.GetBytes(rand.Next(minDarkness, 255))[0],
                BitConverter.GetBytes(rand.Next(minDarkness, 255))[0]);
        } 


        #endregion

        #region XNamed

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

        #endregion

        #region  Service Testing

        /// <summary>
        /// Determines if the given object supports the requested service.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="service">The service.</param>
        /// <returns><c>true</c> if the element support the requested service, otherwise <c>false</c></returns>
        public static bool ElementSupportsService(Object element, String service)
        {
            if (element != null && element is XServiceInfo)
            {
                return ((XServiceInfo)element).supportsService(service);
            }
            return false;
        }

        #endregion

        #region Basic Helper Functions

        /// <summary>
        /// Try to compare two OpenOffice objects to be equal.
        /// Therefore at least the implementation id is compared. 
        /// </summary>
        /// <param name="a">Object a.</param>
        /// <param name="b">Object b.</param>
        /// <returns></returns>
        public static bool AreOoObjectsEqual(Object a, Object b)
        {
            if (a != null && a.Equals(b)) return true;
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

        /// <summary>
        /// Gets the Oo process ID.
        /// </summary>
        /// <returns>The integer id of the process otherwise -1</returns>
        public static int GetOoProcessID()
        {
            Process[] localByName = Process.GetProcessesByName("soffice");
            return (localByName != null && localByName.Length > 0) ? localByName[0].Id : -1;
        }

        /// <summary>
        /// Gets the process id of the soffice.bin process.
        /// </summary>
        /// <returns>the id of the office process if it is running; otherwise -1</returns>
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
        public static byte[] GetBytesFromString(string str)
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
        public static string GetStringFromByte(byte[] bytes)
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

        #endregion

        #region Graphic Handling

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

        /// <summary>
        /// Creates the bitmap URL from a file path, so it can be loaded.
        /// </summary>
        /// <param name="path">The path to the bitmap.</param>
        /// <returns>a confirm bitmap URL string</returns>
        public static string CreateBitmapUrl(string path)
        {
            return @"file:///" + path.Replace("\\", "/");
        }

        #endregion

        #region Coordination Convention

        #region PInvoke user32.dll

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        #endregion

        /// <summary>
        /// Gets the device resolution as pixel per meter.
        /// </summary>
        /// <param name="_ppmX">The pixel per meter in x dimension.</param>
        /// <param name="_ppmY">The pixel per meter in y dimension.</param>
        public static void GetDeviceResolutionPixelPerMeter(out double _ppmX, out double _ppmY)
        {
            int dpiX = 96, dpiY = 96;
            // get DPI resolution x and y
            IntPtr hDC = GetDC(IntPtr.Zero);    // device context for the whole screen, see http://pinvoke.net/default.aspx/gdi32.GetDC
            dpiX = GetDeviceCaps(hDC, 88 /*LOGPIXELSX*/);   // see http://pinvoke.net/default.aspx/gdi32/GetDeviceCaps.html
            dpiY = GetDeviceCaps(hDC, 90 /*LOGPIXELSY*/);
            ReleaseDC(IntPtr.Zero, hDC);        // release device context
            _ppmX = ((double)dpiX / 25.4) * 1000.0;
            _ppmY = ((double)dpiX / 25.4) * 1000.0;
        }

        #endregion

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

        #region Copy / Paste

        internal static XTransferableSupplier FindTransferableSupplier(XModel2 xModel2)
        {
            XTransferableSupplier result = null;

            try
            {
                if (xModel2 != null)
                {
                    XEnumeration xEnumeration = xModel2.getControllers();

                    while (xEnumeration.hasMoreElements())
                    {
                        XController xController = (XController)xEnumeration.nextElement().Value;
                        if (xController is XTransferableSupplier)
                        {
                            result = xController as XTransferableSupplier;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }

            return result;
        }


        public static void Copy(Object document, Object selection)
        {

            Debug.GetAllInterfacesOfObject(document);


            if (document is XModel)
            {
                //return 
                Copy(document as XModel, selection);
            }
            else if (document is XController)
            {
                //return 
                Copy(document as XController, selection);
            }
        }

        internal static void Copy(XModel document, Object selection)
        {

            if (document != null)
            {
                //return 
                Copy(document.getCurrentController(), selection);
            }

        }

        internal static void Copy(XController controller, Object selection)
        {
            XTransferableSupplier xTransferableSupplier = controller as XTransferableSupplier;

            Debug.GetAllInterfacesOfObject(controller);


            Copy(controller, selection, xTransferableSupplier);
        }

        private static void Copy(XController controller, Object selection, XTransferableSupplier xTransferableSupplier)
        {
            if (xTransferableSupplier != null)
            {
                // get selection supplier
                XSelectionSupplier xSelectionSupplier = controller as XSelectionSupplier;

                if (xSelectionSupplier != null)
                {
                    // select the objects to select
                    try
                    {
                        xSelectionSupplier.select(Any.Get(selection));
                    }
                    catch (Exception ex)
                    {
                        throw new ArgumentException("Object can't be set as selection", "selection", ex);
                    }

                    XTransferable transfarable = xTransferableSupplier.getTransferable();


                }
            }
        }


        #endregion

    }

    #region Undo / Redo Classes

    /// <summary>
    /// Class implementing the XUndoAction interface. This can be added to an
    /// Undo manager to support the undo redo of an property change.
    /// </summary>
    public class ParameterUndo : XUndoAction
    {

        private readonly Object Target;
        readonly Dictionary<String, Object> OldParameters;
        readonly Dictionary<String, Object> NewParameters;

        /// <summary>
        /// Initializes a new instance of the <see cref="ParameterUndo"/> class.
        /// A small handler for changing object properties.
        /// </summary>
        /// <param name="title">The title of the done operation - will be listed in the undo/redo history.</param>
        /// <param name="target">The target object to change properties of.</param>
        /// <param name="oldParameters">The old parameters and values.</param>
        /// <param name="newParameters">The new parameters and values.</param>
        public ParameterUndo(string title, Object target, Dictionary<String, Object> oldParameters, Dictionary<String, Object> newParameters)
        {
            Title = title;
            OldParameters = oldParameters;
            NewParameters = newParameters;
            Target = target;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ParameterUndo" /> class.
        /// A small handler for changing object properties.
        /// </summary>
        /// <param name="title">The title of the done operation - will be listed in the undo/redo history.</param>
        /// <param name="target">The target object to change properties of.</param>
        /// <param name="name">The name of the property to change.</param>
        /// <param name="oldValue">The old value.</param>
        /// <param name="newValue">The new value.</param>
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

        /// <summary>
        /// The title of the done operation - will be listed in the undo/redo history.
        /// </summary>
        /// <value>
        /// The title.
        /// </value>
        public string Title
        {
            get;
            private set;
        }

        /// <summary>
        /// sets the new properties again.
        /// </summary>
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

        /// <summary>
        /// restore the old properties.
        /// </summary>
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

        /// <summary>
        /// Is the human-readable, localized description of the action.
        /// </summary>
        /// <value>
        /// The title.
        /// </value>
        public string Title
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ActionUndo"/> class.
        /// </summary>
        /// <param name="title">The title of this change - will be displayed in the undo/redo history.</param>
        /// <param name="undo">The undo action.</param>
        /// <param name="redo">The redo action.</param>
        public ActionUndo(String title, Action undo, Action redo)
        {
            Title = title;
            _undo = undo;
            _redo = redo;

        }

        /// <summary>
        /// Repeats the action represented by the instance, after it had previously been reverted.
        /// </summary>
        public void redo()
        {
            try { _redo.Invoke(); }
            catch { }
        }

        /// <summary>
        /// Reverts the action represented by the instance  .
        /// </summary>
        public void undo()
        {
            try { _undo.Invoke(); }
            catch { }
        }
    }

    #endregion

}