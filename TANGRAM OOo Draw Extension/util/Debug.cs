// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-18-2012
//
// Last Modified By : Admin
// Last Modified On : 09-18-2012
// ***********************************************************************
// <copyright file="Debug.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.beans;
using System.Reflection;
using System.Text;
using System.Collections;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.xforms;

namespace tud.mci.tangram.util
{
    public static class Debug
    {

        /// <summary>
        /// Gets all interfaces of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the interfaces will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns></returns>
        public static Type[] GetAllInterfacesOfObject(Object obj, bool debug = true)
        {
            String output = "";
            if (obj != null)
            {
                try
                {
                    XTypeProvider tp = obj as XTypeProvider;


                    if (tp != null)
                    {
                        var interfaces = tp.getTypes();

                        if (debug)
                        {
                            output += (interfaces.Length + "\tTypes:");

                            foreach (var item in interfaces)
                            {
                                output += "\n" + "\tType: " + item;
                            }
                        }
                        System.Diagnostics.Debug.WriteLine(output);
                        return interfaces;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("Can't get interfaces of Object: " + obj);
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Can't get interfaces of Object: " + e);
                }
            }
            return new Type[0];
        }


        /// <summary>
        /// Returns all services of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the services will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns></returns>
        public static string[] GetAllServicesOfObject(Object obj, bool debug = true)
        {
            String output = "";
            string[] services = new string[0];
            if (obj != null)
            {
                try
                {
                    TimeLimitExecutor.WaitForExecuteWithTimeLimit(200, () =>
                    {

                        XServiceInfo si = obj as XServiceInfo;
                        if (si != null)
                        {
                            services = si.getSupportedServiceNames();

                            if (debug)
                            {
                                output += (services.Length + "\tServices:");

                                foreach (var item in services)
                                {
                                    output += "\n" + ("\tService: " + item);
                                }
                                System.Diagnostics.Debug.WriteLine(output);
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("Can't get services of Object because it is not a XServiceInfo provider.");
                        }
                    }, "GetAllServices");
                }
                catch (unoidl.com.sun.star.lang.DisposedException de) { OO.CheckConnection(); }                   
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Can't get services of Object: " + e);
                }
            }
            return services;
        }

        /// <summary>
        /// prints all properties of the object to the debug output.
        /// </summary>
        /// <param name="obj">The obj.</param>
        public static void PrintAllProperties(Object obj)
        {
            GetAllProperties(obj, true);
        }

        /// <summary>
        /// Gets all properties of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the properies will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns>A dictionary with the properties. The name is the key an the value 
        /// is set as uno.Any. You can access the real value by using uno.Any.Value</returns>
        public static Dictionary<String, uno.Any> GetAllProperties(Object obj, bool debug = true)
        {
            if (obj != null)
            {
                if (obj is XControl)
                {
                    XControlModel model = ((XControl)obj).getModel();
                    return GetAllProperties(model as XPropertySet, debug);
                }else
                return GetAllProperties(obj as XPropertySet, debug);
            }
            return null;
        }

        /// <summary>
        /// Gets all properties of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the properties will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns>A dictionary with the properties. The name is the key an the value 
        /// is set as uno.Any. You can access the real value by using uno.Any.Value</returns>
        public static Dictionary<String, uno.Any> GetAllProperties(XPropertySet obj, bool debug = true)
        {
            String output = "";
            if (obj != null)
            {
                try
                {
                    var b = obj.getPropertySetInfo();
                    if (b != null)
                    {
                        Property[] a = b.getProperties();
                        if (debug)
                            System.Diagnostics.Debug.WriteLine("Object [" + obj + "] has " + a.Length + " Properties:");

                        var properties = new Dictionary<String, uno.Any>();

                        foreach (var item in a)
                        {
                            try
                            {
                                if (obj != null && item != null && item.Name != null && obj.getPropertyValue(item.Name).hasValue())
                                {
                                    if (debug)
                                        output += "\n" + ("\tProperty: " + item.Name + " = " + obj.getPropertyValue(item.Name));
                                    properties.Add(item.Name, obj.getPropertyValue(item.Name));
                                }
                                else
                                {
                                    output += "\n" + "Can't get property - " + item.Name;
                                }
                                
                            }
                            catch (Exception)
                            {
                                output += "\n" + "Can't get property - " + item.Name;
                            }
                        }
                        System.Diagnostics.Debug.WriteLine(output);
                        return properties;
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Can't get properties of object: " + e);
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("object is null and therefore not of type XPropertySet and no properties could be displayed! ");
            }
            return new Dictionary<String, uno.Any>();
        }

        /// <summary>
        /// Gets all named elements of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the named children will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns>
        /// A dictionary with the named elements. The name is the key an the value
        /// is set as uno.Any. You can access the real value by using uno.Any.Value
        /// </returns>
        public static Dictionary<string, uno.Any> GetAllNamedElements(Object obj, bool debug = true) { return GetAllNamedElements(obj as XNameAccess, debug); }
        /// <summary>
        /// Gets all named elements of the object.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="debug">if set to <c>true</c> the named children will be printed
        /// to the System.Diagnostics.Debug output.</param>
        /// <returns>
        /// A dictionary with the named elements. The name is the key an the value
        /// is set as uno.Any. You can access the real value by using uno.Any.Value
        /// </returns>
        public static Dictionary<string, uno.Any> GetAllNamedElements(XNameAccess obj, bool debug = true)
        {
            Dictionary<string, uno.Any> elements = new Dictionary<string, uno.Any>();
            if (obj != null)
            {
                string[] names = obj.getElementNames();

                if (debug)
                    System.Diagnostics.Debug.WriteLine("Object [" + obj + "] has " + names.Length + " Properties:");

                foreach (var name in names)
                {

                    uno.Any element = obj.getByName(name);
                    elements.Add(name, element);

                    if (debug)
                        try
                        {
                            System.Diagnostics.Debug.WriteLine("\tProperty: " + name + " = " + element);
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("Can't get property - " + name);
                        }

                }

            }
            return elements;
        }


        /// <summary>
        /// Var_dump function for the specified obj. [SEEMS NOT TO WORK]
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="recursion">The recursion.</param>
        /// <returns></returns>
        public static string var_dump(object obj, int recursion = 1)
        {
            StringBuilder result = new StringBuilder();

            // Protect the method against endless recursion
            if (recursion < 5)
            {
                // Determine object type
                Type t = obj.GetType();

                // Get array with properties for this object
                PropertyInfo[] properties = t.GetProperties();

                foreach (PropertyInfo property in properties)
                {
                    try
                    {
                        // Get the property value
                        object value = property.GetValue(obj, null);

                        // Create indenting string to put in front of properties of a deeper level
                        // We'll need this when we display the property name and value
                        string indent = String.Empty;
                        string spaces = "|   ";
                        string trail = "|...";

                        if (recursion > 0)
                        {
                            indent = new StringBuilder(trail).Insert(0, spaces, recursion - 1).ToString();
                        }

                        if (value != null)
                        {
                            // If the value is a string, add quotation marks
                            string displayValue = value.ToString();
                            if (value is string) displayValue = String.Concat('"', displayValue, '"');

                            // Add property name and value to return string
                            result.AppendFormat("{0}{1} = {2}\n", indent, property.Name, displayValue);

                            try
                            {
                                if (!(value is ICollection))
                                {
                                    // Call var_dump() again to list child properties
                                    // This throws an exception if the current property value
                                    // is of an unsupported type (eg. it has not properties)
                                    result.Append(var_dump(value, recursion + 1));
                                }
                                else
                                {
                                    // 2009-07-29: added support for collections
                                    // The value is a collection (eg. it's an arraylist or generic list)
                                    // so loop through its elements and dump their properties
                                    int elementCount = 0;
                                    foreach (object element in ((ICollection)value))
                                    {
                                        string elementName = String.Format("{0}[{1}]", property.Name, elementCount);
                                        indent = new StringBuilder(trail).Insert(0, spaces, recursion).ToString();

                                        // Display the collection element name and type
                                        result.AppendFormat("{0}{1} = {2}\n", indent, elementName, element.ToString());

                                        // Display the child properties
                                        result.Append(var_dump(element, recursion + 2));
                                        elementCount++;
                                    }

                                    result.Append(var_dump(value, recursion + 1));
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            // Add empty (null) property to return string
                            result.AppendFormat("{0}{1} = {2}\n", indent, property.Name, "null");
                        }
                    }
                    catch
                    {
                        // Some properties will throw an exception on property.GetValue()
                        // I don't know exactly why this happens, so for now i will ignore them...
                    }
                }
            }

            return result.ToString();
        }


        /// <summary>
        /// Gets the readable implementation id.
        /// </summary>
        /// <param name="id">The id.</param>
        /// <returns></returns>
        public static string getReadableImplementationId(byte[] id)
        {
            return String.Join("|", id);
        }

        public static String rectToString(Rectangle rect)
        {
            return (rect.X + ", " + rect.Y + "    " + rect.Width + " x " + rect.Height);
        }

        public static String homogenMatrix3ToString(unoidl.com.sun.star.drawing.HomogenMatrix3 matrix, bool singeLine=true)
        {
            return
                "[" + matrix.Line1.Column1.ToString() + ", " + matrix.Line1.Column2.ToString() + ", " + matrix.Line1.Column2.ToString() + "]" + (singeLine ? String.Empty : "\r\n") +
                "[" + matrix.Line2.Column1.ToString() + ", " + matrix.Line2.Column2.ToString() + ", " + matrix.Line2.Column2.ToString() + "]" + (singeLine ? String.Empty : "\r\n") +
                "[" + matrix.Line3.Column1.ToString() + ", " + matrix.Line3.Column2.ToString() + ", " + matrix.Line3.Column2.ToString() + "]";
        }

    }
}
