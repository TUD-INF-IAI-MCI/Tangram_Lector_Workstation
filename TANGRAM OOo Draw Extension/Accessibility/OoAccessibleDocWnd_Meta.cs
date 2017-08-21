using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using tud.mci.tangram.models.Interfaces;
using unoidl.com.sun.star.document;
using unoidl.com.sun.star.beans;
using tud.mci.tangram.models;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.frame;
using tud.mci.tangram.util;
using System.IO;

namespace tud.mci.tangram.Accessibility
{
    /// <summary>
    /// Encapsulates meta-data properties.
    /// </summary>
    public partial class OoAccessibleDocWnd : IUpdateable, IDisposable, IDisposingObserver
    {

        #region Meta-Data Access


        public DrawDocMetaDataSet MetaData { get; private set; }


        #endregion

        #region File Access

        /// <summary>
        /// Gets the store location of the document if it is already stored.
        /// </summary>
        /// <value>
        /// The store location; otherwise <c>null</c>.
        /// </value>
        public String StoreLocation
        {
            get
            {

                if (this.DrawPageSupplier != null && DrawPageSupplier is XStorable)
                {
                    if (((XStorable)DrawPageSupplier).hasLocation())
                    {
                        return ((XStorable)DrawPageSupplier).getLocation();
                    }
                }
                return String.Empty;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is read only.        /// 
        /// </summary>
        /// <value>
        ///   <c>true</c> if the data store is read only or opened read only; <c>false</c> otherwise.
        /// </value>
        /// <remarks>It is not possible to call store() successfully when the data store is read-only.</remarks>
        public bool IsReadonly
        {
            get
            {
                if (this.DrawPageSupplier != null && DrawPageSupplier is XStorable)
                {
                    return ((XStorable)DrawPageSupplier).isReadonly();
                }
                return true;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this document was modified since the last saving.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this document is modified; otherwise, <c>false</c>.
        /// </value>
        public bool IsModified
        {
            get
            {
                if (DrawPageSupplier != null && DrawPageSupplier is XModifiable)
                {
                    return ((XModifiable)DrawPageSupplier).isModified();
                }
                return false;
            }
            set
            {
                if (DrawPageSupplier != null && DrawPageSupplier is XModifiable)
                {
                    ((XModifiable)DrawPageSupplier).setModified(value);
                }
            }
        }

        /// <summary>
        /// Saves this instance.
        /// 
        /// Stores the data to the URL from which it was loaded. 
        /// Only objects which know their locations can be stored.
        /// </summary>
        /// <returns><c>true if the save call was sent without errors.</c></returns>
        /// <remarks>a <c>true</c> as return value does not mean, that the file was saved successfully.</remarks>
        public bool Save()
        {
            bool success = false;
            if (DrawPageSupplier != null && DrawPageSupplier is XModifiable && !IsReadonly)
            {
                //TimeLimitExecutor.WaitForExecuteWithTimeLimit(1000, () =>
                //{
                    try
                    {
                        if (((XStorable)DrawPageSupplier).hasLocation())
                        {
                            ((XStorable)DrawPageSupplier).store();
                            success = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] Can't store file:", ex);
                    }
                // }, "Store file");
            }
            return success;
        }


        /// <summary>
        /// Saves this instance to the given file path and name.
        /// 
        /// stores the object's persistent data to a URL and makes this 
        /// URL the new location of the object. This is the normal 
        /// behavior for UI's "save-as" feature.
        /// </summary>
        /// <returns><c>true if the save call was sent without errors.</c></returns>
        /// <remarks>a <c>true</c> as return value does not mean, that the file was saved successfully.</remarks>
        public bool SaveAs(String path, string fileType = "draw8")
        {
            bool success = false;
            if (DrawPageSupplier != null && DrawPageSupplier is XModifiable
                && !IsReadonly && !String.IsNullOrWhiteSpace(path))
            {
                var dir = Path.GetDirectoryName(path);
                if (!String.IsNullOrWhiteSpace(dir))
                {
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    if (Directory.Exists(dir) && !String.IsNullOrWhiteSpace(Path.GetFileName(path)))
                    {
                        // TimeLimitExecutor.WaitForExecuteWithTimeLimit(1000, () =>
                        //{
                            try
                            {
                                var overwrite = new PropertyValue();
                                overwrite.Name = "Overwrite";
                                overwrite.Value = Any.Get(true);

                                var filter = new PropertyValue();
                                filter.Name = "FilterName";
                                filter.Value = Any.Get(fileType);

                                PropertyValue[] arguments = new PropertyValue[2]{
                                    overwrite,
                                    filter
                                };

                                path = path.Replace("\\", "/");
                                ((XStorable)DrawPageSupplier).storeAsURL(@"file:///" + path, arguments);
                                success = true;
                            }
                            catch (Exception ex)
                            {
                                Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] Can't store file:", ex);
                            }
                        // }, "Store file");
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Saves this instance to the given file path and name.
        /// 
        /// Behavior for UI's export feature. This method accepts all kinds of export filters, 
        /// not only combined import/export filters because it implements an exporting capability, 
        /// not a persistence capability.
        /// </summary>
        /// <returns><c>true if the save call was sent without errors.</c></returns>
        /// <remarks>a <c>true</c> as return value does not mean, that the file was saved successfully.</remarks>
        public bool SaveTo(String path, string fileType = "draw8")
        {
            bool success = false;
            if (DrawPageSupplier != null && DrawPageSupplier is XModifiable
                && !IsReadonly && !String.IsNullOrWhiteSpace(path))
            {
                var dir = Path.GetDirectoryName(path);
                if (!String.IsNullOrWhiteSpace(dir))
                {
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    if (Directory.Exists(dir) && !String.IsNullOrWhiteSpace(Path.GetFileName(path)))
                    {
                        // TimeLimitExecutor.WaitForExecuteWithTimeLimit(1000, () =>
                        //{
                            try
                            {
                                var overwrite = new PropertyValue();
                                overwrite.Name = "Overwrite";
                                overwrite.Value = Any.Get(true);

                                var filter = new PropertyValue();
                                filter.Name = "FilterName";
                                filter.Value = Any.Get(fileType);

                                PropertyValue[] arguments = new PropertyValue[2]{
                                    overwrite,
                                    filter
                                };

                                path = path.Replace("\\", "/");
                                ((XStorable)DrawPageSupplier).storeToURL(@"file:///" +path, arguments);
                                success = true;
                            }
                            catch (Exception ex)
                            {
                                Logger.Instance.Log(LogPriority.ALWAYS, this, "[FATAL ERROR] Can't store file:", ex);
                            }
                       // }, "Store file");
                    }
                }
            }
            return success;
        }

        #endregion
    }

    /// <summary>
    /// A structure to get access to meta data information
    /// </summary>
    public struct DrawDocMetaDataSet
    {

        readonly XDocumentPropertiesSupplier docPropertySupplier;
        readonly XDocumentProperties xDocumentProperties;

        /// <summary>
        /// Initializes a new instance of the <see cref="DrawDocMetaDataSet"/> struct.
        /// </summary>
        /// <param name="doc">The document.</param>
        public DrawDocMetaDataSet(XDocumentPropertiesSupplier doc)
        {
            xDocumentProperties = null;
            docPropertySupplier = doc;
            if (docPropertySupplier != null)
            {
                xDocumentProperties = docPropertySupplier.getDocumentProperties();
            }
        }

        #region XDocumentProperties

        /// <summary>
        /// The initial author of the document.
        /// </summary>
        /// <value>The author.</value>
        public String Author
        {
            get
            {
                if (xDocumentProperties != null)
                    return xDocumentProperties.Author;
                return String.Empty;
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    xDocumentProperties.Author = value;
                }
            }
        }

        /// <summary>
        /// The date and time when the document was created. 
        /// </summary>
        /// <value>The date of creation .</value>
        public System.DateTime CreationDate
        {
            get
            {
                if (xDocumentProperties != null)
                {
                    try
                    {

                        var date = xDocumentProperties.CreationDate;
                        return new System.DateTime(date.Year, date.Month, date.Day, date.Hours, date.Minutes, date.Seconds, (int)date.NanoSeconds / 1000000);

                    }
                    catch (Exception) { }
                }
                return new System.DateTime();
            }
            set
            {
                try
                {
                    if (xDocumentProperties != null)
                    {

                        var time = value.ToUniversalTime();
                        xDocumentProperties.CreationDate = new unoidl.com.sun.star.util.DateTime(
                            (uint)(time.Millisecond * 1000000),
                            (ushort)time.Second,
                            (ushort)time.Minute,
                            (ushort)time.Hour,
                            (ushort)time.Day,
                            (ushort)time.Month,
                            (short)time.Year,
                            true);
                    }
                }
                catch (Exception) { }
            }
        }

        /// <summary>
        /// The title of the document.
        /// </summary>
        /// <value>The title.</value>
        public String Title
        {
            get
            {
                if (xDocumentProperties != null)
                    return xDocumentProperties.Title;
                return String.Empty;
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    xDocumentProperties.Title = value;
                }
            }
        }

        /// <summary>
        /// the subject of the document.
        /// </summary>
        /// <value>The subject.</value>
        public String Subject
        {
            get
            {
                if (xDocumentProperties != null)
                    return xDocumentProperties.Subject;
                return String.Empty;
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    xDocumentProperties.Subject = value;
                }
            }
        }

        /// <summary>
        /// a multi-line comment describing the document.
        /// Line delimiters can be UNIX, Macintosh or DOS style.
        /// </summary>
        /// <value>The description.</value>
        public String Description
        {
            get
            {
                if (xDocumentProperties != null)
                    return xDocumentProperties.Description;
                return String.Empty;
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    xDocumentProperties.Description = value;
                }
            }
        }

        /// <summary>
        /// a list of keywords for the document.
        /// </summary>
        /// <value>The keywords.</value>
        public List<String> Keywords
        {
            get
            {
                try
                {
                    if (xDocumentProperties != null)
                    {
                        return new List<String>(xDocumentProperties.Keywords);
                    }
                }
                catch (Exception) { }
                return new List<String>();
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    try
                    {
                        xDocumentProperties.Keywords = value.ToArray();
                    }
                    catch (Exception) { }
                }
            }
        }

        /// <summary>
        /// the default language of the document.
        /// </summary>
        /// <value>The language.</value>
        public String Language
        {
            get
            {
                try
                {
                    if (xDocumentProperties != null)
                    {
                        return xDocumentProperties.Language.Language;
                    }
                }
                catch (Exception) { }
                return String.Empty;
            }
            set
            {
                if (xDocumentProperties != null)
                {
                    try
                    {
                        xDocumentProperties.Language = new unoidl.com.sun.star.lang.Locale();
                        xDocumentProperties.Language.Language = value;
                    }
                    catch (Exception) { }
                }
            }
        }

        #endregion

        #region Custom Properties

        /// <summary>
        /// Gets all custom properties.
        /// </summary>
        /// <returns>a dictionary of names and string-values of properties</returns>
        public Dictionary<String, String> GetAllCustomProperties()
        {
            Dictionary<String, String> properties = new Dictionary<String, String>();
            if (xDocumentProperties != null)
            {
                XPropertyContainer customProperties = xDocumentProperties.getUserDefinedProperties();
                if (customProperties != null && customProperties is XPropertySet)
                {
                    if (customProperties != null && customProperties is XPropertySet)
                    {
                        XPropertySetInfo info = ((XPropertySet)customProperties).getPropertySetInfo();
                        if (info != null)
                        {
                            foreach (Property p in info.getProperties())
                            {
                                if (p != null)
                                {
                                    var val = ((XPropertySet)customProperties).getPropertyValue(p.Name);
                                    if (val.hasValue()) properties.Add(p.Name, val.Value.ToString());
                                }
                            }
                        }
                    }
                }
            }
            return properties;
        }

        /// <summary>
        /// Gets a custom property.
        /// </summary>
        /// <param name="property">The property name.</param>
        /// <returns>The property's value or <c>null</c>.</returns>
        public Object GetCustomProperty(String property)
        {
            if (!String.IsNullOrWhiteSpace(property) && xDocumentProperties != null)
            {
                XPropertyContainer customProperties = xDocumentProperties.getUserDefinedProperties();
                if (customProperties != null && customProperties is XPropertySet)
                {
                    XPropertySetInfo info = ((XPropertySet)customProperties).getPropertySetInfo();
                    if (info != null && info.hasPropertyByName(property))
                    {
                        var val = ((XPropertySet)customProperties).getPropertyValue(property);

                        if (val.hasValue())
                        {
                            // TODO: type the property
                            return val.Value;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Sets a custom property.
        /// </summary>
        /// <param name="property">The property name.</param>
        /// <param name="value">The value for the property.</param>
        /// <returns><c>true</c> if the property was successfully set, [otherwise] <c>false</c></returns>
        public bool SetCustomProperty(String property, Object value)
        {
            if (!String.IsNullOrWhiteSpace(property) && xDocumentProperties != null)
            {
                XPropertyContainer customProperties = xDocumentProperties.getUserDefinedProperties();
                if (customProperties != null && customProperties is XPropertySet)
                {
                    try
                    {
                        XPropertySetInfo info = ((XPropertySet)customProperties).getPropertySetInfo();
                        if (info != null && !info.hasPropertyByName(property))
                        {
                            customProperties.addProperty(property, 0, Any.Get(""));
                        }

                        if (info != null && info.hasPropertyByName(property))
                        {
                            ((XPropertySet)customProperties).setPropertyValue(property, Any.Get(value));
                            return true;
                        }
                    }
                    catch (Exception) { }
                }
            }
            return false;
        }

        #endregion

    }

}
