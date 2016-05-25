using System;
using System.Collections.Generic;
using System.Text;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.drawing;
using tud.mci.tangram.util;
using unoidl.com.sun.star.container;
using System.Collections;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.beans;
using tud.mci.tangram.models.documents;
using tud.mci.tangram.models;
using unoidl.com.sun.star.text;

namespace tud.mci.tangram.classes
{
    /// <summary>
    ///     Class for crabbing all set titles and descriptions in a daw document 
    /// </summary>
    class DescriptionMapper
    {

        private XComponentContext _xContext;
        public XComponentContext xContext
        {
            get
            {
                if (_xContext == null)
                    xContext = OO.GetContext();
                return _xContext;
            }
            set { _xContext = value; }
        }

        private XDrawPagesSupplier _xDps;
        public XDrawPagesSupplier xDrawPagesSupplier
        {
            get { return _xDps; }
            set { _xDps = value; }
        }

        private WriterDocument _writerDoc;
        public WriterDocument WriterDoc
        {
            get
            {
                if (_writerDoc == null)
                    WriterDoc = new WriterDocument(xContext);
                return _writerDoc;
            }
            set { _writerDoc = value; }
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="DescriptionMapper"/> class.
        /// </summary>
        /// <param name="xDrawPagesSupplier">The x draw pages supplier.</param>
        /// <param name="xContext">The x context. (optional)</param>
        public DescriptionMapper(XDrawPagesSupplier xDrawPagesSupplier, XComponentContext xContext = null)
        {
            if (xContext != null)
                this.xContext = xContext;
            this.xDrawPagesSupplier = xDrawPagesSupplier;
        }


        public List<Object> GetAllTitleAndDescriptions()
        {

            var textTobjects = new List<Object>();

            var pages = GetAllDrawPages();
            foreach (var page in pages)
            {
                //foreach (var obj in GetAllDrawShapesOnPage(page))
                //{
                //    textTobjects = textTobjects.Union(getTextObjectsFromComponet(obj as XComponent)).ToList();
                //}
            }

            //open new Textdocument
            addToWriterDoc(textTobjects);
            return textTobjects;
        }

        private bool addToWriterDoc(List<Object> items)
        {
            foreach (var item in items)
            {
                var ps = item as XPropertySet;
                if (ps != null)
                {
                    appendToWriterDoc(ps);
                }
            }
            return false;
        }


        private void appendToWriterDoc(XPropertySet ps)
        {
            if (WriterDoc != null && ps != null)
            {
                var trT = AppendNonEmptyTrimedStyledParagraphToWriterDoc(ps.getPropertyValue("Title").Value.ToString(), WriterDocument.ParaStyleName.HEADING_1);
                var trD = AppendNonEmptyTrimedStyledParagraphToWriterDoc(ps.getPropertyValue("Description").Value.ToString());
                var trN = AppendNonEmptyTrimedStyledParagraphToWriterDoc(ps.getPropertyValue("Name").Value.ToString(), WriterDocument.ParaStyleName.MARGINALIA);

                //add bookmarks
                WriterDoc.InsertBookmark(trT, ps.GetHashCode() + "_Title");
                WriterDoc.InsertBookmark(trD, ps.GetHashCode() + "_Desc");
                WriterDoc.InsertBookmark(trN, ps.GetHashCode() + "_Name");

            }
        }

        private XTextRange AppendNonEmptyTrimedStyledParagraphToWriterDoc(String text,
            String paraStyle = WriterDocument.ParaStyleName.NONE,
            String charStyle = WriterDocument.CharStyleName.NONE)
        {
            var t = text.Trim();
            if (!t.Equals(""))
            {
                var tr = WriterDoc.AppendNewStyledTextParagraphAndGetTextRange(t, paraStyle, charStyle);

                registerListener(tr);
                return tr;

            }
            return null;
        }






        private List<Object> getTextObjectsFromComponet(XComponent comp)
        {
            var objList = new List<Object>();

            if (comp != null)
            {
                if (HasTitleAndDescription(comp as XPropertySet))
                {
                    objList.Add(comp);
                }

                if (comp is XIndexAccess)
                {
                    //foreach (var child in GetAllObjectsFromXIndexAccess(comp as XIndexAccess))
                    //{
                    //    objList = objList.Union(getTextObjectsFromComponet(child as XComponent)).ToList();
                    //}
                }
            }

            return objList;
        }

        public static bool HasTitleAndDescription(XPropertySet properties)
        {
            if (properties != null
                && properties.getPropertyValue("Name").Value != null
                && properties.getPropertyValue("Title").Value != null
                && properties.getPropertyValue("Description").Value != null)
            {
                return true;
            }
            return false;
        }

        #region Stuff

        #region GetAllObjectsFromXIndexAccess
        private List<XDrawPage> GetAllDrawPages()
        {
            return GetAllObjectsFromXIndexAccess<XDrawPage>(xDrawPagesSupplier != null ? xDrawPagesSupplier.getDrawPages() as XIndexAccess : null) as List<XDrawPage>;
        }

        private List<Object> GetAllDrawShapesOnPage(XDrawPage xDrawPage)
        {
            return GetAllObjectsFromXIndexAccess(xDrawPage as XIndexAccess) as List<Object>;
        }

        private IList<Object> GetAllObjectsFromXIndexAccess(XIndexAccess container)
        {
            return GetAllObjectsFromXIndexAccess<Object>(container);
        }

        private IList<T> GetAllObjectsFromXIndexAccess<T>(XIndexAccess container) where T : class
        {
            IList<T> list = new List<T>();
            if (container != null && container.hasElements())
            {
                for (int i = 0; i < container.getCount(); i++)
                {
                    var child = container.getByIndex(i).Value;
                    if (child != null && (child is T))
                        list.Add(child as T);
                }
            }
            return list;
        }
        #endregion

        #endregion

        #region debug and testing
        private void TestTextRangeEqualsText(XTextRange textRange, String text)
        {
            //if (!text.Equals(textRange.getString()))
            //    System.Diagnostics.Debug.WriteLine("WARNING: TextRange does not equals text");
        }

        private void registerListener(XTextRange tr)
        {
            //if (tr != null)
            //{
            //    //starterCursor
            //    var start = WriterDoc.CreateTextCursor(tr.getStart());

            //    //endCursor
            //    var end = WriterDoc.CreateTextCursor(tr.getEnd());

            //    //Debug.GetAllInterfacesOfObject(start);
            //    //Debug.GetAllProperties(start);


            //    /* XPropertySet:
            //     * addPropertyChangeListener
            //     * addVetoableChangeListener
            //    */

            //    ((XPropertySet)start).addPropertyChangeListener("", new TextRangeListener());

            //}

        }


        #endregion

    }


    public class TextRangeListener : XPropertyChangeListener
    {
        public TextRangeListener()
        {

        }

        public void propertyChange(PropertyChangeEvent evt)
        {
            throw new NotImplementedException();
        }

        public void disposing(EventObject Source)
        {
            throw new NotImplementedException();
        }
    }
}