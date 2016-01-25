using System;
using System.Collections.Generic;
using System.Text;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using tud.mci.tangram.util;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.beans;
using System.Text.RegularExpressions;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.awt;

namespace tud.mci.tangram.models.documents
{
    public class WriterDocument
    {
        private XComponent doc;
        public XComponent Document
        {
            get { return doc; }
            set { doc = value; }
        }

        private XComponentContext xContext;
        private XMultiComponentFactory _xMcFactory;
        private XMultiComponentFactory xMcFactory
        {
            get
            {
                if (_xMcFactory == null)
                {
                    if (xContext == null)
                        xContext = OO.GetContext();
                    _xMcFactory = OO.GetMultiComponentFactory(xContext);
                }
                return _xMcFactory;
            }
            set { _xMcFactory = value; }
        }

        public WriterDocument(XComponentContext xContext = null, XMultiComponentFactory xMcFactory = null)
        {
            this.xContext = xContext;
            this.xMcFactory = xMcFactory;
            doc = OOoDocument.OpenNewDocumentComponent(OO.DocTypes.SWRITER, xContext, xMcFactory);
        }

        public WriterDocument(XTextDocument xTextDocument, XComponentContext xContext = null, XMultiComponentFactory xMcFactory = null)
        {
            this.xContext = xContext;
            this.xMcFactory = xMcFactory;
            doc = xTextDocument;
        }

        public XText Text
        {
            get
            {
                if (Document != null && Document is XTextDocument)
                {
                    return ((XTextDocument)Document).getText();
                }
                return null;
            }
        }
        #region Textframe
        public XTextFrame AddTextFrame(int height, int width)
        {
            var textFrame = this.xMcFactory.createInstanceWithContext(OO.Services.TEXT_FRAME, xContext);

            if (textFrame != null)
            {
                if (textFrame is XTextFrame)
                {
                    Size size = OoUtils.MakeSize(width, height);
                    OoUtils.SetProperty(textFrame, "Size", size);
                    OoUtils.SetProperty(textFrame, "AnchorType", TextContentAnchorType.AS_CHARACTER);
                    OoUtils.SetProperty(textFrame, "LeftMargin", 0);
                    //util.Debug.GetAllProperties(textFrame);
                    XTextRange xTextRange = Text.getEnd();
                    NewParagraph();
                    Text.insertTextContent(xTextRange, textFrame as XTextFrame, false);
                    NewParagraph();
                }

            }
            return textFrame as XTextFrame;
        }
        public void AddTextToFrameAndNewParagraph(ref XTextFrame textFrame, String text, int charWeight = 100, int charHeight = 12)
        {
            if (textFrame != null)
            {
                XText xText = ((XTextFrame)textFrame).getText();
                XTextCursor xFrameTextCurser = xText.createTextCursor();
                xFrameTextCurser.gotoEnd(false);
                OoUtils.SetProperty(xFrameTextCurser, "CharWeight", charWeight);
                OoUtils.SetProperty(xFrameTextCurser, "CharHeight", charHeight);

                xText.insertString(xFrameTextCurser, text, false);
                xText.insertControlCharacter(xFrameTextCurser, ControlCharacter.PARAGRAPH_BREAK, false);
            }
        }
        /// <summary>
        /// Add bold text.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="charWeight"></param>
        /// <param name="charHeight"></param>
        public void AddText(String text, int charWeight = 100, int charHeight = 12)
        {

            if (Text != null)
            {
                XText xText = Text.getText();
                XTextCursor xFrameTextCurser = xText.createTextCursor();
                xFrameTextCurser.gotoEnd(false);
                OoUtils.SetProperty(xFrameTextCurser, "CharWeight", charWeight);
                OoUtils.SetProperty(xFrameTextCurser, "CharHeight", charHeight);

                xText.insertString(xFrameTextCurser, text, false);

            }

        }
        /// <summary>
        /// Add styled Text in the same line and add a new paragraph
        /// </summary>
        /// <param name="textFrame"></param>
        /// <param name="text"></param>
        /// <param name="charWeight"></param>
        /// <param name="charHeight"></param>
        public void AddTextAndNewParagraph(String text, int charWeight = 100, int charHeight = 12)
        {
            if (Text != null)
            {
                XText xText = Text.getText();
                XTextCursor xFrameTextCurser = xText.createTextCursor();
                xFrameTextCurser.gotoEnd(false);
                OoUtils.SetProperty(xFrameTextCurser, "CharWeight", charWeight);
                OoUtils.SetProperty(xFrameTextCurser, "CharHeight", charHeight);

                xText.insertString(xFrameTextCurser, text, false);
                xText.insertControlCharacter(xFrameTextCurser, ControlCharacter.PARAGRAPH_BREAK, false);
            }
        }

        public void AddTextToFrame(ref XTextFrame textFrame, String text, int charWeight = 100, int charHeight = 12)
        {
            if (textFrame != null)
            {
                XText xText = ((XTextFrame)textFrame).getText();
                XTextCursor xFrameTextCurser = xText.createTextCursor();
                xFrameTextCurser.gotoEnd(false);
                OoUtils.SetProperty(xFrameTextCurser, "CharWeight", charWeight);
                OoUtils.SetProperty(xFrameTextCurser, "CharHeight", charHeight);

                xText.insertString(xFrameTextCurser, text, false);

            }
        }
        public XTextTable AddTableToFrame(ref XTextFrame textFrame, int col, int rows)
        {
            var Table = this.xMcFactory.createInstanceWithContext(OO.Services.TEXT_TEXT_TABLE, xContext);

            if (Table != null)
            {
                if (Table is XTextTable)
                {
                    ((XTextTable)Table).initialize(rows, col);
                }
                XText xText = ((XTextFrame)textFrame).getText();
                XTextCursor xFrameTextCurser = xText.createTextCursor();
                xFrameTextCurser.gotoEnd(false);
                xText.insertTextContent(xFrameTextCurser, Table as XTextContent, false);
                xText.insertControlCharacter(xFrameTextCurser, ControlCharacter.PARAGRAPH_BREAK, false);
            }
            return Table as XTextTable;
        }
        #endregion
        #region text stuff

        private XTextViewCursor xTextViewCursor;
        /// <summary>
        /// Gets the text view cursor. It is the visible and actal cursor of the Document.
        /// </summary>
        public XTextViewCursor TextViewCursor
        {
            get
            {
                if (xTextViewCursor == null && Document != null)
                {
                    // the controller gives us the TextViewCursor
                    // query the viewcursor supplier interface
                    XTextViewCursorSupplier xViewCursorSupplier =
                            (XTextViewCursorSupplier)GetController((XModel)Document);
                    // get the cursor
                    xTextViewCursor = xViewCursorSupplier.getViewCursor();
                }
                return xTextViewCursor;
            }
        }

        /// <summary>
        /// Sets the style for the xTextViewCursor.
        /// </summary>
        /// <param name="paraStyleName">Name of the paragraph style. <see cref="tud.mci.tangram.models.documents.WriterDocument.ParaStyleName"/></param>
        /// <param name="charStyleName">Name of the character style. <see cref="tud.mci.tangram.models.documents.WriterDocument.CharStyleName"/></param>
        /// <returns><code>true</code> if all styles wa successfully set, otherwise <code>false</code></returns>
        public bool SetStyle(String paraStyleName = ParaStyleName.NONE, String charStyleName = CharStyleName.NONE)
        {
            bool success = false;
            try
            {
                // query its XPropertySet interface, we want to set character and paragraph properties
                XPropertySet xCursorPropertySet = (XPropertySet)TextViewCursor;

                // set the appropriate properties for character and paragraph style
                try
                {
                    xCursorPropertySet.setPropertyValue("ParaStyleName", Any.Get(paraStyleName));
                    success = true;
                }
                catch (IllegalArgumentException ex)
                {
                    System.Diagnostics.Debug.WriteLine("ParaStyleName: '" + paraStyleName + "' was not accepted.\n" + ex);
                }
                try
                {
                    xCursorPropertySet.setPropertyValue("CharStyleName", Any.Get(charStyleName));
                    success = true;
                }
                catch (IllegalArgumentException ex)
                {
                    System.Diagnostics.Debug.WriteLine("CharStyleName: '" + charStyleName + "' was not accepted.\n" + ex);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return success;
        }

        /// <summary>
        /// Append new text to the end of the document.
        /// </summary>
        /// <param name="text">The text.</param>
        public void AppendText(String text)
        {
            if (Text != null) Text.insertString(Text.getEnd(), text, false);
        }
        //public void AppendStyledText(String text,String paraStyle = WriterDocument.ParaStyleName.NONE,
        //    String charStyle = WriterDocument.CharStyleName.NONE)
        //{

        //    if (Document != null)
        //    {

        //        SetStyle(paraStyle, charStyle);
        //        AppendText(text);
        //    }
        //}
        /// <summary>
        /// Appends the text and get text range.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        public XTextRange AppendTextAndGetTextRange(String text)
        {
            if (TextViewCursor != null)
            {
                AppendText(text);
                var end = CreateTextCursor(TextViewCursor.getEnd());
                end.goLeft((short)text.Length, true);
                return end;
            }
            return null;
        }

        /// <summary>
        /// Append a new paragraph to the end of the document.
        /// </summary>
        public void NewParagraph()
        {
            AppendControllCharacter(unoidl.com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK);
        }

        /// <summary>
        /// Appends a controll character at the end of the text.
        /// </summary>
        /// <param name="cChar">The c char. From unoidl.com.sun.star.text.ControlCharacter</param>
        public void AppendControllCharacter(short cChar)
        {
            if (Text != null) Text.insertControlCharacter(Text.getEnd(), cChar, false);
        }

        /// <summary>
        /// Appends the new styled text paragraph.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="paraStyle">The para style.</param>
        /// <param name="charStyle">The char style.</param>
        public void AppendNewStyledTextParagraph(String text,
            String paraStyle = WriterDocument.ParaStyleName.NONE,
            String charStyle = WriterDocument.CharStyleName.NONE)
        {
            if (Document != null)
            {
                if (!IsLastCharParagraphBrake())
                {
                    NewParagraph();
                }
                SetStyle(paraStyle, charStyle);
                AppendText(text);
            }
        }

        /// <summary>
        /// Appends the new styled text paragraph.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="paraStyle">The para style.</param>
        /// <param name="charStyle">The char style.</param>
        public XTextRange AppendNewStyledTextParagraphAndGetTextRange(String text,
            String paraStyle = WriterDocument.ParaStyleName.NONE,
            String charStyle = WriterDocument.CharStyleName.NONE)
        {
            if (Document != null)
            {
                if (!IsLastCharParagraphBrake())
                {
                    NewParagraph();
                }
                SetStyle(paraStyle, charStyle);
                return AppendTextAndGetTextRange(text);
            }
            return null;
        }


        #region TextCursor
        /// <summary>
        /// Creates the text cursor.
        /// NOTE: You can have many non-visible Text Cursors as you want but there is only ONE View Cursor per document.
        /// </summary>
        public XTextCursor CreateTextCursorFromViewCursor()
        {
            return TextViewCursor != null ? Text.createTextCursorByRange(TextViewCursor) : null;
        }

        /// <summary>
        /// Creates the text cursor.
        /// NOTE: You can have many non-visible Text Cursors as you want but there is only ONE View Cursor per document.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns></returns>
        public XTextCursor CreateTextCursor(XTextRange range)
        {
            return Text != null ? Text.createTextCursorByRange(range) : null;
        }

        /// <summary>
        /// Creates the text cursor.
        /// NOTE: You can have many non-visible Text Cursors as you want but there is only ONE View Cursor per document.
        /// </summary>
        public XTextCursor CreateTextCursor()
        {
            return Text != null ? Text.createTextCursor() : null;
        }
        #endregion

        private Regex prargraphRegex = new Regex(@"(?:(\r\n)|\r|\n)", RegexOptions.Singleline);
        /// <summary>
        /// Determines whether [is last char paragraph brake].
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if [is last char paragraph brake]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsLastCharParagraphBrake()
        {
            TextViewCursor.gotoEnd(false);
            var tc = CreateTextCursorFromViewCursor();
            //expeand the textCursor selection
            tc.goLeft(1, true);
            if (prargraphRegex.IsMatch(tc.getString()) || IsDocumentEmpty())
            {
                return true;
            }
            return false;
        }

        #region Text ranges
        private XTextRange getTextRange(XTextRange from, XTextRange to)
        {
            if (from != null && to != null)
            {
                return setXTextRange(from, to);
            }
            return null;
        }

        #region setXTextRange
        public XTextCursor setXTextRange(XTextRange start, XTextRange end)
        {
            return setXTextRange(CreateTextCursor(), start, end);
        }
        public XTextCursor setXTextRange(XTextCursor xTextCursor, XTextRange start, XTextRange end)
        {
            if (xTextCursor != null && start != null && end != null)
            {
                xTextCursor.gotoRange(start.getStart(), false);
                xTextCursor.gotoRange(end.getEnd(), true);
            }
            return xTextCursor;
        }
        #endregion

        #endregion

        /// <summary>
        /// Determines whether [document is empty].
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if [document is empty]; otherwise, <c>false</c>.
        /// </returns>
        public bool IsDocumentEmpty()
        {
            if (Text != null && Text.getString().Length > 0)
                return false;
            return true;
        }
        #endregion

        #region Table stuff

        /// <summary>
        /// Adds a new table to the document.
        /// </summary>
        /// <param name="col">The amount of columns</param>
        /// <param name="rows">The amount of rows</param>
        /// <returns></returns>
        public XTextTable AddTable(int col, int rows)
        {
            var Table = this.xMcFactory.createInstanceWithContext(OO.Services.TEXT_TEXT_TABLE, xContext);

            if (Table != null)
            {
                if (Table is XTextTable)
                {
                    ((XTextTable)Table).initialize(rows, col);
                }
                XTextRange xTextRange = Text.getEnd();
                Text.insertTextContent(xTextRange, Table as XTextContent, false);
            }
            return Table as XTextTable;
        }

        /// <summary>
        /// Don´t useful function, because the size you have to set is a relative value
        /// </summary>
        /// <param name="table">Table you want to resize </param>
        /// <param name="size">relative size you want to set, look for the actual TableColumnSeparators you want to fit.</param>
        /// <param name="index">index of the column you want to resize. From left to right column. if you want to fit the first you have to set the index 0. The rightest column cant resize, you have to set the column before smaller </param>
        public void SetColSize(ref XTextTable table, int size, int index)
        {
            XTableColumns columns = table.getColumns();
            int colCount = columns.getCount();
            if (index > colCount)
            {
                System.Console.WriteLine("index was to big");
                return;
            }

            int iWidth = (int)OoUtils.GetIntProperty(table, "Width");

            short sTableColumnRelativeSum = (short)OoUtils.GetProperty(table, "TableColumnRelativeSum");

            double dRatio = (double)sTableColumnRelativeSum / (double)iWidth;
            double dRelativeWidth = (double)2000 * dRatio;
            double dposition = sTableColumnRelativeSum - dRelativeWidth;
            TableColumnSeparator[] a = OoUtils.GetProperty(table, "TableColumnSeparators") as TableColumnSeparator[];
            if (a != null)
            {
                System.Console.WriteLine(a[index].Position.ToString() + "was the position before");
                a[index].Position = (short)size;

                //for (int i = a.Length - 1; i >= 0; i--)
                //{
                //    a[i].Position = (short)Math.Ceiling(dposition);
                //    System.Console.WriteLine(dposition);
                //    dposition -= dRelativeWidth;

                //}
                OoUtils.SetProperty(table, "TableColumnSeparators", a);
            }

            //if (column != null)
            //{

            //    //OoUtils.SetIntProperty(column, "Width", size);
            //}

        }





        /// <summary>
        /// Adds new rows to a table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="count">The count of rows to add.</param>
        /// <param name="index">[optional] The index where to add. If &lt; 0 the row will be added at the bottom of the table. </param>
        /// <param name="countOfColumns">[optional]after a colspan you have to set the columns</param>
        /// <returns><c>true</c> if the row could been added [otherwise] <c>false</c>.</returns>
        public static bool AddRowToTable(ref XTextTable table, int count = 1, int index = -1, uint countOfColumns = 1)
        {

            if (table != null)
            {

                var rows = table.getRows();
                int rowCount = rows != null ? rows.getCount() : 0;

                count = Math.Max(0, count);
                index = index < 0 ? rowCount : index;

                if (rows != null)
                {
                    try
                    {
                        rows.insertByIndex(index, count);
                        if (countOfColumns > 1)
                        {
                            doColSplit(ref table, "A" + ++index, (uint)table.getColumns().getCount() - 1, false);
                        }
                    }
                    catch { return false; }
                }
                return true;
            }
            return false;
        }

        private static void doColSplit(ref XTextTable table, string startCellName, uint steps, bool horizontal)
        {
            if (table != null && !String.IsNullOrEmpty(startCellName) && steps > 1)
            {
                //get cursor for start cell
                var cursor = table.createCursorByCellName(startCellName);
                if (cursor != null)
                {

                    cursor.splitRange((short)steps, horizontal);
                }
            }
        }


        /// <summary>
        /// Adds new columns to a table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="count">The count of columns to add.</param>
        /// <param name="index">[optional] The index where to add. If &lt; 0 the column will be added at the right of the table.</param>
        /// <returns><c>true</c> if the column could been added [otherwise] <c>false</c>.</returns>
        public static bool AddColToTable(ref XTextTable table, int count = 1, int index = -1)
        {
            if (table != null)
            {
                var cols = table.getColumns();
                int colCount = cols != null ? cols.getCount() : 0;

                count = Math.Max(0, count);
                index = index < 0 ? colCount : index;

                if (cols != null)
                {
                    try
                    {
                        cols.insertByIndex(index, count);
                    }
                    catch { return false; }
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Merges table cols in one row to one cell.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="startCellName">Name of the start cell for the colspan. Important: the new cell will have this name.</param>
        /// <param name="steps">The steps to move right for merging e.g. 1 moves merges two cells and results in a colspan of 2.</param>
        public static void DoColSpan(ref XTextTable table, String startCellName, uint steps)
        {
            if (table != null && !String.IsNullOrEmpty(startCellName) && steps > 1)
            {
                //get cursor for start cell
                var cursor = table.createCursorByCellName(startCellName);
                if (cursor != null)
                {
                    bool succ = cursor.goRight((short)steps, true);
                    cursor.mergeRange();
                }
            }
        }


        /// <summary>
        /// Sets the text value of a cell .
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="cellName">Name of the cell.</param>
        /// <param name="value">The value.</param>
        public static void SetCellTextValue(ref XTextTable table, string cellName, string value, string fontName = "Verdana", int charWeight = 100, short VertOrient = 1, int charHeight = 12, int paraAdjust = 0)
        {
            var a1 = table.getCellByName(cellName);
            if (a1 != null && a1 is XText)
            {
                // set the style of the cell
                // set style for the TextRange before inserting text content!
                var start = ((XTextRange)a1).getStart();
                if (start != null && start is XTextRange)
                {
                    //util.Debug.GetAllProperties(start);
                    OoUtils.SetStringProperty(start, "CharFontName", fontName);
                    OoUtils.SetProperty(start, "CharWeight", (float)charWeight);
                    OoUtils.SetProperty(a1, "VertOrient", VertOrient);
                    OoUtils.SetProperty(start, "CharHeight", charHeight);
                    OoUtils.SetProperty(start, "ParaAdjust", paraAdjust);
                }
                ((XText)a1).setString(value);
            }
        }


        #endregion
        #region Lists


        #endregion
        # region bookmarks

        /// <summary>
        /// Inserts a bookmark.
        /// </summary>
        /// <param name="xTextRange">The text range the bookmark should stand for.</param>
        /// <param name="sBookName">Name of the bookmark.</param>
        public void InsertBookmark(XTextRange xTextRange, String sBookName)
        {
            if (xTextRange == null)
                return;
            if (sBookName.Equals(""))
                sBookName = FunctionHelper.GetUniqueFrameName() + "_" + DateTime.Now;

            // create a bookmark on a TextRange
            try
            {
                // the bookmark service is a context depended service, you need
                // the MultiServiceFactory from the document
                Object xObject = xMcFactory.createInstanceWithContext(OO.Services.TEXT_BOOKMARK, this.xContext);

                // set the name from the bookmark
                XNamed xNameAccess = xObject as XNamed;

                if (xNameAccess != null)
                {
                    xNameAccess.setName(sBookName);

                    // create a XTextContent, for the method 'insertTextContent'
                    XTextContent xTextContent = (XTextContent)xNameAccess;

                    // insertTextContent need a TextRange not a cursor to specify the
                    // position from the bookmark
                    ((XTextDocument)Document).getText().insertTextContent(xTextRange, xTextContent, true);

                    System.Diagnostics.Debug.WriteLine("Insert bookmark: " + sBookName);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: Insert bookmark faild because no MultiServiceFactory was created!");
                }



            }
            catch (unoidl.com.sun.star.uno.Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Error while insering bookmark '" + sBookName + "':\n" + e);
            }
        }

        /// <summary>
        /// Gets all avalible bookmarks in tzhe actual writer document.
        /// </summary>
        /// <returns>XNameAccess PropertySet or <c>NULL</c></returns>
        public XNameAccess getAllBookmarks()
        {
            XBookmarksSupplier xBookmarksSupplier = doc as XBookmarksSupplier;
            if (xBookmarksSupplier != null)
            {
                XNameAccess xNameAccess = xBookmarksSupplier.getBookmarks();
                return xNameAccess;
            }
            return null;
        }

        /// <summary>
        /// Trys to gets a bookmark by name.
        /// </summary>
        /// <param name="sBookName">Name of the bookmark.</param>
        /// <returns>Bookmark or <c>NULL</c> if no bookmark where found</returns>
        public object getBookmarkByName(string sBookName)
        {
            XNameAccess xNameAccess = getAllBookmarks();
            if (xNameAccess != null)
            {
                try
                {
                    return xNameAccess.getByName(sBookName).Value;
                }
                catch (unoidl.com.sun.star.uno.Exception e)
                {
                    System.Diagnostics.Debug.WriteLine("Error while getting bookmark '" + sBookName + "':\n" + e);
                }
            }
            return null;
        }


        # endregion

        #region Maybe allover Stuff
        /// <summary>
        /// Gets the current controller of the given model.
        /// </summary>
        /// <param name="xModel">The x model.</param>
        /// <returns>Current controller or <code>null</code></returns>
        public XController GetController(XModel xModel)
        {
            return xModel != null ? xModel.getCurrentController() : null;
        }

        #endregion

        #region Constants

        public static class CharStyleName
        {
            //TODO: extend this
            public const string NONE = "Standard";
        }

        public static class ParaStyleName
        {
            public const string NONE = "Standard";
            public const string HEADING_1 = "Heading 1";
            public const string HEADING_2 = "Heading 2";
            public const string HEADING_3 = "Heading 3";
            public const string HEADING_4 = "Heading 4";
            public const string HEADING_5 = "Heading 5";
            public const string HEADING_6 = "Heading 6";
            public const string HEADING_7 = "Heading 7";
            public const string HEADING_8 = "Heading 8";
            public const string HEADING_9 = "Heading 9";
            public const string HEADING_10 = "Heading 10";
            public const string SIGNATURE = "Signature";
            public const string LIST_INTENT = "List Indent";
            public const string SALUTATION = "Salutation";
            public const string MARGINALIA = "Marginalia";
            public const string TEXT_BODY = "Text body";
            public const string TEXT_BODY_INTENT = "Text body indent";
            public const string TEXT_BODY_FIRST_LINE_INTENT = "First line indent";
            public const string HEADING_INTENT = "Hanging indent";
            public const string HEADING = "Heading";

            //public const string NAME = "Standard";
            //public const string NAME = "Heading";
            //public const string NAME = "Text body";
            //public const string NAME = "List";
            //public const string NAME = "Caption";
            //public const string NAME = "Index";
            //public const string NAME = "First line indent";
            //public const string NAME = "Hanging indent";
            //public const string NAME = "Text body indent";
            //public const string NAME = "Salutation";
            //public const string NAME = "Signature";
            //public const string NAME = "List Indent";
            //public const string NAME = "Marginalia";
            //public const string NAME = "Heading 1";
            //public const string NAME = "Heading 2";
            //public const string NAME = "Heading 3";
            //public const string NAME = "Heading 4";
            //public const string NAME = "Heading 5";
            //public const string NAME = "Heading 6";
            //public const string NAME = "Heading 7";
            //public const string NAME = "Heading 8";
            //public const string NAME = "Heading 9";
            //public const string NAME = "Heading 10";
            public const string TITLE = "Title";
            //public const string NAME = "Subtitle";
            //public const string NAME = "Numbering 1 Start";
            //public const string NAME = "Numbering 1";
            //public const string NAME = "Numbering 1 End";
            //public const string NAME = "Numbering 1 Cont.";
            //public const string NAME = "Numbering 2 Start";
            //public const string NAME = "Numbering 2";
            //public const string NAME = "Numbering 2 End";
            //public const string NAME = "Numbering 2 Cont.";
            //public const string NAME = "Numbering 3 Start";
            //public const string NAME = "Numbering 3";
            //public const string NAME = "Numbering 3 End";
            //public const string NAME = "Numbering 3 Cont.";
            //public const string NAME = "Numbering 4 Start";
            //public const string NAME = "Numbering 4";
            //public const string NAME = "Numbering 4 End";
            //public const string NAME = "Numbering 4 Cont.";
            //public const string NAME = "Numbering 5 Start";
            //public const string NAME = "Numbering 5";
            //public const string NAME = "Numbering 5 End";
            //public const string NAME = "Numbering 5 Cont.";
            public const string LIST_1_START = "List 1 Start";
            public const string LIST_1 = "List 1";
            public const string LIST_1_END = "List 1 End";
            public const string LIST_1_CONT = "List 1 Cont.";
            //public const string NAME = "List 2 Start";
            //public const string NAME = "List 2";
            //public const string NAME = "List 2 End";
            //public const string NAME = "List 2 Cont.";
            //public const string NAME = "List 3 Start";
            //public const string NAME = "List 3";
            //public const string NAME = "List 3 End";
            //public const string NAME = "List 3 Cont.";
            //public const string NAME = "List 4 Start";
            //public const string NAME = "List 4";
            //public const string NAME = "List 4 End";
            //public const string NAME = "List 4 Cont.";
            //public const string NAME = "List 5 Start";
            //public const string NAME = "List 5";
            //public const string NAME = "List 5 End";
            //public const string NAME = "List 5 Cont.";
            //public const string NAME = "Index Heading";
            //public const string NAME = "Index 1";
            //public const string NAME = "Index 2";
            //public const string NAME = "Index 3";
            //public const string NAME = "Index Separator";
            //public const string NAME = "Contents Heading";
            //public const string NAME = "Contents 1";
            //public const string NAME = "Contents 2";
            //public const string NAME = "Contents 3";
            //public const string NAME = "Contents 4";
            //public const string NAME = "Contents 5";
            //public const string NAME = "User Index Heading";
            //public const string NAME = "User Index 1";
            //public const string NAME = "User Index 2";
            //public const string NAME = "User Index 3";
            //public const string NAME = "User Index 4";
            //public const string NAME = "User Index 5";
            //public const string NAME = "Contents 6";
            //public const string NAME = "Contents 7";
            //public const string NAME = "Contents 8";
            //public const string NAME = "Contents 9";
            //public const string NAME = "Contents 10";
            //public const string NAME = "Illustration Index Heading";
            //public const string NAME = "Illustration Index 1";
            //public const string NAME = "Object index heading";
            //public const string NAME = "Object index 1";
            //public const string NAME = "Table index heading";
            //public const string NAME = "Table index 1";
            //public const string NAME = "Bibliography Heading";
            //public const string NAME = "Bibliography 1";
            //public const string NAME = "User Index 6";
            //public const string NAME = "User Index 7";
            //public const string NAME = "User Index 8";
            //public const string NAME = "User Index 9";
            //public const string NAME = "User Index 10";
            //public const string NAME = "Header";
            //public const string NAME = "Header left";
            //public const string NAME = "Header right";
            //public const string NAME = "Footer";
            //public const string NAME = "Footer left";
            //public const string NAME = "Footer right";
            //public const string NAME = "Table Contents";
            public const string TABLE_HEADING = "Table Heading";
            //public const string NAME = "Illustration";
            //public const string NAME = "Table";
            //public const string NAME = "Text";
            public const string FRAME_CONTENT = "Frame contents";
            //public const string NAME = "Footnote";
            //public const string NAME = "Addressee";
            //public const string NAME = "Sender";
            //public const string NAME = "Endnote";
            //public const string NAME = "Drawing";
            //public const string NAME = "Quotations";
            //public const string NAME = "Preformatted Text";
            //public const string NAME = "Horizontal Line";
            //public const string NAME = "List Contents";
            //public const string NAME = "List Heading";





        }

        #endregion



    }
}
