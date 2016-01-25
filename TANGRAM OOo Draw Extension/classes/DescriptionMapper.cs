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



    //public class TextRangeObservation
    //{
    //    public XTextRange TextRange { get; set; }
    //    public XTextCursor StartTextCursor { get; set; }
    //    public XTextCursor EndTextCursor { get; set; }
    //    //public WriterDocument WriterDoc {get; set;}
    //    public int MyProperty { get; set; }
    //    public string OriginalString { get; set; }


    //    public TextRangeObservation(XTextRange xTextRange, WriterDocument writerDoc)
    //    {
    //        TextRange = xTextRange;
    //        var WriterDoc = writerDoc;
    //        OriginalString = TextRange != null ? TextRange.getString() : String.Empty;

    //        //starterCursor
    //        StartTextCursor = (WriterDoc != null && TextRange != null) ? WriterDoc.CreateTextCursor(TextRange.getStart()) : null;
    //        //endCursor
    //        EndTextCursor = (WriterDoc != null && TextRange != null) ? WriterDoc.CreateTextCursor(TextRange.getEnd()) : null;

    //    }
    //}


}



/*************************************************************************
 *
 *  The Contents of this file are made available subject to the terms of
 *  the BSD license.
 *
 *  Copyright 2000, 2010 Oracle and/or its affiliates.
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions
 *  are met:
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *  3. Neither the name of Sun Microsystems, Inc. nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 *  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 *  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 *  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 *  COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 *  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 *  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS
 *  OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 *  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 *  USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 *************************************************************************/

//***************************************************************************
// comment: Step 1: get the Desktop object from the office
//          Step 2: open an empty text document
//          Step 3: enter a example text
//          Step 4: insert some bookmarks
//
//          Chapter 5.1.1.4 Inserting bookmarks
//***************************************************************************

//import com.sun.star.uno.UnoRuntime;

//public class BookmarkInsertion {

//    public static void main(String args[]) {
//        // You need the desktop to create a document
//        // The getDesktop method does the UNO bootstrapping, gets the
//        // remote servie manager and the desktop object.
//        com.sun.star.frame.XDesktop xDesktop = null;
//        xDesktop = getDesktop();

//        // create text document
//        com.sun.star.text.XTextDocument xTextDocument = null;
//        xTextDocument = createTextdocument(xDesktop);

//        // put example text in document
//        createExampleData(xTextDocument);


//        String mOffending[] = { "negro(e|es)?","bor(ed|ing)?",
//                                "bloody?", "bleed(ing)?" };
//        String mBad[] = { "possib(le|ilit(y|ies))", "real(ly)+", "brilliant" };

//        String sOffendPrefix = "Offending";
//        String sBadPrefix = "BadStyle";

//        markList(xTextDocument, mOffending, sOffendPrefix);
//        markList(xTextDocument, mBad, sBadPrefix);

//        System.out.println("Done");

//        System.exit(0);
//    }

//    public static void markList(com.sun.star.text.XTextDocument xTextDocument,
//                                String mList[], String sPrefix) {
//        int iCounter=0;
//        com.sun.star.uno.XInterface xSearchInterface = null;
//        com.sun.star.text.XTextRange xSearchTextRange = null;

//        try {
//            for( iCounter = 0; iCounter < mList.length; iCounter++ ) {
//                // the findfirst returns a XInterface
//                xSearchInterface = (com.sun.star.uno.XInterface)FindFirst(
//                    xTextDocument, mList[ iCounter ] );

//                if( xSearchInterface != null ) {
//                    // get the TextRange form the XInterface
//                    xSearchTextRange = (com.sun.star.text.XTextRange)
//                        UnoRuntime.queryInterface(
//                            com.sun.star.text.XTextRange.class, xSearchInterface);

//                    InsertBookmark(xTextDocument, xSearchTextRange,
//                                   sPrefix + iCounter);
//                }
//            }
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//            System.exit(1);
//        }
//    }


//    public static void InsertBookmark(com.sun.star.text.XTextDocument xTextDocument,
//                                      com.sun.star.text.XTextRange xTextRange,
//                                      String sBookName) {
//        // create a bookmark on a TextRange
//        try {
//            // get the MultiServiceFactory from the text document
//            com.sun.star.lang.XMultiServiceFactory xDocMSF;
//            xDocMSF = (com.sun.star.lang.XMultiServiceFactory)
//                UnoRuntime.queryInterface(
//                    com.sun.star.lang.XMultiServiceFactory.class, xTextDocument);

//            // the bookmark service is a context dependend service, you need
//            // the MultiServiceFactory from the document
//            Object xObject = xDocMSF.createInstance("com.sun.star.text.Bookmark");

//            // set the name from the bookmark
//            com.sun.star.container.XNamed xNameAccess = null;
//            xNameAccess = (com.sun.star.container.XNamed)
//                UnoRuntime.queryInterface(
//                    com.sun.star.container.XNamed.class, xObject);

//            xNameAccess.setName(sBookName);

//            // create a XTextContent, for the method 'insertTextContent'
//            com.sun.star.text.XTextContent xTextContent = null;
//            xTextContent = (com.sun.star.text.XTextContent)
//                UnoRuntime.queryInterface(
//                    com.sun.star.text.XTextContent.class, xNameAccess);

//            // insertTextContent need a TextRange not a cursor to specify the
//            // position from the bookmark
//            xTextDocument.getText().insertTextContent(xTextRange, xTextContent, true);

//            System.out.println("Insert bookmark: " + sBookName);
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//        }
//    }

//    protected static com.sun.star.uno.XInterface FindFirst(
//        com.sun.star.text.XTextDocument xTextDocument, String sSearchString)
//    {
//        com.sun.star.util.XSearchDescriptor xSearchDescriptor = null;
//        com.sun.star.util.XSearchable xSearchable = null;
//        com.sun.star.uno.XInterface xSearchInterface = null;

//        try {
//            xSearchable = (com.sun.star.util.XSearchable)
//                UnoRuntime.queryInterface(
//                    com.sun.star.util.XSearchable.class, xTextDocument);
//            xSearchDescriptor = (com.sun.star.util.XSearchDescriptor)
//                xSearchable.createSearchDescriptor();

//            xSearchDescriptor.setSearchString(sSearchString);

//            com.sun.star.beans.XPropertySet xPropertySet = null;
//            xPropertySet = (com.sun.star.beans.XPropertySet)
//                UnoRuntime.queryInterface(
//                    com.sun.star.beans.XPropertySet.class, xSearchDescriptor);

//            xPropertySet.setPropertyValue("SearchRegularExpression",
//                                          new Boolean( true ) );

//            xSearchInterface = (com.sun.star.uno.XInterface)
//                xSearchable.findFirst(xSearchDescriptor);
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//        }

//        return xSearchInterface;
//    }

//    protected static void createExampleData(
//        com.sun.star.text.XTextDocument xTextDocument )
//    {
//        com.sun.star.text.XTextCursor xTextCursor = null;

//        try {
//            xTextCursor = (com.sun.star.text.XTextCursor)
//                xTextDocument.getText().createTextCursor();

//            xTextCursor.setString( "He heard quiet steps behind him. That didn't bode well. Who could be following him this late at night and in this deadbeat part of town? And at this particular moment, just after he pulled off the big time and was making off with the greenbacks. Was there another crook who'd had the same idea, and was now watching him and waiting for a chance to grab the fruit of his labor?" );
//            xTextCursor.collapseToEnd();
//            xTextCursor.setString( "Or did the steps behind him mean that one of many bloody officers in town was on to him and just waiting to pounce and snap those cuffs on his wrists? He nervously looked all around. Suddenly he saw the alley. Like lightening he darted off to the left and disappeared between the two warehouses almost falling over the trash can lying in the middle of the sidewalk. He tried to nervously tap his way along in the inky darkness and suddenly stiffened: it was a dead-end, he would have to go back the way he had come" );
//            xTextCursor.collapseToEnd();
//            xTextCursor.setString( "The steps got louder and louder, he saw the black outline of a figure coming around the corner. Is this the end of the line? he thought pressing himself back against the wall trying to make himself invisible in the dark, was all that planning and energy wasted? He was dripping with sweat now, cold and wet, he could smell the brilliant fear coming off his clothes. Suddenly next to him, with a barely noticeable squeak, a door swung quietly to and fro in the night's breeze." );

//            xTextCursor.gotoStart(false);
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//        }

//    }

//    public static com.sun.star.frame.XDesktop getDesktop() {
//        com.sun.star.frame.XDesktop xDesktop = null;
//        com.sun.star.lang.XMultiComponentFactory xMCF = null;

//        try {
//            com.sun.star.uno.XComponentContext xContext = null;

//            // get the remote office component context
//            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();

//            // get the remote office service manager
//            xMCF = xContext.getServiceManager();
//            if( xMCF != null ) {
//                System.out.println("Connected to a running office ...");

//                Object oDesktop = xMCF.createInstanceWithContext(
//                    "com.sun.star.frame.Desktop", xContext);
//                xDesktop = (com.sun.star.frame.XDesktop) UnoRuntime.queryInterface(
//                    com.sun.star.frame.XDesktop.class, oDesktop);
//            }
//            else
//                System.out.println( "Can't create a desktop. No connection, no remote office servicemanager available!" );
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//            System.exit(1);
//        }


//        return xDesktop;
//    }

//    public static com.sun.star.text.XTextDocument createTextdocument(
//        com.sun.star.frame.XDesktop xDesktop )
//    {
//        com.sun.star.text.XTextDocument aTextDocument = null;

//        try {
//            com.sun.star.lang.XComponent xComponent = CreateNewDocument(xDesktop,
//                                                                        "swriter");
//            aTextDocument = (com.sun.star.text.XTextDocument)
//                UnoRuntime.queryInterface(
//                    com.sun.star.text.XTextDocument.class, xComponent);
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.err);
//        }

//        return aTextDocument;
//    }


//    protected static com.sun.star.lang.XComponent CreateNewDocument(
//        com.sun.star.frame.XDesktop xDesktop,
//        String sDocumentType )
//    {
//        String sURL = "private:factory/" + sDocumentType;

//        com.sun.star.lang.XComponent xComponent = null;
//        com.sun.star.frame.XComponentLoader xComponentLoader = null;
//        com.sun.star.beans.PropertyValue xValues[] =
//            new com.sun.star.beans.PropertyValue[1];
//        com.sun.star.beans.PropertyValue xEmptyArgs[] =
//            new com.sun.star.beans.PropertyValue[0];

//        try {
//            xComponentLoader = (com.sun.star.frame.XComponentLoader)
//                UnoRuntime.queryInterface(
//                    com.sun.star.frame.XComponentLoader.class, xDesktop);

//            xComponent  = xComponentLoader.loadComponentFromURL(
//                sURL, "_blank", 0, xEmptyArgs);
//        }
//        catch( Exception e) {
//            e.printStackTrace(System.out);
//        }

//        return xComponent ;
//    }
//}

