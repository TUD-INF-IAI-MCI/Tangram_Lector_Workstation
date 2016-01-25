using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using unoidl.com.sun.star.awt;
using tud.mci.tangram.util;
using unoidl.com.sun.star.beans;

namespace tud.mci.tangram.models.dialogs
{
    public class HtmlHandler
    {
        int Width { get; set; }
        int Height { get; set; }
        public ResizingContainer Container;
        HtmlFontDescriptors fonts = new HtmlFontDescriptors();


        public HtmlHandler(int width)
        {
            this.Width = width;
            Height = 0;
            Container = new ResizingContainer(Width, Height, "container____");
        }

        public void parseHtmlToDialogeTexts(String text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                string xmlWithRoot = "<xml>" + text + "</xml>";
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlWithRoot);
                if (xmlDoc != null && xmlDoc.HasChildNodes)
                {
                    XmlNodeList childs = xmlDoc.FirstChild.ChildNodes;

                    if (childs != null && childs.Count > 0)
                    {
                        foreach (XmlNode child in childs)
                        {
                            if (child != null)
                            {
                                System.Diagnostics.Debug.WriteLine("child: " + child);
                                BuildControllsFromNode(child);
                            }
                        }
                    }
                }
            }
        }


        public void BuildControllsFromNode(XmlNode node)
        {
            if (node != null)
            {
                //build control
                BuildControlFromNode(node);

                //handle children
                if (node.HasChildNodes)
                {
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        BuildControllsFromNode(child);
                    }
                }
            }
        }


        public void BuildControlFromNode(XmlNode node)
        {
            if (node != null)
            {
                if (!String.IsNullOrWhiteSpace(node.InnerXml) || !String.IsNullOrWhiteSpace(node.Value))
                {

                    switch (node.LocalName.ToLowerInvariant())
                    {
                        case "#text":
                            System.Diagnostics.Debug.WriteLine("\t\ttext to render: " + node.Value);

                            handleText(node.Value);

                            break;
                        case "p":
                            System.Diagnostics.Debug.WriteLine("\t\tparagraph to render: " + node.Value);
                            break;
                        case "h1":
                        case "h2":
                        case "h3":
                        case "h4":
                        case "h5":
                        case "h6":
                            System.Diagnostics.Debug.WriteLine("\t\tHeader to render: " + node.Value);
                            break;
                        default:
                            System.Diagnostics.Debug.WriteLine("\t\tunknown node type " + node.LocalName + " should be handled: " + node);
                            break;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("\t\t\tempty node?!: " + node.Value);
                }
            }
        }

        private void handleText(string p)
        {
            // build a label
            var label = CreateFixedLabel(p, 0, 0, 0, 0, 0, null, p.GetHashCode().ToString());
            if (Container != null && label != null)
            {
                Container.addElementToTheEndAndAdoptTheSize(label as XControl, p.GetHashCode().ToString(), 0, 2);
            }
        }

        protected virtual XFixedText CreateFixedLabel(String text, int _nPosX, int _nPosY, int _nWidth, int _nHeight, int _nStep, XMouseListener _xMouseListener, String sName = "")
        {
            XFixedText xFixedText = null;
            try
            {

                // create a controlmodel at the multiservicefactory of the dialog model...
                Object oFTModel = OO.GetMultiServiceFactory().createInstance(OO.Services.AWT_CONTROL_TEXT_FIXED_MODEL);
                Object xFTControl = OO.GetMultiServiceFactory().createInstance(OO.Services.AWT_CONTROL_TEXT_FIXED);
                XMultiPropertySet xFTModelMPSet = (XMultiPropertySet)oFTModel;
                // Set the properties at the model - keep in mind to pass the property names in alphabetical order!

                xFTModelMPSet.setPropertyValues(
                        new String[] { "Height", "Name", "PositionX", "PositionY", "Step", "Width", "Label" },
                        Any.Get(new Object[] { _nHeight, sName, _nPosX, _nPosY, _nStep, _nWidth, text }));

                if (oFTModel != null && xFTControl != null && xFTControl is XControl)
                {
                    ((XControl)xFTControl).setModel(oFTModel as XControlModel);
                }

                xFixedText = xFTControl as XFixedText;
                if (_xMouseListener != null && xFTControl is XWindow)
                {
                    XWindow xWindow = (XWindow)xFTControl;
                    xWindow.addMouseListener(_xMouseListener);
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("uno.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("System.Exception:");
                System.Diagnostics.Debug.WriteLine(ex);
            }
            return xFixedText;
        }



        #region



        #endregion


        /// <summary>
        /// Generates the basic font descriptor.
        /// Set this as property 'FontDescriptor'
        /// </summary>
        /// <returns></returns>
        public static FontDescriptor GenerateBasicFontDescriptor()
        {
            FontDescriptor basicFontDescriptor = new FontDescriptor();
            return basicFontDescriptor;
        }

    }


    public struct HtmlFontDescriptors
    {
        public FontDescriptor Normal;
        public FontDescriptor H1;
        public FontDescriptor H2;
        public FontDescriptor H3;
        public FontDescriptor H4;
        public FontDescriptor H5;
        public FontDescriptor H6;

        public HtmlFontDescriptors(int normalFontHeight = 0)
        {
            Normal = HtmlHandler.GenerateBasicFontDescriptor();
            Normal.Height = (short)normalFontHeight;
            #region Header
            H1 = Normal;
            H1.Height += 5;

            H2 = H1;
            H1.Height -= 2;

            H3 = H2;
            H1.Height -= 2;

            H4 = H3;
            H5 = H4;
            H6 = H5;
            #endregion
        }

    }

}
