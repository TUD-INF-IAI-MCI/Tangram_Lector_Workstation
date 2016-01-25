using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.container;
using uno;
using unoidl.com.sun.star.beans;

namespace tud.mci.tangram.util
{
    class DocumentListener : IDisposable, XFrameActionListener
    {
        #region Members
        private readonly Object factorykLock = new Object();
        private XMultiComponentFactory _xMcf;
        public XMultiComponentFactory XMcf
        {
            get
            {
                lock (factorykLock)
                {
                    if (_xMcf == null) _xMcf = OO.GetMultiComponentFactory(XCc);
                    return _xMcf;
                }
            }
            set
            {
                lock (factorykLock)
                {
                    _xMcf = value;
                }
            }
        }

        private readonly Object contextLock = new Object();
        private XComponentContext _xCc;
        public XComponentContext XCc
        {
            get
            {
                lock (contextLock)
                {
                    if (_xCc == null) _xCc = OO.GetContext();
                    return _xCc;
                }
            }
            set
            {
                lock (contextLock)
                {
                    _xCc = value;
                }
            }
        }

        private readonly Object deskLock = new Object();
        private XDesktop _xDesktop;
        public XDesktop XDesktop
        {
            get
            {
                lock (deskLock)
                {
                    if (_xDesktop == null) _xDesktop = OO.GetDesktop(XMcf, XCc);
                    return _xDesktop;
                }
            }
            set
            {
                lock (deskLock)
                {
                    _xDesktop = value;
                }
            }

        }

        private volatile bool serach = false;
        protected volatile int WAIT_TIME = 150;
        private Thread serachThread;

        private readonly Object dictLock = new Object();
        private Dictionary<String, XComponent> _desktopChilds = new Dictionary<String, XComponent>();
        public Dictionary<String, XComponent> DesktopDocumentComponents
        {
            get
            {
                lock (dictLock)
                {
                    return _desktopChilds;
                }
            }
            set
            {
                lock (dictLock)
                {
                    _desktopChilds = value;
                }
            }
        }



        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentListener"/> class.
        /// </summary>
        /// <param name="xMcf">The XMultiComponentFactory.</param>
        public DocumentListener(XMultiComponentFactory xMcf = null)
        {
            XMcf = xMcf;
            if (XDesktop != null)
            {
                var xF = XDesktop as XFrame;
                if (xF != null) { xF.addFrameActionListener(this); }
            }
        }

        ~DocumentListener()
        {
            Dispose();
        }

        /// <summary>
        /// Starts the search for document windows.
        /// </summary>
        public void startSearch()
        {
            serach = true;
            serachThread = serachThread != null ? serachThread : new Thread(searchForDocs);
            if (serachThread != null && !serachThread.IsAlive)
                serachThread.Start();
        }

        /// <summary>
        /// Aborts the search for documents.
        /// </summary>
        public void abortSearch()
        {
            //FIXME: for fixing
            Console.WriteLine("Abort the serach for document windows");
            serach = false;
        }

        private void searchForDocs()
        {
            while (serach)
            {
                try
                {
                    if (XDesktop != null)
                    {
                        XEnumerationAccess xEnumerationAccess = XDesktop.getComponents();
                        XEnumeration enummeraration = xEnumerationAccess.createEnumeration();

                        var _uids = DesktopDocumentComponents.Keys;

                        while (enummeraration.hasMoreElements())
                        {
                            Any anyelemet = enummeraration.nextElement();
                            XComponent element = anyelemet.Value as XComponent;
                            if (element != null)
                            {
                                //FIXME: for debugging
                                //System.Diagnostics.Debug.WriteLine("Window from Desktop found _________________");
                                //Debug.GetAllInterfacesOfObject(element);
                                //System.Diagnostics.Debug.WriteLine("\tProperties _________________");
                                //Debug.GetAllProperties(element);
                                XPropertySet ps = element as XPropertySet;
                                if (ps != null)
                                {
                                    var uid = OoUtils.GetStringProperty(ps, "RuntimeUID");
                                    if (uid != null && !uid.Equals(""))
                                    {
                                        if (!DesktopDocumentComponents.ContainsKey(uid))
                                        {
                                            //FIXME: for fixing
                                            Console.WriteLine("Found new Desktop child with uid:" + uid);
                                            DesktopDocumentComponents.Add(uid, element);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (DisposedException dex) { 
                    this.Dispose(); 
                    throw dex; 
                }
                catch (System.Exception) { }


                //Console.WriteLine("\t\tSearch for Docs");
                Thread.Sleep(WAIT_TIME);
            }
        }


        /// <summary>
        /// Führt anwendungsspezifische Aufgaben durch, die mit der Freigabe, der Zurückgabe oder dem Zurücksetzen von nicht verwalteten Ressourcen zusammenhängen.
        /// </summary>
        public void Dispose()
        {
            serach = false;
            if (serachThread != null)
                serachThread.Abort();
            _desktopChilds = null;
            _xCc = null;
            _xDesktop = null;
            _xMcf = null;
        }
        
        #region XFrameActionListener
        public void frameAction(FrameActionEvent Action)
        {
            throw new NotImplementedException();
        }

        public void disposing(EventObject Source)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
