using System;
using System.Collections.Generic;
using System.Text;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using uno.util;

namespace tud.mci.tangram.util
{

    //xURLTransformer->parseStrict( aURL );

    //aArgs[0].Name = rtl::OUString( RTL_CONSTASCII_USTRINGPARAM( "FileName" );

    //aArgs[0].Value = uno::makeAny( rtl::OUString( RTL_CONSTASCII_USTRINGPARAM( "file:///home/test/test.odt" )) );

    //xDispatch->dispatch( aURL, aArgs )

    //There is also a way to pass simple types via the URL. Keep in mind that you have to escape certain characters (e.g. ?,&,' ',...) if you use them in your arguments.


    //command?arg1:type1=value2&arg2:type2=value2

    //".uno:Open?FileName:string=file:///home/test/test.odt"

    public static class CommandUrlHelper
    {
        public static URL getCommandUrl(string protocol, string path)
        {
            var aUrl = new URL();
            aUrl.Protocol = protocol;
            aUrl.Path = path;
            aUrl.Complete = (protocol.EndsWith(":") ? protocol : protocol + ":") + path;
            return aUrl;
        }

        public static URL getCommandUrl(string complete)
        {
            var aUrl = new URL();
            aUrl.Complete = complete;
            return aUrl;
        }


        public static URL getCommandUrl(string protocol, string path, List<PropertyValue> arguments)
        {
            var aUrl = getCommandUrl(protocol, path);
            return appendArgumentsToURL(aUrl, getPropertyValueArgumentString(arguments));
        }

        public static URL getCommandUrl(string complete, List<PropertyValue> arguments)
        {
            var aUrl = getCommandUrl(complete);
            return appendArgumentsToURL(aUrl, getPropertyValueArgumentString(arguments));
        }

        private static URL appendArgumentsToURL(URL aURL, string arguments)
        {
            aURL.Arguments = arguments;
            aURL.Complete += (aURL.Complete.EndsWith("?") ? arguments : "?" + arguments);
            //aURL = transformUrl(aURL);
            return aURL;
        }

        #region Arguments

        //There is also a way to pass simple types via the URL. Keep in mind that you have to escape certain characters (e.g. ?,&,' ',...) if you use them in your arguments.
        // command?arg1:type1=value2&arg2:type2=value2
        // service:MyAddon.MyService?arg1=blah&arg2=123


        public static string getPropertyValueArgumentString(List<PropertyValue> arguments)
        {
            string result = "";
            foreach (var argument in arguments)
            {
                string aS = getPropertyValueArgumentString(argument).Trim();
                result += aS.Equals("") ? "" : (result.Length > 1 ? "&" + aS : aS);
            }
            return result;
        }

        public static string getPropertyValueArgumentString(PropertyValue argument)
        {
            string result = "";
            if (argument != null)
            {
                //arg1:type1=value
                if (argument.Name != null && !argument.Name.Equals(""))
                {
                    result += argument.Name 
                        //+ ":" + argument.Value.Type.ToString().Trim() 
                        + "=" + getEncodedPropertyValue(argument.Value.Value.ToString());
                }

            }
            return result;
        }

        public static string getEncodedPropertyValue(string value)
        {
            //TODO: do this
            return value.Trim();
        }
        

        public static string getPathFromCommandUrl(string url)
        {
            string _url = String.Empty;
            int iqs = url.IndexOf('?');

            if (iqs == -1) { _url = url; }
            // If query string variables exist, put them in a string. 
            else if (iqs >= 0)
            {
                _url = url.Substring(0, url.IndexOf("?"));
            }
            return _url;
        }

        public static Dictionary<string, string> getParameterFromCommandUrl(string url)
        {
            Dictionary<string, string> ps = new Dictionary<string, string>();
            int iqs = url.IndexOf('?');

            if (iqs == -1) { }
            // If query string variables exist, put them in a string. 
            else if (iqs >= 0)
            {
                var querystring = (iqs < url.Length - 1) ? url.Substring(iqs + 1) : String.Empty;
                var querrys = querystring.Split("&".ToCharArray());

                foreach (var querry in querrys)
                {
                    var q = querry.Split("=".ToCharArray());
                    if (q.Length > 1) { ps.Add(q[0], q[1]); }
                    else if (q.Length > 0) { ps.Add(q[0], null); }
                }

            }
            return ps;
        }




        //public static URL transformUrl(URL aUrl)
        //{
        //    try
        //    {
        //        //XComponentContext xOfficeCtx = OOo.GetContext();


        //        XComponentContext xOfficeCtx = Bootstrap.defaultBootstrap_InitialComponentContext();

        //        // Create special service for parsing of given URL.
        //        var xParser = (XURLTransformer)xOfficeCtx.getServiceManager().createInstanceWithContext(
        //                    OOo.Services.UTIL_URL_TRANSFORMER, xOfficeCtx);

        //        if (xParser != null)
        //        {
        //            if (!xParser.assemble(ref aUrl))
        //                System.Diagnostics.Debug.WriteLine("URL can't be assembled");
        //        }
        //    }
        //    catch (unoidl.com.sun.star.uno.Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine("Error while try to assemble URL\n" + ex);
        //    }

        //    return aUrl;
        //}
        #endregion
    }
}
