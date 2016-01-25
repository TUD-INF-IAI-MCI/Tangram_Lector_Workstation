using System;
using System.Collections.Generic;
using System.Threading;
using tud.mci.tangram.util;
using unoidl.com.sun.star.accessibility;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.drawing;
using unoidl.com.sun.star.lang;

namespace tud.mci.tangram.Accessibility
{
    public static class OoAccessibility
    {
        /// <summary>
        /// Gets the accessible role of the XAccessible obj if possible.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>AccessibleRole or AccessibleRole.UNKNOWN</returns>
        public static AccessibleRole GetAccessibleRole(XAccessible obj)
        {
            AccessibleRole r = AccessibleRole.UNKNOWN;

            if (obj != null)
            {
                TimeLimitExecutor.WaitForExecuteWithTimeLimit(400, () =>
                {
                    try
                    {
                        r = GetAccessibleRole(obj.getAccessibleContext());
                    }
                    catch (unoidl.com.sun.star.lang.DisposedException)
                    {
                        r = AccessibleRole.DISPOSED;
                    }
                    catch (ThreadAbortException) { }
                    catch (ThreadInterruptedException) { }
                    catch { r = AccessibleRole.INVALID; }
                }, "");
            }
            return AccessibleRole.UNKNOWN;
        }

        const int maxTrys = 3;

        /// <summary>
        /// Gets the accessible role of the XAccessible obj if possible.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>AccessibleRole or AccessibleRole.UNKNOWN</returns>
        public static AccessibleRole GetAccessibleRole(XAccessibleContext obj)
        {
            if (obj != null)
            {
                AccessibleRole r = AccessibleRole.UNKNOWN;
                try
                {
                    r = tryGetAccessibleRole(obj);
                    return r;
                }
                catch (unoidl.com.sun.star.lang.DisposedException)
                {
                    return AccessibleRole.INVALID;
                    //Logger.Instance.Log(LogPriority.DEBUG, "can't get accessibility role - the object is already disposed", ex); 
                }

                catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, "can't get accessibility role", ex); }
            }
            return AccessibleRole.UNKNOWN;
        }

        private static AccessibleRole tryGetAccessibleRole(XAccessibleContext obj)
        {
            AccessibleRole r = AccessibleRole.UNKNOWN;
            try
            {
                TimeLimitExecutor.WaitForExecuteWithTimeLimit(100,
                             () =>
                             {
                                 try
                                 {
                                     short role = obj.getAccessibleRole();
                                     r = GetRoleFromShort(role);
                                 }
                                 catch (unoidl.com.sun.star.lang.DisposedException)
                                 {
                                     r = AccessibleRole.INVALID;
                                 }
                                 catch (ThreadAbortException) { r = AccessibleRole.UNKNOWN; }
                             }, "GetAccRole4XAccCont");
            }
            catch (System.Threading.ThreadAbortException) { }
            return r;
        }


        /// <summary>
        /// Gets the name of the accessible.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        public static String GetAccessibleName(XAccessible obj) { return obj != null ? GetAccessibleName(obj.getAccessibleContext()) : String.Empty; }
        /// <summary>
        /// Gets the name of the accessible.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        public static String GetAccessibleName(XAccessibleContext obj)
        {
            try
            {
                string name = obj != null ? obj.getAccessibleName() : String.Empty;
                return name;
            }
            catch (System.Exception ex)
            {
                try
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "Could not get the Accessible name for object: " + ex);
                    //System.Diagnostics.Debug.WriteLine("------------- ERROR FOR ELEMENT ---------------");
                    //util.Debug.GetAllServicesOfObject(obj);
                    //util.Debug.GetAllInterfacesOfObject(obj);
                    //System.Diagnostics.Debug.WriteLine("------------- END OF ERROR MESSAGE ---------------");
                }
                catch { }
                return String.Empty;
            }
        }

        /// <summary>
        /// Gets the 'Name' part of the accessible name.
        /// The accessible name is a combination of 'Name' + " " + 'Title' [or Text] + [...]
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>The first part of the accessible name up to the first appearence of an space character</returns>
        public static String GetAccessibleNamePart(XAccessible obj) { return obj != null ? GetAccessibleNamePart(obj.getAccessibleContext()) : String.Empty; }

        /// <summary>
        /// Gets the 'Name' part of the accessible name.
        /// The accessible name is a combination of 'Name' + " " + 'Title' [or Text] + [...]
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>The first part of the accessible name up to the first appearence of an space character</returns>
        public static String GetAccessibleNamePart(XAccessibleContext obj)
        {
            String name = String.Empty;
            String accName = GetAccessibleName(obj);
            if (!String.IsNullOrWhiteSpace(accName))
            {
                int deviderPos = accName.IndexOf(' ');
                if (deviderPos > 0)
                {
                    name = accName.Substring(0, deviderPos);
                }
                else
                {
                    name = accName;
                }
            }
            return name;
        }

        /// <summary>
        /// Gets the accessible description.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        public static String GetAccessibleDesc(XAccessible obj) { return obj != null ? GetAccessibleDesc(obj.getAccessibleContext()) : String.Empty; }
        /// <summary>
        /// Gets the accessible description.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        public static String GetAccessibleDesc(XAccessibleContext obj) { return obj != null ? obj.getAccessibleDescription() : String.Empty; }

        /// <summary>
        /// Prints some accessible infos into the Debug Output.
        /// </summary>
        /// <param name="obj">The obj.</param>
        public static String PrintAccessibleInfos(XAccessibleContext accessibleContext, bool printStates = false)
        {
            String output = "";
            if (accessibleContext != null)
            {
                var role = GetRoleFromShort(accessibleContext.getAccessibleRole());

                var name = accessibleContext.getAccessibleName();

                var desc = accessibleContext.getAccessibleDescription();
                output += "Role: " + role + " | Name: '" + name + "' Description: '" + desc + "'";

                var child = accessibleContext.getAccessibleChildCount();
                output += " (has " + child + " children)";

                if (printStates)
                {
                    var states = GetAccessibleStates(accessibleContext.getAccessibleStateSet());
                    foreach (var item in states)
                    {
                        output += "\r\nState: " + item;
                    }
                }
            }
            return output;
        }
        /// <summary>
        /// Prints some accessible infos into the Debug Output.
        /// </summary>
        /// <param name="obj">The obj.</param>
        public static String PrintAccessibleInfos(XAccessible obj, bool printStates = false)
        {
            if (obj != null)
            {
                XAccessibleContext accessibleContext = obj.getAccessibleContext();
                return PrintAccessibleInfos(accessibleContext, printStates);
            }
            return String.Empty;
        }

        /// <summary>
        /// Converts a short (index) into an AccessibleRole.
        /// </summary>
        /// <param name="role">The role.</param>
        /// <returns>The corresponding AccessibleRole to the role.</returns>
        public static AccessibleRole GetRoleFromShort(short role)
        {
            try { return (AccessibleRole)Enum.ToObject(typeof(AccessibleRole), role); }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "can't cast short '" + role + "' to accessibility role", ex);
                return AccessibleRole.UNKNOWN;
            }
        }

        /// <summary>
        /// Converts a short (index) into an AccessibleEventId.
        /// </summary>
        /// <param name="id">The id.</param>
        /// <returns>The corresponding AccessibleEventId to the id.</returns>
        public static AccessibleEventId GetAccessibleEventIdFromShort(short id)
        {
            try { return (AccessibleEventId)Enum.ToObject(typeof(AccessibleEventId), id); }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "can't cast short '" + id + "' to AccessibleEventId", ex);
                return AccessibleEventId.NONE;
            }
        }

        /// <summary>
        /// Converts a short (index) into an AccessibleRelationType.
        /// </summary>
        /// <param name="relation">The relation.</param>
        /// <returns>The corresponding AccessibleRelationType to the relation id.</returns>
        public static AccessibleRelationType GetAccessibleRelationTypeFromShort(short relation)
        {
            try { return (AccessibleRelationType)Enum.ToObject(typeof(AccessibleRelationType), relation); }
            catch (Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "can't cast short '" + relation + "' to AccessibleRelationType", ex);
                return AccessibleRelationType.INVALID;
            }
        }

        /// <summary>
        /// Converts the accessible states of the element into a list of AccessibleStateType.
        /// </summary>
        /// <param name="states">The XAccessibleStateSet of an element returned by XAccessibleContext.getAccessibleStateSet().</param>
        /// <returns>A list of AccessibleStateType of the element.</returns>
        public static List<AccessibleStateType> GetAccessibleStates(XAccessibleStateSet states)
        {
            List<AccessibleStateType> s = new List<AccessibleStateType>(1);
            if (states != null && !states.isEmpty())
            {
                foreach (short item in states.getStates())
                {
                    try { s.Add(getAccessibleStateTypeFromShort(item)); }
                    catch { }

                }
                return s;
            }
            else s.Add(AccessibleStateType.INVALID);
            return s;
        }

        private static AccessibleStateType getAccessibleStateTypeFromShort(short item)
        {
            try
            {
                return (AccessibleStateType)Enum.ToObject(typeof(AccessibleStateType), item);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "can't cast short '" + item + "' to AccessibleStateType", ex);
                return AccessibleStateType.INVALID;
            }
        }

        /// <summary>
        /// Determines whether [has accessible state] [the specified obj].
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="state">The state.</param>
        /// <returns>
        /// 	<c>true</c> if [has accessible state] [the specified obj]; otherwise, <c>false</c>.
        /// </returns>
        public static bool HasAccessibleState(XAccessible obj, AccessibleStateType state)
        {
            if (obj != null)
            {
                try
                {
                    XAccessibleContext context = obj.getAccessibleContext();
                    if (context != null)
                    {
                        var states = context.getAccessibleStateSet();
                        if (states != null)
                        {
                            bool contains = states.contains((short)state);
                            if (contains)
                            {
                                short[] stateArray = states.getStates();
                                bool exists = Array.Exists(stateArray,
                                                delegate(short s) { return s.Equals((short)state); }
                                            );
                                contains = exists;
                            }
                            return contains;
                        }
                    }
                }
                catch (DisposedException ex) { }
            }
            return false;
        }

        /// <summary>
        /// Try to gets the root pane parent from the element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns>the XAccessible parent with role <c>AccessibleRole.ROOT_PANE</c> or <c>null</c></returns>
        public static XAccessible GetRootPaneFromElement(XAccessible element)
        {

            //XAccessible root = null;
            if (element != null)
            {
                XAccessibleContext cnt = element.getAccessibleContext();
                XAccessible parent = cnt.getAccessibleParent();

                int trys = 0;

                while (cnt != null
                    && parent != null
                    && !element.Equals(parent)
                    )
                {
                    // only for debug
                    var role = GetAccessibleRole(cnt);

                    if (!role.Equals(AccessibleRole.UNKNOWN))
                    {

                        trys = 0;
                        element = parent;

                        if (element != null) cnt = element.getAccessibleContext();
                        else cnt = null;

                        var role2 = GetAccessibleRole(cnt);

                        if (role2.Equals(AccessibleRole.ROOT_PANE))
                        {
                            return element;
                        }
                    }
                    else
                    {
                        if (++trys > 5) break;
                        Thread.Sleep(100);
                    }
                }

                Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "Can't get root pane - loop break");
            }

            return null;
        }

        /// <summary>
        /// Determines whether the the specified element contains accessible state (Must not be right!!).
        /// </summary>
        /// <param name="obj">The element to test.</param>
        /// <param name="state">The AccessibleStateType to test for.</param>
        /// <returns>
        /// 	<c>true</c> if [contains accessible state]; otherwise, <c>false</c>.
        /// </returns>
        public static bool ContainsAccessibleState(XAccessible obj, AccessibleStateType state)
        {
            if (obj != null)
            {
                XAccessibleContext context = obj.getAccessibleContext();
                if (context != null)
                {
                    var states = context.getAccessibleStateSet();
                    if (states != null)
                    {
                        return states.contains((short)state); ;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Try to gets all selected child elements.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <returns>A list of elements that have the AccessibleStateType.SELECTED = 23</returns>
        public static List<XAccessible> GetSelectedChildObjects(XAccessible parent)
        {
            List<XAccessible> selections = new List<XAccessible>();

            if (parent != null)
            {
                XAccessibleContext context = parent.getAccessibleContext();
                if (context != null)
                {
                    if (HasAccessibleState(parent, AccessibleStateType.SELECTED))
                    {
                        selections.Add(parent); //TODO: check type (context or XAccessible)
                    }
                    else
                    {
                        if (context.getAccessibleChildCount() > 0)
                        {
                            for (int i = 0; i < context.getAccessibleChildCount(); i++)
                            {
                                selections.AddRange(GetSelectedChildObjects(context.getAccessibleChild(i)));
                            }
                        }
                    }
                }
            }
            return selections;
        }

        /// <summary>
        /// Try to find the accessible counterpart to a XShape in the document.
        /// </summary>
        /// <param name="needle">The needle.</param>
        /// <param name="haystack">The XAccessible haystack (parent or document).</param>
        /// <returns>The accessible counterpart (if the name is equal) or NULL</returns>
        public static XAccessible _GetAccessibleCounterpart(XShape needle, XAccessibleContext haystack)
        {
            if (haystack != null && needle != null)
            {
                String name = OoUtils.GetStringProperty(needle, "Name");
                String title = OoUtils.GetStringProperty(needle, "Title");
                String search = (String.IsNullOrWhiteSpace(name) ? "" : name)
                    + (String.IsNullOrWhiteSpace(title) ? "" : " " + title);

                return GetAccessibleChildWithName(search, haystack);
            }
            return null;
        }


        public static XAccessible GetAccessibleCounterpartFromHash(XShape needle, XAccessibleContext haystack)
        {
            XAccessible counterpart = null;
            if (haystack != null && needle != null)
            {
                string hash = "###" + needle.GetHashCode().ToString() + "###";
                try
                {
                    string desc = OoUtils.GetStringProperty(needle, "Description");
                    bool success = OoUtils.SetStringProperty(needle, "Description", hash + desc);

                    if (success)
                    {

                        // find the accessible counterpart with the description
                        // starting with the corresponding hash pattern
                        counterpart = GetAccessibleChildDescriptionStartsWith(hash, haystack);

                        if (counterpart == null)
                        {
                            Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "[ERORR] could not find ax XAccessible width the given Description");
                        }

                    }
                    else
                    {
                        // FIXME: what todo if no description could been set
                    }


                    // if found reset the Description
                    success = OoUtils.SetStringProperty(needle, "Description", desc);
                    if (!success)
                    {
                        Logger.Instance.Log(LogPriority.DEBUG, "OoAccessibility", "[ERROR] Could not revert Description change for XShape while searching for their counterpart");
                    }


                }
                catch { }


            }
            return counterpart;
        }


        /// <summary>
        /// Gets the name of the accessible child with.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="haystack">The haystack.</param>
        /// <returns>The first found XAccessible with a partly equal name or NULL</returns>
        public static XAccessible GetAccessibleChildWithName(String name, XAccessible haystack)
        {
            if (haystack != null && !String.IsNullOrWhiteSpace(name))
            {
                if (haystack is XAccessibleContext) return GetAccessibleChildWithName(name, haystack as XAccessibleContext);
                else return GetAccessibleChildWithName(name, haystack.getAccessibleContext());
            }
            return null;
        }
        /// <summary>
        /// Gets the name of the accessible child with.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="haystack">The haystack.</param>
        /// <returns>The first found XAccessible with a partly equal name or NULL</returns>
        public static XAccessible GetAccessibleChildWithName(String name, XAccessibleContext haystack)
        {
            try
            {
                if (haystack != null && !String.IsNullOrWhiteSpace(name))
                {
                    if (GetAccessibleName(haystack).Equals(name))
                    {
                        return haystack as XAccessible;
                    }

                    if (haystack.getAccessibleChildCount() > 0)
                    {
                        int childCount = haystack.getAccessibleChildCount();
                        for (int i = 0; i < childCount; i++)
                        {
                            var c = haystack.getAccessibleChild(i);
                            if (c != null)
                            {
                                if (c.Equals(haystack)) continue;
                                var child = GetAccessibleChildWithName(name, c);
                                if (child != null) return child;
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }


        public static XAccessible GetAccessibleChildDescriptionStartsWith(String descriptionStart, XAccessibleContext haystack)
        {
            try
            {
                if (haystack != null && !String.IsNullOrWhiteSpace(descriptionStart))
                {
                    string desc = GetAccessibleDesc(haystack);
                    if (desc != null && desc.StartsWith(descriptionStart))
                    {
                        return haystack as XAccessible;
                    }

                    if (haystack.getAccessibleChildCount() > 0)
                    {
                        int childCount = haystack.getAccessibleChildCount();
                        for (int i = 0; i < childCount; i++)
                        {
                            var c = haystack.getAccessibleChild(i);
                            if (c != null)
                            {
                                if (c.Equals(haystack)) continue;
                                var child = GetAccessibleChildDescriptionStartsWith(descriptionStart, c.getAccessibleContext());
                                if (child != null) return child;
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }


        /// <summary>
        /// Gets the children of an accessible object.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <returns>list of accessible children of this node</returns>
        public static List<XAccessible> GetChildrenOfAccessibleObject(XAccessible parent)
        {
            List<XAccessible> children = new List<XAccessible>();

            if (parent != null)
            {
                XAccessibleContext context = parent.getAccessibleContext();
                if (context != null)
                {
                    if (context.getAccessibleChildCount() > 0)
                    {
                        for (int i = 0; i < context.getAccessibleChildCount(); i++)
                        {
                            try
                            {
                                var child = context.getAccessibleChild(i);
                                if (child != null) children.Add(child);
                            }
                            catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, "can't get child from XAccessible", ex); }
                        }
                    }
                }
            }
            return children;
        }

        /// <summary>
        /// Gets all children and sub children under the accessible object.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <returns>list of all available accessible children and sub children under this node</returns>
        public static List<XAccessible> GetAllChildrenOfAccessibleObject(XAccessible parent)
        {
            List<XAccessible> children = new List<XAccessible>();

            if (parent != null)
            {
                XAccessibleContext context = parent.getAccessibleContext();
                if (context != null)
                {
                    if (context.getAccessibleChildCount() > 0)
                    {
                        for (int i = 0; i < context.getAccessibleChildCount(); i++)
                        {
                            try
                            {
                                var child = context.getAccessibleChild(i);
                                if (child != null)
                                {
                                    children.Add(child);
                                    children.AddRange(GetAllChildrenOfAccessibleObject(child));
                                }
                            }
                            catch (Exception ex) { Logger.Instance.Log(LogPriority.DEBUG, "can't get all childen from XAccessible", ex); }
                        }
                    }
                }
            }
            return children;
        }

        /// <summary>
        /// Returns the parent of the element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns>The parent XAccessible or <c>null</c></returns>
        public static XAccessible GetParent(XAccessible element)
        {
            if (element != null)
            {
                XAccessibleContext cont = element.getAccessibleContext();
                if (cont != null)
                {
                    XAccessible parent = cont.getAccessibleParent();
                    if (!parent.Equals(element)) return element;
                }
            }
            return null;
        }

        /// <summary>
        /// Try to get the parent page.
        /// </summary>
        /// <param name="acc">The acc.</param>
        /// <returns></returns>
        public static XAccessible GetParentPage(XAccessible acc)
        {
            if (acc != null)
            {
                XAccessibleContext accCont = acc.getAccessibleContext();
                if (accCont != null)
                {
                    AccessibleRole role = GetAccessibleRole(accCont);
                    if (role == AccessibleRole.PAGE) return accCont as XAccessible;
                    else if (
                        role == AccessibleRole.WINDOW
                        || role == AccessibleRole.DOCUMENT
                        || role == AccessibleRole.ROOT_PANE
                        || role == AccessibleRole.MENU_BAR
                        ) return null;
                    else
                    {
                        XAccessible parent = GetParent(acc);
                        if (parent != null && !parent.Equals(acc) && !parent.Equals(accCont))
                        {
                            return GetParentPage(parent);
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Gets all OpenOffice top windows.
        /// </summary>
        /// <param name="extTollkit">The XExtendedToolkit if available.</param>
        /// <returns>List of top windows.</returns>
        public static List<Object> GetAllTopWindows(XExtendedToolkit extTollkit = null)
        {
            List<Object> result = new List<Object>();
            if (extTollkit == null) extTollkit = OO.GetExtTooklkit();
            if (extTollkit != null && extTollkit.getTopWindowCount() > 0)
            {
                for (int i = 0; i < extTollkit.getTopWindowCount(); i++)
                {
                    var tpw = extTollkit.getTopWindow(i);
                    if (tpw != null) result.Add(tpw);
                }
            }

            return result;
        }

        /// <summary>
        /// Gets the active top window.
        /// </summary>
        /// <param name="extTollkit">The ext tollkit.</param>
        /// <returns></returns>
        public static Object GetActiveTopWindow(XExtendedToolkit extTollkit = null)
        {
            if (extTollkit == null) extTollkit = OO.GetExtTooklkit();
            if (extTollkit != null)
            {
                try
                {
                    return extTollkit.getActiveTopWindow();
                }
                catch (Exception e)
                {
                    Logger.Instance.Log(LogPriority.IMPORTANT, "OoAccessibility", "[ERROR] while trying to get the active top window", e);
                    extTollkit = null;
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// Gets all "real" document windows from the available top windows.
        /// </summary>
        /// <returns></returns>
        public static List<Object> GetAllDocumentWindows()
        {
            List<Object> docWindows = new List<Object>();
            List<Object> tpws = GetAllTopWindows();
            foreach (var tpwnd in tpws)
            {
                if (tpwnd != null)
                {
                    if (tpwnd is XAccessible)
                    {
                        AccessibleRole role = GetAccessibleRole(tpwnd as XAccessible);
                        if (role.Equals(AccessibleRole.ROOT_PANE)) // get only ROOT_PANES as real "windows"
                        {
                            if (!HasAccessibleState(tpwnd as XAccessible, AccessibleStateType.INVALID)) // don't get invalid root panes
                            {
                                docWindows.Add(tpwnd);
                            }
                        }
                    }
                }
            }
            return docWindows;
        }

        #region OO DRAW

        /// <summary>
        /// Try to gets all available top windows of draw documents (DRAW windows).
        /// </summary>
        /// <returns>List of top window objects relating to draw documents</returns>
        public static List<Object> GetAllDrawWindows()
        {
            List<Object> drawWindows = new List<Object>();

            List<Object> docWnds = GetAllDocumentWindows();
            foreach (var docWnd in docWnds)
            {
                if (docWnd != null && docWnd is XAccessible)
                {
                    object draw = IsDrawWindow(docWnd as XAccessible);
                    if (!draw.Equals(false))
                    {
                        drawWindows.Add(draw);
                    }
                }
            }
            return drawWindows;
        }

        /// <summary>
        /// Determines whether the window [is a draw element].
        /// This function works only for determining if the main window of the DRAW application is sent to this function e.g. through a top window event.
        /// If you want to check if any other element could bee the parent of an DRAW document us the <see cref="IsDrawWindowParentPart"/> function (mutch slower).
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns><c>False</c> or the DOCUMENT</returns>
        public static object IsDrawWindow(XAccessible obj)
        {
            //util.Debug.GetAllServicesOfObject(obj);
            //util.Debug.GetAllProperties(obj);
            //util.Debug.GetAllInterfacesOfObject(obj);

            if (!(obj is XSystemDependentWindowPeer)) return false;

            return IsDrawWindowParentPart(obj);
        }


        /// <summary>
        /// Determines whether the window [is a draw element].
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns>False or the DOCUMENT</returns>
        public static object IsDrawWindowParentPart(XAccessible obj)
        {
            if (obj != null)
            {
                //check if object is com.sun.star.drawing.AccessibleDrawDocumentView
                if (OoUtils.ElementSupportsService(obj, OO.Services.DRAWING_ACCESSIBLE_DOC))
                {
                    return obj;
                }

                // cannot have any children windows
                if (!(obj is XVclContainer)) return false;

                AccessibleRole role = AccessibleRole.UNKNOWN;
                XAccessibleContext cont = obj is XAccessibleContext ? obj as XAccessibleContext : obj.getAccessibleContext();

                // if you cannot obtain the context than the other stuff 
                // will run into trouble - so break!
                if (cont == null) return false;

                //// try to check for drawing window by comparing children with existing drawing controller windows
                //try
                //{
                //    // get libre office desktop
                //    unoidl.com.sun.star.frame.XDesktop desktop = OO.GetDesktop(OO.GetMultiComponentFactory(OO.GetContext()), OO.GetContext());
                //    if (desktop != null)
                //    {
                //        // iterate desktop components
                //        unoidl.com.sun.star.container.XEnumeration desktopComponentsEnumeration = desktop.getComponents().createEnumeration();
                //        while (desktopComponentsEnumeration.hasMoreElements())
                //        {

                //            object comp = desktopComponentsEnumeration.nextElement().Value;
                //            if (comp != null && comp is XDrawPagesSupplier)
                //            {
                //                XDrawPagesSupplier drawPageSuppliesModel = (XDrawPagesSupplier)comp;
                //                // check only desktop components that is are drawPagesSuppliers (these are models too)
                //                if (drawPageSuppliesModel != null && OoUtils.ElementSupportsService(drawPageSuppliesModel, "com.sun.star.drawing.DrawingDocument") && drawPageSuppliesModel is unoidl.com.sun.star.frame.XModel)
                //                {
                //                    // get the draw page components controller and check if its supporting the Service DrawingDocumentDrawView
                //                    unoidl.com.sun.star.frame.XController currentController = ((unoidl.com.sun.star.frame.XModel)drawPageSuppliesModel).getCurrentController();
                //                    if (currentController != null && OoUtils.ElementSupportsService(currentController, "com.sun.star.drawing.DrawingDocumentDrawView"))
                //                    {
                //                        // get the component window
                //                        XWindow currentControllerComponentWindow = (XWindow)currentController.getFrame().getComponentWindow();
                //                        // get its rectangle
                //                        Rectangle currentRect = (Rectangle)currentControllerComponentWindow.getPosSize();
                //                        // search sources children for visible children if one matches the rectangle which indicates a match
                //                        if (obj is XVclContainer)
                //                        {
                //                            XWindow[] sourceChildWindows = ((XVclContainer)obj).getWindows();
                //                            foreach (XWindow sourceChildWindow in sourceChildWindows)
                //                            {
                //                                Rectangle sourceChildRect = (Rectangle)sourceChildWindow.getPosSize();
                //                                System.Diagnostics.Debug.WriteLine("Source Child Rect: " + sourceChildRect.X + ", " + sourceChildRect.Y + "    " + sourceChildRect.Width + " x " + sourceChildRect.Height);

                //                                if (((XWindow2)sourceChildWindow).isVisible() &&
                //                                    ((XWindow2)currentControllerComponentWindow).isVisible() &&
                //                                    ((XWindow2)sourceChildWindow).hasFocus() == ((XWindow2)currentControllerComponentWindow).hasFocus() &&
                //                                    sourceChildRect.X == currentRect.X &&
                //                                    sourceChildRect.Y == currentRect.Y &&
                //                                    sourceChildRect.Width == currentRect.Width &&
                //                                    sourceChildRect.Height == currentRect.Height
                //                                    )
                //                                {
                //                                    //System.Diagnostics.Debug.WriteLine("match rects");
                //                                    //return currentController;
                //                                }
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }

                //    }
                //}
                //catch (Exception ex)
                //{

                //}
                int ccount = 0;

                try
                {

                    // fast check to search further or stop
                    // if top window and the accessible name ends with a "- OpenOffice Draw"

                    if (obj is XWindow2 && ((XWindow2)obj).isVisible())
                    {
                        ccount = cont.getAccessibleChildCount();

                        // don't go to deep -- if the element has no children 
                        // then it is not the document or a valid container for a document
                        if (ccount < 1) return false;
                    }
                    else { return false; } // stop if obj is not a window or is invisible


                    if (denieObjectBecauseOfServiceCheck(cont)) return false;

                    int trys = 0;
                    while (role == AccessibleRole.UNKNOWN && (trys++ <= 2))
                    {
                        if (trys > 1) Thread.Sleep(5 * trys);
                        role = GetAccessibleRole(cont);
                    }

                    // if role is a surrounding container type, search for the draw doc children
                    if (role.Equals(AccessibleRole.PANEL)
                        || role.Equals(AccessibleRole.SCROLL_PANE)
                        || role.Equals(AccessibleRole.ROOT_PANE)
                        || role.Equals(AccessibleRole.SPLIT_PANE)
                        || role.Equals(AccessibleRole.STATUS_BAR)
                        || role.Equals(AccessibleRole.UNKNOWN) // FIXME: see this as a bad hack - remove this if possible
                        )
                    {
                        if (cont != null)
                        {
                            for (int i = 0; i < ccount; i++)
                            {
                                try
                                {
                                    XAccessible child = cont.getAccessibleChild(i);
                                    if (child != null)
                                    {
                                        Object result = IsDrawWindowParentPart(child);
                                        if (!(result is bool))
                                        {
                                            return result;
                                        }
                                        else continue;
                                    }
                                }
                                catch { return false; }
                            }
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        /// <summary>
        /// List of services that should be denied when try to decide if the opened window is a
        /// DRAW document or not
        /// </summary>
        private static readonly List<String> deniedServices = new List<String>
        {
            OO.Services.AWT_ACCESSIBILITY_MENU_SEPERATOR,
            OO.Services.AWT_ACCESSIBILITY_MENU,
            OO.Services.AWT_ACCESSIBILITY_MENU_ITEM,
            OO.Services.AWT_ACCESSIBILITY_MENUBAR,
            OO.Services.AWT_ACCESSIBILITY_POPUP_MENU,
            OO.Services.AWT_ACCESSIBILITY_FIXEDTEXT,
            OO.Services.AWT_ACCESSIBILITY_EDIT,
            OO.Services.AWT_ACCESSIBILITY_BUTTON,
            OO.Services.AWT_ACCESSIBILITY_SCROLLBAR,
            OO.Services.AWT_ACCESSIBILITY_TAB_PAGE,
            OO.Services.AWT_ACCESSIBILITY_TAB_CONTROL
        };

        /// <summary>
        /// List of services that are allowed for a deeper search.
        /// Is only used for debugging reasons. Check denied against accepted helps
        /// to find unknown services.
        /// </summary>
        private static readonly List<String> acceptedServices = new List<string>
        {
            OO.Services.AWT_ACCESSIBILITY_WINDOW
        };

        /// <summary>
        /// Checks if the service returned by the obj should be denied or further investigated
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        private static bool denieObjectBecauseOfServiceCheck(Object obj)
        {
            if (obj is XServiceInfo)
            {
                string[] services = util.Debug.GetAllServicesOfObject(obj, false);

                foreach (var service in services)
                {
                    //TODO: maybe change this and only allow accessible windows
                    if (acceptedServices.Contains(service)) return false;
                    else return true;
                }

                //FIXME: Only for debug 
                System.Diagnostics.Debug.WriteLine("...  [HINT]  Can't decide if Service should be accepted or not: " + String.Join(";", services));

            }
            return false;
        }


        #endregion
    }

    #region Enums

    /// <summary>
    /// Collection of roles. 
    /// This collection of constants defines the set of possible roles of 
    /// classes implementing the XAccessible interface according to the 
    /// java class javax.accessibility.AccessibleRole. The role of an 
    /// object describes its generic function like 'button', 'menu', or 
    /// 'text'. You can obtain an object's role by calling the 
    /// getAccessibleRole method of the XAccessibleContext interface.
    /// We are using constants instead of a more typesafe enum. The 
    /// reason for this is that IDL enums may not be extended. Therefore,
    /// in order to include future extensions to the set of roles we have
    /// to use constants here.
    /// For some roles there exist two labels with the same value. Please
    /// use the one with the underscores. The other ones are somewhat 
    /// deprecated and will be removed in the future.
    /// </summary>
    public enum AccessibleRole : short
    {
        /// <summary>
        /// It is knows that the object is already disposed
        /// </summary>
        DISPOSED = -2,
        /// <summary>
        /// the request for the role could not return a valid value
        /// </summary>
        INVALID = -1,
        /// <summary>
        /// Object is used to alert the user about something.
        /// </summary>
        ALERT = 1,
        /// <summary>
        /// Button dropdown role  
        /// </summary>
        BUTTON_DROPDOWN = 68,
        /// <summary>
        /// Button menu role  
        /// </summary>
        BUTTON_MENU = 69,
        /// <summary>
        /// Object that can be drawn into and is used to trap events. 
        /// </summary>
        CANVAS = 3,
        /// <summary>
        /// Caption role
        /// </summary>
        CAPTION = 70,
        /// <summary>
        /// Chart role  
        /// </summary>
        CHART = 71,
        /// <summary>
        /// Check box role.
        /// </summary>
        CHECK_BOX = 4,
        /// <summary>
        /// This role is used for check buttons that are menu items.
        /// </summary>
        CHECK_MENU_ITEM = 5,
        /// <summary>
        /// A specialized dialog that lets the user choose a color.
        /// </summary>
        COLOR_CHOOSER = 6,
        /// <summary>
        /// The header for a column of data. 
        /// </summary>
        COLUMN_HEADER = 2,
        /// <summary>
        /// Combo box role. 
        /// </summary>
        COMBO_BOX = 7,
        /// <summary>
        /// Comment role 
        /// </summary>
        COMMENT = 81,
        /// <summary>
        /// Comment end role  
        /// </summary>
        COMMENT_END = 82,
        /// <summary>
        /// Date editor role.
        /// </summary>
        DATE_EDITOR = 8,
        /// <summary>
        /// An iconified internal frame in a DESKTOP_PANE.
        /// </summary>
        DESKTOP_ICON = 9,
        /// <summary>
        /// Desktop pane role.
        /// </summary>
        DESKTOP_PANE = 10,
        /// <summary>
        /// Dialog box role.
        /// </summary>
        DIALOG = 12,
        /// <summary>
        /// Directory pane role. 
        /// </summary>
        DIRECTORY_PANE = 11,
        /// <summary>
        /// View of a document.
        /// </summary>
        DOCUMENT = 13,
        /// <summary>
        /// Edit bar role 
        /// </summary>
        EDIT_BAR = 72,
        /// <summary>
        /// Embedded (OLE) object. 
        /// </summary>
        EMBEDDED_OBJECT = 14,
        /// <summary>
        /// Text that is used as an endnote (footnote at the end 
        /// of a chapter or section.
        /// </summary>
        END_NOTE = 15,
        /// <summary>
        /// File chooser role.
        /// </summary>
        FILE_CHOOSER = 16,
        /// <summary>
        /// Filler role.
        /// </summary>
        FILLER = 17,
        /// <summary>
        /// Font chooser role.
        /// </summary>
        FONT_CHOOSER = 18,
        /// <summary>
        /// Footer of a document page.
        /// </summary>
        FOOTER = 19,
        /// <summary>
        /// Text that is used as a footnote. 
        /// </summary>
        FOOTNOTE = 20,
        /// <summary>
        /// Form role  
        /// </summary>
        FORM = 73,
        /// <summary>
        /// Frame role.
        /// </summary>
        FRAME = 21,
        /// <summary>
        /// Glass pane role.
        /// </summary>
        GLASS_PANE = 22,
        /// <summary>
        /// Graphical object.
        /// </summary>
        GRAPHIC = 23,
        /// <summary>
        /// Group box role.
        /// </summary>
        GROUP_BOX = 24,
        /// <summary>
        /// Header of a document page.
        /// </summary>
        HEADER = 25,
        /// <summary>
        /// Chapter or section heading. 
        /// </summary>
        HEADING = 26,
        /// <summary>
        /// A hypertext anchor
        /// </summary>
        HYPER_LINK = 27,
        /// <summary>
        /// A small fixed size picture, typically 
        /// used to decorate components
        /// </summary>
        ICON = 28,
        /// <summary>
        /// Image map role  
        /// </summary>
        IMAGE_MAP = 74,
        /// <summary>
        /// Internal frame role.
        /// </summary>
        INTERNAL_FRAME = 29,
        /// <summary>
        /// An object used to present an icon or short 
        /// string in an interface.
        /// </summary>
        LABEL = 30,
        /// <summary>
        /// layered pane role.  
        /// </summary>
        LAYERED_PANE = 31,
        /// <summary>
        /// List role.
        /// </summary>
        LIST = 32,
        /// <summary>
        /// List item role.
        /// </summary>
        LIST_ITEM = 33,
        /// <summary>
        /// Menu role.
        /// </summary>
        MENU = 34,
        /// <summary>
        /// Menu bar role.
        /// </summary>
        MENU_BAR = 35,
        /// <summary>
        /// Menu item role.
        /// </summary>
        MENU_ITEM = 36,
        /// <summary>
        /// Note role  
        /// </summary>
        NOTE = 75,
        /// <summary>
        /// A specialized pane whose primary use is inside a DIALOG. 
        /// </summary>
        OPTION_PANE = 37,
        /// <summary>
        /// Page role 
        /// </summary>
        PAGE = 76,
        /// <summary>
        /// Page tab role.
        /// </summary>
        PAGE_TAB = 38,
        /// <summary>
        /// Page tab list role.
        /// </summary>
        PAGE_TAB_LIST = 39,
        /// <summary>
        /// A generic container that is often used to group objects.
        /// </summary>
        PANEL = 40,
        /// <summary>
        /// Paragraph of text.
        /// </summary>
        PARAGRAPH = 41,
        /// <summary>
        /// Password text role.
        /// </summary>
        PASSWORD_TEXT = 42,
        /// <summary>
        /// Popup menu role.
        /// </summary>
        POPUP_MENU = 43,
        /// <summary>
        /// An object used to indicate how much of a task has been 
        /// completed.
        /// </summary>
        PROGRESS_BAR = 45,
        /// <summary>
        /// Push button role.
        /// </summary>
        PUSH_BUTTON = 44,
        /// <summary>
        /// Radio button role.
        /// </summary>
        RADIO_BUTTON = 46,
        /// <summary>
        /// This role is used for radio buttons that are menu items. 
        /// </summary>
        RADIO_MENU_ITEM = 47,
        /// <summary>
        /// Root pane role.
        /// </summary>
        ROOT_PANE = 49,
        /// <summary>
        /// The header for a row of data.
        /// </summary>
        ROW_HEADER = 48,
        /// <summary>
        /// Ruler role  
        /// </summary>
        RULER = 77,
        /// <summary>
        /// Scroll bar role.
        /// </summary>
        SCROLL_BAR = 50,
        /// <summary>
        /// Scroll pane role.
        /// </summary>
        SCROLL_PANE = 51,
        /// <summary>
        /// Section role 
        /// </summary>
        SECTION = 78,
        /// <summary>
        /// Separator role.  
        /// </summary>
        SEPARATOR = 53,
        /// <summary>
        /// Object with graphical representation used to represent 
        /// content on draw pages. 
        /// </summary>
        SHAPE = 52,
        /// <summary>
        /// Slider role. 
        /// </summary>
        SLIDER = 54,
        /// <summary>
        /// Spin box role.  
        /// </summary>
        SPIN_BOX = 55,
        /// <summary>
        /// Split pane role.  
        /// </summary>
        SPLIT_PANE = 56,
        /// <summary>
        /// Status bar role.  
        /// </summary>
        STATUS_BAR = 57,
        /// <summary>
        /// Table component.  
        /// </summary>
        TABLE = 58,
        /// <summary>
        /// Single cell in a table.  
        /// </summary>
        TABLE_CELL = 59,
        /// <summary>
        /// Text role.  
        /// </summary>
        TEXT = 60,
        /// <summary>
        /// Collection of objects that constitute a logical text entity. 
        /// </summary>
        TEXT_FRAME = 61,
        /// <summary>
        /// Toggle button role.  
        /// </summary>
        TOGGLE_BUTTON = 62,
        /// <summary>
        /// Tool bar role.  
        /// </summary>
        TOOL_BAR = 63,
        /// <summary>
        /// Tool tip role. 
        /// </summary>
        TOOL_TIP = 64,
        /// <summary>
        /// Tree role.  
        /// </summary>
        TREE = 65,
        /// <summary>
        /// Tree item role  
        /// </summary>
        TREE_ITEM = 79,
        /// <summary>
        /// Tree table role 
        /// </summary>
        TREE_TABLE = 80,
        /// <summary>
        /// Unknown role. 
        /// </summary>
        UNKNOWN = 0,
        /// <summary>
        /// Viewport role.  
        /// </summary>
        VIEW_PORT = 66,
        /// <summary>
        /// A top level window with no title or border.  
        /// </summary>
        WINDOW = 67,
    }

    /// <summary>
    /// Collection of state types. 
    /// This list of constants defines the available set of states that an 
    /// object that implements XAccessibleContext can be in.
    /// The comments describing the states is taken verbatim from the Java 
    /// Accessibility API 1.4 documentation.
    /// We are using constants instead of a more typesafe enum. The reason 
    /// for this is that IDL enums may not be extended. Therefore, in order 
    /// to include future extensions to the set of roles we have to use 
    /// constants here.
    /// </summary>
    public enum AccessibleStateType : short
    {
        /// <summary>
        /// Indicates a window is currently the active window.  
        /// </summary>
        ACTIVE = 1,
        /// <summary>
        /// Indicates that the object is armed.  
        /// </summary>
        ARMED = 2,
        /// <summary>
        /// Indicates the current object is busy.  
        /// </summary>
        BUSY = 3,
        /// <summary>
        /// Indicates this object is currently checked.  
        /// </summary>
        CHECKED = 4,
        /// <summary>
        /// 
        /// </summary>
        COLLAPSE = 34,
        /// <summary>
        /// 
        /// </summary>
        DEFAULT = 32,
        /// <summary>
        /// User interface object corresponding to this object no longer 
        /// exists.  
        /// </summary>
        DEFUNC = 5,
        /// <summary>
        /// Indicates the user can change the contents of this object.
        /// </summary>
        EDITABLE = 6,
        /// <summary>
        /// Indicates this object is enabled.
        /// </summary>
        ENABLED = 7,
        /// <summary>
        /// Indicates this object allows progressive disclosure of its 
        /// children. 
        /// </summary>
        EXPANDABLE = 8,
        /// <summary>
        /// Indicates this object is expanded. 
        /// </summary>
        EXPANDED = 9,
        /// <summary>
        /// Object can accept the keyboard focus.
        /// </summary>
        FOCUSABLE = 10,
        /// <summary>
        /// Indicates this object currently has the keyboard focus.
        /// </summary>
        FOCUSED = 11,
        /// <summary>
        /// Indicates the orientation of this object is horizontal.
        /// </summary>
        HORIZONTAL = 12,
        /// <summary>
        /// Indicates this object is minimized and is represented only
        /// by an icon. 
        /// </summary>
        ICONIFIED = 13,
        /// <summary>
        /// Sometimes UI elements can have a state indeterminate. 
        /// This can happen e.g. if a check box reflects the bold state 
        /// of text in a text processor. When the current selection 
        /// contains text which is bold and also text which is not bold, 
        /// the state is indeterminate.
        /// </summary>
        INDETERMINATE = 14,
        /// <summary>
        /// Indicates an invalid state.  
        /// </summary>
        INVALID = 0,
        /// <summary>
        /// Indicates the most (all) children are transient and it is 
        /// not necessary to add listener to the children. Only the 
        /// active descendant (given by the event) should be not 
        /// transient to make it possible to add listener to this 
        /// object and recognize changes in this object. The state is 
        /// added to make a performance improvement. Now it is no longer 
        /// necessary to iterate over all children to find out whether 
        /// they are transient or not to decide whether to add listener 
        /// or not. If there is a object with this state no one should 
        /// iterate over the children to add listener. Only the active 
        /// descendant should get listener if it is not transient.
        /// </summary>
        MANAGES_DESCENDANTS = 15,
        /// <summary>
        /// Object is modal.
        /// </summary>
        MODAL = 16,
        /// <summary>
        /// 
        /// </summary>
        MOVEABLE = 31,
        /// <summary>
        /// Indicates this (text) object can contain multiple lines 
        /// of text 
        /// </summary>
        MULTI_LINE = 17,
        /// <summary>
        /// More than one child may be selected at the same time.
        /// </summary>
        MULTI_SELECTABLE = 18,
        /// <summary>
        /// 
        /// </summary>
        OFFSCREEN = 33,
        /// <summary>
        /// Indicates this object paints every pixel within its 
        /// rectangular region.
        /// </summary>
        OPAQUE = 19,
        /// <summary>
        /// Indicates this object is currently pressed.
        /// </summary>
        PRESSED = 20,
        /// <summary>
        /// Indicates the size of this object is not fixed.
        /// </summary>
        RESIZABLE = 21,
        /// <summary>
        /// Object is selectable.
        /// </summary>
        SELECTABLE = 22,
        /// <summary>
        /// Object is selected.
        /// </summary>
        SELECTED = 23,
        /// <summary>
        /// Indicates this object is sensitive.
        /// </summary>
        SENSITIVE = 24,
        /// <summary>
        /// Object is displayed on the screen.
        /// </summary>
        SHOWING = 25,
        /// <summary>
        /// Indicates this (text) object can contain only a 
        /// single line of text
        /// </summary>
        SINGLE_LINE = 26,
        /// <summary>
        /// 
        /// </summary>
        STALE = 27,
        /// <summary>
        /// Indicates this object is transient.
        /// </summary>
        TRANSIENT = 28,
        /// <summary>
        /// Indicates the orientation of this object is vertical.  
        /// </summary>
        VERTICAL = 29,
        /// <summary>
        /// Object wants to be displayed on the screen.  
        /// </summary>
        VISIBLE = 30,
    }

    /// <summary>
    /// These constants identify the type of AccessibleEventObject objects. 
    /// The AccessibleEventObject::OldValue and AccessibleEventObject::NewValue 
    /// fields contain, where applicable and not otherwise stated, the old and 
    /// new value of the property in question.
    /// </summary>
    public enum AccessibleEventId : short
    {
        /// <summary>
        /// 
        /// </summary>
        NONE = 0,
        /// <summary>
        /// The change of the number or attributes of actions of an 
        /// accessible object is signaled by events of this type.
        /// </summary>
        ACTION_CHANGED = 3,
        /// <summary>
        /// Constant used to determine when the active descendant of a
        /// component has changed. The active descendant is used in 
        /// objects with transient children. The
        /// AccessibleEventObject::NewValue contains the now active 
        /// object. The AccessibleEventObject::OldValue contains the 
        /// previously active child. Empty references indicate that 
        /// no child has been respectively is currently active.
        /// </summary>
        ACTIVE_DESCENDANT_CHANGED = 5,
        /// <summary>
        /// 
        /// </summary>
        ACTIVE_DESCENDANT_CHANGED_NOFOCUS = 34,
        /// <summary>
        /// This event indicates a change of the bounding rectangle of 
        /// an accessible object with respect only to its size or relative 
        /// position. If the absolute position changes but not the 
        /// relative position then its is not necessary to send an event.
        /// </summary>
        BOUNDRECT_CHANGED = 6,
        /// <summary>
        /// Events of this type are sent when the caret has moved to a 
        /// new position. The old and new position can be found in the 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields.
        /// </summary>
        CARET_CHANGED = 20,
        /// <summary>
        /// A child event indicates the addition of a new or the removal 
        /// of an existing child. The contents of the 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields determines which of 
        /// both has taken place.
        /// </summary>
        CHILD = 7,
        /// <summary>
        /// 
        /// </summary>
        COLUMN_CHANGED = 40,
        /// <summary>
        /// Identifies the change of a relation set: The content flow has 
        /// changed.
        /// </summary>
        CONTENT_FLOWS_FROM_RELATION_CHANGED = 12,
        /// <summary>
        /// Identifies the change of a relation set: The content flow has 
        /// changed.
        /// </summary>
        CONTENT_FLOWS_TO_RELATION_CHANGED = 13,
        /// <summary>
        /// Identifies the change of a relation set: The target object 
        /// that is doing the controlling has changed. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and 
        /// new controlling objects.
        /// </summary>
        CONTROLLED_BY_RELATION_CHANGED = 14,
        /// <summary>
        /// Identifies the change of a relation set: The controller for 
        /// the target object has changed. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and 
        /// new number of controlled objects.
        /// </summary>
        CONTROLLER_FOR_RELATION_CHANGED = 15,
        /// <summary>
        /// Use this event type to indicate a change of the description 
        /// string of an accessible object. The AccessibleEventObject::OldValue 
        /// and AccessibleEventObject::NewValue fields contain the description
        /// before and after the change.
        /// </summary>
        DESCRIPTION_CHANGED = 2,
        /// <summary>
        /// Constant used to indicate that a hypertext element has received 
        /// focus. The AccessibleEventObject::OldValue field contains 
        /// the start index of previously focused element. 
        /// The AccessibleEventObject::NewValue field holds the start index 
        /// in the document of the current element that has focus. A value 
        /// of -1 indicates that an element does not or did not have focus. 
        /// </summary>
        HYPERTEXT_CHANGED = 24,
        /// <summary>
        /// Use this event to tell the listeners to re-retrieve the whole 
        /// set of children. This should be used by a parent object which 
        /// exchanges all or most of its children. It is a short form of 
        /// first sending one CHILD event for every old child indicating 
        /// that this child is about to be removed and then sending one 
        /// CHILD for every new child indicating that this child has been 
        /// added to the list of children.
        /// </summary>
        INVALIDATE_ALL_CHILDREN = 8,
        /// <summary>
        /// Identifies the change of a relation set: The target group for 
        /// a label has changed. The AccessibleEventObject::OldValue 
        /// and AccessibleEventObject::NewValue fields contain the old 
        /// and new number labeled objects.
        /// </summary>
        LABEL_FOR_RELATION_CHANGED = 16,
        /// <summary>
        /// Identifies the change of a relation set: The objects that are 
        /// doing the labeling have changed. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and 
        /// new accessible label.
        /// </summary>
        LABELED_BY_RELATION_CHANGED = 17,
        /// <summary>
        /// Constant used to indicate that a list box entry has been 
        /// collapsed. AccessibleEventObject::OldValue is empty. 
        /// AccessibleEventObject::NewValue contains the collapsed list 
        /// box entry.
        /// </summary>
        LISTBOX_ENTRY_COLLAPSED = 33,
        /// <summary>
        /// Constant used to indicate that a list box entry has been 
        /// expanded. AccessibleEventObject::OldValue is empty. 
        /// AccessibleEventObject::NewValue contains the expanded list 
        /// box entry.
        /// </summary>
        LISTBOX_ENTRY_EXPANDED = 32,
        /// <summary>
        /// Identifies the change of a relation set: The group membership 
        /// has changed. The AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and 
        /// new number of members.
        /// </summary>
        MEMBER_OF_RELATION_CHANGED = 18,
        /// <summary>
        /// Use this event type to indicate a change of the name string 
        /// of an accessible object. The AccessibleEventObject::OldValue 
        /// and AccessibleEventObject::NewValue fields contain the name 
        /// before and after the change.  
        /// </summary>
        NAME_CHANGED = 1,
        /// <summary>
        /// 
        /// </summary>
        PAGE_CHANGED = 38,
        /// <summary>
        /// </summary>
        SECTION_CHANGED = 39,
        /// <summary>
        /// Events of this type indicate changes of the selection. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields remain empty.
        /// </summary>
        SELECTION_CHANGED = 9,
        /// <summary>
        /// Events of this type indicate changes of the selection. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields remain empty.
        /// </summary>
        SELECTION_CHANGED_ADD = 35,
        /// <summary>
        /// Events of this type indicate changes of the selection. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields remain empty.
        /// </summary>
        SELECTION_CHANGED_REMOVE = 36,
        /// <summary>
        /// Events of this type indicate changes of the selection. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields remain empty. 
        /// </summary>
        SELECTION_CHANGED_WITHIN = 37,
        /// <summary>
        /// State changes are signaled with this event type. 
        /// Use one event for every state that is set or reset. 
        /// The AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old 
        /// and new value respectively. To set a state put the state 
        /// id into the AccessibleEventObject::NewValue field and 
        /// leave AccessibleEventObject::OldValue empty. To reset a 
        /// state put the state id into the 
        /// AccessibleEventObject::OldValue field and leave 
        /// AccessibleEventObject::NewValue empty.
        /// </summary>
        STATE_CHANGED = 4,
        /// <summary>
        /// Identifies the change of a relation set: The sub-window-of 
        /// relation has changed. The AccessibleEventObject::OldValue 
        /// and AccessibleEventObject::NewValue fields contain the old 
        /// and new accessible parent window objects.
        /// </summary>
        SUB_WINDOW_OF_RELATION_CHANGED = 19,
        /// <summary>
        /// Constant used to indicate that the table caption has changed. 
        /// The AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and 
        /// new accessible objects representing the table caption.
        /// </summary>
        TABLE_CAPTION_CHANGED = 25,
        /// <summary>
        /// Constant used to indicate that the column description has 
        /// changed. The AccessibleEventObject::NewValue field contains 
        /// the column index. The AccessibleEventObject::OldValue is left 
        /// empty.
        /// </summary>
        TABLE_COLUMN_DESCRIPTION_CHANGED = 26,
        /// <summary>
        /// Constant used to indicate that the column header has changed. 
        /// The AccessibleEventObject::OldValue is empty, the 
        /// AccessibleEventObject::NewValue field contains an 
        /// AccessibleTableModelChange representing the header change.
        /// </summary>
        TABLE_COLUMN_HEADER_CHANGED = 27,
        /// <summary>
        /// Constant used to indicate that the table data has changed. 
        /// The AccessibleEventObject::OldValue is empty, the 
        /// AccessibleEventObject::NewValue field contains an 
        /// AccessibleTableModelChange representing the data change.
        /// </summary>
        TABLE_MODEL_CHANGED = 28,
        /// <summary>
        /// Constant used to indicate that the row description has 
        /// changed. The AccessibleEventObject::NewValue field contains 
        /// the row index. The AccessibleEventObject::OldValue is left 
        /// empty.
        /// </summary>
        TABLE_ROW_DESCRIPTION_CHANGED = 29,
        /// <summary>
        /// Constant used to indicate that the row header has changed. 
        /// The AccessibleEventObject::OldValue is empty, the 
        /// AccessibleEventObject::NewValue field contains an 
        /// AccessibleTableModelChange representing the header change.
        /// </summary>
        TABLE_ROW_HEADER_CHANGED = 30,
        /// <summary>
        /// Constant used to indicate that the table summary has changed. 
        /// The AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields contain the old and new 
        /// accessible objects representing the table summary.
        /// </summary>
        TABLE_SUMMARY_CHANGED = 31,
        /// <summary>
        /// This entry is reserved for future extension. 
        /// Don't use it right now.
        /// </summary>
        TEXT_ATTRIBUTE_CHANGED = 23,
        /// <summary>
        /// Use this id to indicate general text changes, i.e. 
        /// changes to text that is exposed through the XAccessibleText 
        /// and XAccessibleEditableText interfaces.
        /// </summary>
        TEXT_CHANGED = 22,
        /// <summary>
        /// Events of this type signal changes of the selection. The 
        /// old or new selection is not available through the event 
        /// object. You have to query the XAccessibleText interface 
        /// of the event source for this information. The type of 
        /// content of the AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields is not specified 
        /// at the moment. This may change in the future.
        /// </summary>
        TEXT_SELECTION_CHANGED = 21,
        /// <summary>
        /// This constant indicates changes of the value of an 
        /// XAccessibleValue interface. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue field contain the old and 
        /// new value as a number. Its exact type is implementation 
        /// dependent but has to be the same as is returned by the 
        /// XAccessibleValue::getCurrentValue function.
        /// </summary>
        VALUE_CHANGED = 11,
        /// <summary>
        /// A visible data event indicates the change of the visual 
        /// appearance of an accessible object. This includes for example 
        /// most of the attributes available over the XAccessibleComponent 
        /// and XAccessibleExtendedComponent interfaces. The 
        /// AccessibleEventObject::OldValue and 
        /// AccessibleEventObject::NewValue fields are left empty.
        /// </summary>
        VISIBLE_DATA_CHANGED = 10
    }

    /// <summary>
    /// Collection of relation types. 
    /// This list of constants defines the availabe types of relations 
    /// that are usable by AccessibleRelation.
    /// We are using constants instead of a more typesafe enum. 
    /// The reason for this is that IDL enums may not be extended. 
    /// Therefore, in order to include future extensions to the set of 
    /// roles we have to use constants here.
    /// </summary>
    public enum AccessibleRelationType : short
    {
        /// <summary>
        /// Content-flows-from relation.  
        /// </summary>
        CONTENT_FLOWS_FROM = 1,
        /// <summary>
        /// Content-flows-to relation.  
        /// </summary>
        CONTENT_FLOWS_TO = 2,
        /// <summary>
        /// Controlled-by relation type.  
        /// </summary>
        CONTROLLED_BY = 3,
        /// <summary>
        /// Controller-for relation type.  
        /// </summary>
        CONTROLLER_FOR = 4,
        /// <summary>
        /// 
        /// </summary>
        DESCRIBED_BY = 10,
        /// <summary>
        /// Invalid relation type.
        /// Indicates an invalid relation type. This is used to indicate 
        /// that a retrieval method could not find a requested relation.
        /// </summary>
        INVALID = 0,
        /// <summary>
        /// Lable-for relation type.  
        /// </summary>
        LABEL_FOR = 5,
        /// <summary>
        /// Labeled-by relation type.  
        /// </summary>
        LABELED_BY = 6,
        /// <summary>
        /// Member-of relation type.  
        /// </summary>
        MEMBER_OF = 7,
        /// <summary>
        /// Node-Child-of relation type. 
        /// Indicates an object is a cell in a tree or treetable 
        /// which is displayed because a cell in the same column 
        /// is expanded and identifies that cell.
        /// </summary>
        NODE_CHILD_OF = 9,
        /// <summary>
        /// Sub-Window-of relation type. 
        /// With this relation you can realize an alternative 
        /// parent-child relationship. The target of the relation 
        /// contains the parent window. Note that there is no 
        /// relation that points the other way, from the parent 
        /// window to the child window. 
        /// </summary>
        SUB_WINDOW_OF = 8,
    }

    #endregion

}