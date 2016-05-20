using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;

namespace tud.mci.tangram.Uia
{
    public static class UiaPicker
    {
        #region Members

        static readonly LL ll = new LL(Properties.Resources.Language);

        #endregion

        /// <summary>
        /// Gets element at the given screen position
        /// </summary>
        /// <param name="x">x value of screen position</param>
        /// <param name="y">y value of screen position</param>
        /// <returns>UIA element at the given screen position</returns>
        public static AutomationElement GetElementFromScreenPosition(int x, int y)
        {
            AutomationElement element = AutomationElement.FromPoint(new System.Windows.Point(x, y));
            if (element != null) try
                {
                    Console.WriteLine("Element at position " + x + "," + y + " is '" + element.Current.Name + "'");
                    Logger.Instance.Log(LogPriority.MIDDLE, "UIA Picker", "[PICK] Element at position " + x + "," + y + " is '" + element.Current.Name + "'");                    
                }
                catch (Exception) {}
            return element;
        }

        /// <summary>
        /// Speaks name of the UIA element at the given screen position.
        /// </summary>
        /// <param name="x">x value of the screen position</param>
        /// <param name="y">y value of the screen position</param>
        public static void SpeakElementOnSreenPosition(int x, int y)
        {
            AutomationElement element = GetElementFromScreenPosition(x, y);
            SpeakElement(element);
        }

        /// <summary>
        /// Speaks name and type of the UIA element.
        /// </summary>
        /// <param name="element">UIA element</param>
        public static void SpeakElement(AutomationElement element)
        {
            if (element != null)
            {
                try
                {
                    var typeName = element.Current.ControlType.ProgrammaticName;
                    string type = ll.GetTrans("Name." + (String.IsNullOrWhiteSpace(typeName) ? "ControlType.unknown" : typeName));
                    AudioRenderer.Instance.PlaySoundImmediately(type + " " + element.Current.Name);
                }
                catch (Exception)
                {
                }
            }
        }

        /// <summary>
        /// Gets the focused element by the UIA interface.
        /// </summary>
        /// <returns>The focused element</returns>
        public static AutomationElement GetFocusedElement()
        {
            AutomationElement element = AutomationElement.FocusedElement;
            SpeakElement(element);
            return element;
        }
    }
}
