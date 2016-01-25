using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using tud.mci.tangram.audio;

namespace tud.mci.tangram
{
    public static class UiaPicker
    {
        /// <summary>
        /// Gets element at the given screen position
        /// </summary>
        /// <param name="x">x value of screen position</param>
        /// <param name="y">y value of screen position</param>
        /// <returns>uia element at the given screen position</returns>
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
        /// Speaks name of the uia element at the given screen position.
        /// </summary>
        /// <param name="x">x value of the screen position</param>
        /// <param name="y">y value of the screen position</param>
        public static void SpeakElementOnSreenPosition(int x, int y)
        {
            AutomationElement element = GetElementFromScreenPosition(x, y);
            SpeakElement(element);
        }

        /// <summary>
        /// Speaks name and type of the uia element.
        /// </summary>
        /// <param name="element">uia element</param>
        public static void SpeakElement(AutomationElement element)
        {
            if (element != null)
            {
                try
                {
                    string type = getGermanControlTypeName(element.Current.ControlType);
                    AudioRenderer.Instance.PlaySoundImmediately(type + element.Current.Name);
                }
                catch (Exception)
                {
                }
            }
        }

        /// <summary>
        /// Get the German word for the given ControlType.
        /// </summary>
        /// <param name="type">ControlType</param>
        /// <returns>German word for speaking the ControlType</returns>
        private static string getGermanControlTypeName(ControlType type)
        {
            if (type.Equals(ControlType.Button))
                return "Schalter ";
            else if (type.Equals(ControlType.CheckBox))
                return "Kontrollfeld ";
            else if (type.Equals(ControlType.Edit))
                return "Eingabefeld ";
            else if (type.Equals(ControlType.Hyperlink))
                return "Link ";
            else if (type.Equals(ControlType.Image))
                return "Grafik ";
            else if (type.Equals(ControlType.Menu) || type.Equals(ControlType.MenuItem))
                return "Menü ";
            else if (type.Equals(ControlType.RadioButton))
                return "Auswahlschalter ";
            else if (type.Equals(ControlType.Table))
                return "Tabelle ";           
            else return "";
        }

        public static AutomationElement GetFocusedElement()
        {
            AutomationElement element = AutomationElement.FocusedElement;
            SpeakElement(element);

            //if (element.Current.ControlType.Equals(ControlType.Window))
            //{
            //    int id = element.Current.ProcessId;
            //    //HwndSource hwndSource = HwndSource.FromHwnd(id);
            //    //JFrame.getFocusOwner();
            //}
            return element;
        }
    }
}
