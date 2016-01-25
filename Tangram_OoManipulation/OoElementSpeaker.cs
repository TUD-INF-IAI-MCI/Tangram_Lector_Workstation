using System;
using tud.mci.tangram;
using tud.mci.tangram.Accessibility;
using tud.mci.tangram.controller.observer;
using tud.mci.tangram.TangramLector;
using tud.mci.LanguageLocalization;
using tud.mci.tangram.audio;

namespace TangramLector.OO
{
    public static class OoElementSpeaker
    {
        static AudioRenderer audio = AudioRenderer.Instance;
        static readonly LL LL = new LL(tud.mci.tangram.TangramLector.Properties.Resources.Language);

        /// <summary>
        /// Gets the audio text for the given element and send it to the audio renderer.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties and befor the label (if available).</param>
        /// <returns>The describing string of the element in the following form: [ROLE] [NAME] [TITLE] (addintionalText) [TEXT]</returns>
        public static string PlayElement(OoShapeObserver element, string additionalText = "")
        {
            String text = GetElementAudioText(element, additionalText);
            if (!String.IsNullOrWhiteSpace(text))
            {
                audio.PlaySound(text);
            }
            return text;
        }

        /// <summary>
        /// Gets the audio text for the given element and send it to the audio renderer.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties and befor the label (if available).</param>
        /// <returns>The describing string of the element in the following form: [ROLE] [NAME] (addintionalText)</returns>
        public static string PlayElement(OoAccComponent element, string additionalText = "")
        {
            String text = GetElementAudioText(element, additionalText);
            if (!String.IsNullOrWhiteSpace(text))
            {
                audio.PlaySound(text);
            }
            return text;
        }

        /// <summary>
        /// Gets the audio text for the given element and send it immediately to the audio renderer.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties and befor the label (if available).</param>
        /// <returns>The describing string of the element in the following form: [ROLE] [NAME] [TITLE] (addintionalText) [TEXT]</returns>
        public static string PlayElementImmediately(OoShapeObserver element, string additionalText = "")
        {
            String text = GetElementAudioText(element, additionalText);
            if (!String.IsNullOrWhiteSpace(text))
            {
                audio.PlaySoundImmediately(text);
            }
            return text;
        }

        /// <summary>
        /// Gets the audio text for the given element and send it immediately to the audio renderer.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties and befor the label (if available).</param>
        /// <returns>The describing string of the element in the following form: [ROLE] [NAME] (addintionalText)</returns>
        public static string PlayElementImmediately(OoAccComponent element, string additionalText = "")
        {
            String text = GetElementAudioText(element, additionalText);
            if (!String.IsNullOrWhiteSpace(text))
            {
                audio.PlaySoundImmediately(text);
            }
            return text;
        }

        public static string PlayElementTitleAndDescriptionImmediately(OoShapeObserver element)
        {
            String text = GetElementAudioText(element);
            if (!String.IsNullOrWhiteSpace(element.Description)) text += " - " + LL.GetTrans("tangram.oomanipulation.element_speaker.description", element.Description);
            if (!String.IsNullOrWhiteSpace(text))
            {
                audio.PlaySoundImmediately(text);
            }
            return text;
        }

        /// <summary>
        /// Gets the audio text for the given element.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties and befor the label (if available).</param>
        /// <returns>A string that should describe the element in the following form: [ROLE] [NAME] [TITLE] (addintionalText) [TEXT]</returns>
        public static String GetElementAudioText(OoShapeObserver element, string additionalText = "")
        {
            String result = String.Empty;

            if (element != null && element.IsValid())
            {
                try
                {
                    bool isText = element.IsText;

                    #region default element handling

                    if (isText)
                    {
                        result = LL.GetTrans("tangram.oomanipulation.element_speaker.text_element", element.Text, element.Title);
                    }
                    else
                    {
                        //result += element.AccComponent.Role + " " + element.Name;
                        result += element.Name;

                        if (!String.IsNullOrWhiteSpace(element.Title))
                        {
                            result += (result.EndsWith(" ") ? "" : " ") + element.Title;
                        }
                    }

                    if (!String.IsNullOrEmpty(additionalText))
                    {
                        result += (result.EndsWith(" ") ? "" : " ") + additionalText;
                    }
                    else
                    {
                        result = result.Trim();
                    }

                    #endregion

                    #region object text

                    if (!isText && !String.IsNullOrWhiteSpace(element.Text))
                    {
                        result += " - " +LL.GetTrans("tangram.oomanipulation.element_speaker.text_value", element.Text);
                    }

                    #endregion

                    #region Group

                    if (element.IsGroup || element.HasChildren)
                    {
                        int cCount = element.ChildCount;
                        result += (result.EndsWith(" ") ? "" : " ") + "- "
                            + LL.GetTrans("tangram.oomanipulation.element_speaker.has_" + (cCount == 1 ? "one_child" : "children"), cCount.ToString());
                    }

                    #endregion

                    #region Group Member

                    if (element.IsGroupMember && element.Parent != null)
                    {
                        result += " - " + LL.GetTrans("tangram.oomanipulation.element_speaker.child_of", element.Parent.Name);
                    }

                    #endregion

                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Log(LogPriority.DEBUG, "OoElementSpeaker", "[ERROR] Can't play element: " + ex);
                }
            }

            return String.IsNullOrWhiteSpace(result) ? LL.GetTrans("tangram.oomanipulation.element_speaker.default") : result;
        }

        /// <summary>
        /// Gets the audio text for the given element.
        /// </summary>
        /// <param name="element">The element to get the text of.</param>
        /// <param name="additionalText">An addintional text, that is put behind the element properties</param>
        /// <returns>A string that should describe the element in the following form: [ROLE] [NAME] (addintionalText)</returns>
        public static String GetElementAudioText(OoAccComponent element, string additionalText = "")
        {
            String result = String.Empty;

            try
            {
                if (element != null)
                {
                    result += LL.GetTrans("tangram.oomanipulation.element_speaker.audio.element", element.Role.ToString(), element.Name.ToString(), additionalText.ToString());
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Log(LogPriority.DEBUG, "OoElementSpeaker", "[ERROR] Can't play element: " + ex);
            }

            return result;
        }

        /// <summary>
        /// Speak description of focused element.
        /// </summary>
        public static String GetElementDescriptionText(OoShapeObserver shape)
        {
            if (shape == null)
            {
                return null;
            }
            string desc = "";
            if (String.IsNullOrWhiteSpace(shape.Description))
            {
                desc = LL.GetTrans("tangram.oomanipulation.element_speaker.no_description");
            }
            else
            {
                desc = LL.GetTrans("tangram.oomanipulation.element_speaker.detail_description", shape.Name, shape.Description);
            }

            return desc;
        }

        #region ILocalizable

        static void SetLocalizationCulture(System.Globalization.CultureInfo culture)
        {
            if (LL != null) LL.SetStandardCulture(culture);
        }

        #endregion

    }
}