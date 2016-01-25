using System;
using System.Collections.Generic;
using System.Linq;

namespace tud.mci.tangram.TangramLector
{
    internal static class Utils
    {
        /// <summary>
        /// Check if two lists are equal.
        /// </summary>
        /// <param name="list1">The list1.</param>
        /// <param name="values">The values to check.</param>
        /// <returns><c>true</c> if the lists are the same, otherwise <c>false</c></returns>
        public static bool ListsAreEqual(List<String> list1, params String[] values)
        {
            if (values.Length != list1.Count) return false;
            return ListsAreEqual(list1, values.ToList());
        }
        /// <summary>
        /// Check if two lists are equal.
        /// </summary>
        /// <param name="list1">The list1.</param>
        /// <param name="list2">The values to check.</param>
        /// <returns><c>true</c> if the lists are the same, otherwise <c>false</c></returns>
        public static bool ListsAreEqual(List<String> list1, List<String> list2) { return ListsAreEqual<String>(list1, list2); }
        /// <summary>
        /// Check if two lists are equal.
        /// </summary>
        /// <typeparam name="T">the <c>Type</c> of the lists</typeparam>
        /// <param name="list1">The list1.</param>
        /// <param name="list2">The values to check.</param>
        /// <returns>
        /// 	<c>true</c> if the lists are the same, otherwise <c>false</c>
        /// </returns>
        public static bool ListsAreEqual<T>(List<T> list1, List<T> list2)
        {
            if (list1.Intersect(list2).ToList().Count == list1.Count) return true;
            return false;
        }


        public static float GetResoultion(float dpiX, float dpiY)
        {
            return (dpiX + dpiY) / 2;
        }
        
        [System.Runtime.InteropServices.DllImport("gdi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        private static extern int GetDeviceCaps(IntPtr hDC, int nIndex);
        private enum DeviceCap
        {
            /// <summary>
            /// Logical pixels inch in X
            /// </summary>
            LOGPIXELSX = 88,
            /// <summary>
            /// Logical pixels inch in Y
            /// </summary>
            LOGPIXELSY = 90

            // Other constants may be founded on pinvoke.net
        }

        /// <summary>
        /// Gets the screen dpi.
        /// </summary>
        /// <returns></returns>
        public static float GetScreenDpi()
        {
            System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();

            int dpiX = GetDeviceCaps(desktop, (int)DeviceCap.LOGPIXELSX);
            int dpiY = GetDeviceCaps(desktop, (int)DeviceCap.LOGPIXELSY);

            return GetResoultion((float)dpiX, (float)dpiY);
        }
    }
}
