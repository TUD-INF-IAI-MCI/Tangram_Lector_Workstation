// ***********************************************************************
// Assembly         : TANGRAM OOo Draw Extention
// Author           : Admin
// Created          : 09-14-2012
//
// Last Modified By : Admin
// Last Modified On : 10-10-2012
// ***********************************************************************
// <copyright file="Any.cs" company="">
//     . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;

namespace tud.mci.tangram.models
{
    public static class Any
    {

        public static uno.Any Get(Object o)
        {
            if (o is uno.Any) return (uno.Any) o;
            var a = new uno.Any(o != null ? o.GetType() : (new Object()).GetType(), o);
            return a; 
        }

        public static uno.Any[] Get(Array i)
        {
            var a = new uno.Any[i.Length];
            for (int j = 0; j < i.Length; j++)
            {
                a[j] = Get(i.GetValue(j));
            }

            return a;

        }

        public static uno.Any GetAsOne(Array i)
        {
            return new uno.Any(i.GetType(),i);
        }        
    }
}
