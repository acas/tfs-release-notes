using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReleaseNotesWeb.Utility
{
    public static class Utility
    {
        /// <summary>
        /// Gets the value with the specified key as a specific type T. If the key
        /// doesn't exist, the default value for T is returned.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobject"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static T GetValueOrDefault<T>(this Newtonsoft.Json.Linq.JObject jobject, string key, T replacementValue = default(T))
        {
            try
            {
                if (replacementValue != null)
                {
                    return jobject[key] != null ? jobject[key].ToObject<T>() : replacementValue;
                }
                else
                {
                    return jobject[key] != null ? jobject[key].ToObject<T>() : default(T);
                }
            }
            catch (ArgumentException)
            {
                if (replacementValue != null)
                {
                    return replacementValue;
                } 
                else 
                {
                    return default(T);
                }
            }
        }
    }
}