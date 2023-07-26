using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Web.Mvc;
using System.Web.Mvc.Properties;
using System.Web.Script.Serialization;

namespace Quality.Helper
{
    public sealed class MyJsonValueProviderFactory : ValueProviderFactory
    {
        private class EntryLimitedDictionary
        {
            private static int _maximumDepth = GetMaximumDepth();
            private readonly IDictionary<string, object> _innerDictionary;
            private int _itemCount;

            public EntryLimitedDictionary(IDictionary<string, object> innerDictionary)
            {
                this._innerDictionary = innerDictionary;
            }

            public void Add(string key, object value)
            {
                if (++this._itemCount > _maximumDepth)
                {
                    //throw new InvalidOperationException(MvcResources.JsonValueProviderFactory_RequestTooLarge);
                    throw new InvalidOperationException("itemCount is over maximumDepth");
                }
                this._innerDictionary.Add(key, value);
            }

            private static int GetMaximumDepth()
            {
                NameValueCollection appSettings = ConfigurationManager.AppSettings;
                if (appSettings != null)
                {
                    string[] values = appSettings.GetValues("aspnet:MaxJsonDeserializerMembers");
                    int result;
                    if (values != null && values.Length > 0 && int.TryParse(values[0], out result))
                    {
                        return result;
                    }
                }
                return int.MaxValue;
            }
        }

        private static void AddToBackingStore(EntryLimitedDictionary backingStore, string prefix, object value)
        {
            IDictionary<string, object> dictionary = value as IDictionary<string, object>;
            if (dictionary != null)
            {
                foreach (KeyValuePair<string, object> current in dictionary)
                {
                    AddToBackingStore(backingStore, MakePropertyKey(prefix, current.Key), current.Value);
                }
                return;
            }
            IList list = value as IList;
            if (list != null)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    AddToBackingStore(backingStore, MakeArrayKey(prefix, i), list[i]);
                }
                return;
            }
            backingStore.Add(prefix, value);
        }

        private static object GetDeserializedObject(ControllerContext controllerContext)
        {
            if (!controllerContext.HttpContext.Request.ContentType.StartsWith("application/json", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            StreamReader streamReader = new StreamReader(controllerContext.HttpContext.Request.InputStream);
            string text = streamReader.ReadToEnd();
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }
            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            // 解決這個問題：
            // 使用 JSON JavaScriptSerializer 序列化或還原序列化期間發生錯誤。字符串的長度超過在 maxJsonLength 屬性上設定的值。
            javaScriptSerializer.MaxJsonLength = int.MaxValue;
            // ----------------------------------------
            return javaScriptSerializer.DeserializeObject(text);
        }

        public override IValueProvider GetValueProvider(ControllerContext controllerContext)
        {
            if (controllerContext == null)
            {
                throw new ArgumentNullException("controllerContext");
            }
            object deserializedObject = GetDeserializedObject(controllerContext);
            if (deserializedObject == null)
            {
                return null;
            }
            Dictionary<string, object> dictionary = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            EntryLimitedDictionary backingStore = new EntryLimitedDictionary(dictionary);
            AddToBackingStore(backingStore, string.Empty, deserializedObject);
            return new DictionaryValueProvider<object>(dictionary, CultureInfo.CurrentCulture);
        }

        private static string MakeArrayKey(string prefix, int index)
        {
            return prefix + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
        }

        private static string MakePropertyKey(string prefix, string propertyName)
        {
            if (!string.IsNullOrEmpty(prefix))
            {
                return prefix + "." + propertyName;
            }
            return propertyName;
        }
    }
}