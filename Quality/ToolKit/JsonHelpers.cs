using Newtonsoft.Json;
using System;
using System.Xml.Linq;

namespace ToolKit
{
    public class JsonHelpers
    {
        public JsonHelpers()
        {
        }

        public static D Decoding<D>(string JSOData)
        {
            return JsonConvert.DeserializeObject<D>(JSOData);
        }

        public static XDocument DeserializeXNode(string JSON)
        {
            return JsonConvert.DeserializeXNode(JSON);
        }

        public static string Encoding<T>(T Item)
        {
            return JsonConvert.SerializeObject(Item);
        }

        public static object JsonToObject(string jsonString, object obj)
        {
            return JsonConvert.DeserializeObject(jsonString, obj.GetType());
        }

        public static string ObjectToJson(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }
    }
}
