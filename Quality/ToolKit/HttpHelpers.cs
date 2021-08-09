using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ToolKit
{
    public class HttpHelpers
    {
        public static HttpResponseMessage PostJsonDataHttpClient<T>(T Model, string targetUrl, Encoding encoding, string mediaType)
        {
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage();
            using (HttpClient httpClient = new HttpClient())
            {
                HttpContent stringContent = new StringContent(JsonHelpers.Encoding<T>(Model), encoding, mediaType);
                httpResponseMessage = httpClient.PostAsync(targetUrl, stringContent).Result;
            }
            return httpResponseMessage;
        }


        public static HttpResponseMessage GetJsonDataHttpClient(string targetUrl, string parameter)
        {
            HttpResponseMessage httpResponseMessage = new HttpResponseMessage();
            using (HttpClient httpClient = new HttpClient())
            {
                httpResponseMessage = httpClient.GetAsync(targetUrl + "?" + parameter).Result;
            }
            return httpResponseMessage;
        }
    }
}
