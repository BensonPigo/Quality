using System.Net.Http;
using System.Text;
using ToolKit;

namespace Library
{
    public class Http
    {
        public static HttpResponseMessage PostData<T>(T Model, string targetUrl)
        {
            return HttpHelpers.PostJsonDataHttpClient<T>(Model, targetUrl, Encoding.UTF8, "application/json");
        }

        public static HttpResponseMessage GetData(string targetUrl, string parameter)
        {
            return HttpHelpers.GetJsonDataHttpClient(targetUrl, parameter);
        }
    }
}
