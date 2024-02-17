using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using DatabaseObject;
using BusinessLogicLayer.Interface;

namespace BusinessLogicLayer.Service
{
    public class ADAuthAPICheck : IADAuthAPICheck
    {
        private string aecKey = "SPORTSCITY-MISTOKENFO-RPORTALSYS-TEMUSERBUI-LDINLINEEE-FFFFFFFFFF-VERSION101";
        private string urlIssueTracking = "https://sciitjobtracking.sportscity.com.tw:8002/";

        public ADAuthAPICheck()
        {
#if DEBUG
            urlIssueTracking = "https://sciitjobtracking.sportscity.com.tw:8003/";
#endif
            this.CheckHttps(ref this.urlIssueTracking);
        }

        public BaseResult ADAuthByRegion(string region, string domainName)
        {
            BaseResult result = new BaseResult();
            string domainNameAES = this.AESEncrypt(domainName, out string iv);
            using (HttpClient clientADAuthByRegion = new HttpClient())
            {
                var clientContentADAuthByRegion = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("region", region),
                    new KeyValuePair<string, string>("domainName", domainNameAES),
                    new KeyValuePair<string, string>("iv", iv),
                });

                clientADAuthByRegion.BaseAddress = new Uri(this.urlIssueTracking);
                JObject jObj = new JObject();
                try
                {
                    HttpResponseMessage responseADAuthByRegion = clientADAuthByRegion.PostAsync($"api/ADAuthAPI/ADAuthByRegion", clientContentADAuthByRegion).Result;
                    string returnContent = responseADAuthByRegion.Content.ReadAsStringAsync().Result;

                    if (responseADAuthByRegion.StatusCode == HttpStatusCode.NotFound)
                    {
                        ErrorModel errorObject = JsonConvert.DeserializeObject<ErrorModel>(returnContent);
                        throw new Exception(errorObject.Message);
                    }

                    result.Result = true;
                }
                catch (Exception ex)
                {
                    result.Result = false;
                    result.ErrorMessage = ex.Message;
                    result.Exception = ex;
                }
            }

            return result;
        }

        private string AESEncrypt(string data, out string iv)
        {
            iv = this.GetRandomIV();
            var dataBytes = Encoding.UTF8.GetBytes(data);
            var keyBytes = Encoding.UTF8.GetBytes(this.aecKey.Substring(0, 16));
            var ivBytes = Encoding.UTF8.GetBytes(iv);
            using (Aes aes = Aes.Create())
            {
                aes.Key = keyBytes.ToArray();
                aes.IV = ivBytes.ToArray();
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                var encryptor = aes.CreateEncryptor().TransformFinalBlock(dataBytes, 0, dataBytes.Length);
                return Convert.ToBase64String(encryptor);
            }
        }

        private string GetRandomIV()
        {
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                byte[] byteRandom = new byte[16];
                rng.GetBytes(byteRandom);
                string iv = Convert.ToBase64String(byteRandom);
                return iv.Substring(0, 16);
            }
        }

        private void CheckHttps(ref string url)
        {
            try
            {
                HttpWebRequest obj = WebRequest.Create(url) as HttpWebRequest;
                obj.Method = "HEAD";
                HttpWebResponse obj2 = obj.GetResponse() as HttpWebResponse;
                obj2.Close();
                if (obj2.StatusCode != HttpStatusCode.OK)
                {
                    url = url.Replace("https://", "http://");
                }
            }
            catch
            {
                url = url.Replace("https://", "http://");
            }
        }
    }
}
