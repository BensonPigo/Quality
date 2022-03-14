using System.Collections.Generic;
using System.Net.Mail;
using System.Web;

namespace DatabaseObject.RequestModel
{
    public class SendMail_Request
    {
        public string To { get; set; }
        public string CC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }

        /// <summary>
        /// 用於將產生報表自動夾上, EX:使用GetPDF回傳的 (路徑+檔案名稱) 填入此欄位
        /// </summary>
        public List<string> FileonServer { get; set; }

        /// <summary>
        /// 路徑+檔案名稱
        /// </summary>
        public List<HttpPostedFileBase> FileUploader { get; set; }

        /// <summary>
        /// 信件圖片附檔
        /// </summary>
        public AlternateView alternateView { get; set; }
    }
}
