using System.Collections.Generic;

namespace DatabaseObject.RequestModel
{
    public class SendMail_Request
    {
        public string From { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }

        /// <summary>
        /// 路徑+檔案名稱
        /// </summary>
        public List<string> FileList { get; set; }
    }
}
