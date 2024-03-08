using System.Collections.Generic;
using System.Web;

namespace DatabaseObject.RequestModel
{
    public class MockupFailMail_Request
    {
        public string To { get; set; }
        public string CC { get; set; }
        public string ReportNo { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<HttpPostedFileBase> Files { get; set; }
    }
}
