using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.RequestModel
{
    public class PivotTransferRequest
    {
        public string InspectionType { get; set; }
        public string InspectionID { get; set; } = string.Empty;

        public string BaseUri { get; set; }

        public string RequestUri { get; set; }

        public Dictionary<string, string> Headers { get; set; } = new Dictionary<string, string>();
    }
}
