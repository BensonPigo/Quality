using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.RequestModel
{
    public class StyleResult_Request
    {
        public enum EnumCallType
        {
            PrintBarcode,
            StyleResult
        }

        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string MDivisionID { get; set; }
        public string InspectionTableName { get; set; }
        public string StyleUkey { get; set; }
        public string OrderTypeSerialKey { get; set; }

        public EnumCallType CallType { get; set; }
    }
}
