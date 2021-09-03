using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.RequestModel
{
    public class Measurement_Request
    {
        public string OrderID { get; set; }        
        public string OrderTypeID { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string Unit { get; set; }
        public string Article { get; set; }
        public List<string> Articles { get; set; } = new List<string>();
        public bool Result { get; set; }
        public string ErrMsg { get; set; }
        public string Factory { get; set; }
        public string FileName { get; set; }

    }
}
