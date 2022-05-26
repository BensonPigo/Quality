using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.SampleRFT
{
    public class QueryInspectionBySP_ViewModel : ResultModelBase<QueryInspectionBySP>
    {
        public string SP { get; set; }
        public string CustPONO { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }

        public DateTime? InspDateStart { get; set; }
        public DateTime? InspDateEnd { get; set; }
    }
    public class QueryInspectionBySP
    {
        public long ID{ get; set; }
        public string SP { get; set; }
        public string CustPONO { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string Article { get; set; }
        public string SampleStage { get; set; }
        public string SewingLineID { get; set; }
        public string InspectionTimes { get; set; }
        public string Inspector { get; set; }
        public string Result { get; set; }
    }
}
