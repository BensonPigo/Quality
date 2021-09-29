using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ResultModel.FinalInspection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class PoSelect : ResultModelBase<PoSelect_Result>
    {
        public string SP { get; set; }
        public string CustPONO { get; set; }
        public string StyleID { get; set; }
        public DateTime? SciDeliveryStart { get; set; } = null;
        public DateTime? SciDeliveryEnd { get; set; } = null;
        public string InspectionResult { get; set; }
    }
}
