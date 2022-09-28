using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel.EtoEFlowChart
{
    public class FinalInspection
    {        
        public decimal PassRate { get; set; }
        public decimal SQR { get; set; }
        public decimal ChinaPassRate { get; set; }
        public decimal JapanPassRate { get; set; }
        public bool IsAllCnOrder{ get; set; }
        public bool IsAllJpOrder { get; set; }
    }
}
