using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class ReworkList_ViewModel : RFT_Inspection
    {
        public string ReworkNo { get; set; }

        public string POID { get; set; }

        public string Defect { get; set; }
    }
}
