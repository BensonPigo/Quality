using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class InspectionSave_ViewModel
    {
        public bool Result { get; set; }

        public string ErrMsg { get; set; }

        public string StyleID { get; set; }

        public RFT_Inspection rft_Inspection { get; set; }

        public List<RFT_Inspection_Detail> fT_Inspection_Details { get; set; }
    }
}
