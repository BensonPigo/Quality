using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class RFT_Inspection_Measurement_ViewModel : RFT_Inspection_Measurement
    {
        public string SizeUnit { get; set; }

        public string Description { get; set; }

        public string Tol1 { get; set; }

        public string Tol2 { get; set; }

        public bool IsPatternMeas { get; set; }
    }
}
