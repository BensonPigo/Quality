using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class RFT_PicDuringDummyFitting_ViewModel : RFT_PicDuringDummyFitting
    {
        public bool Result { get; set; }
        public string ErrMsg { get; set; }
    }
}
