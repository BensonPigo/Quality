using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class RFT_OrderComments_ViewModel : RFT_OrderComments
    {
        public bool Result { get; set; }
        public string ErrMsg { get; set; }

        public string PMS_RFTCommentsDescription { get; set; }
    }
}
