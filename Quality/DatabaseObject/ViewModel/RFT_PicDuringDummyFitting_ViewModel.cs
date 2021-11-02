using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class RFT_PicDuringDummyFitting_ViewModel : ResultModelBase<RFT_PicDuringDummyFitting>
    {

        // for SampleRFT.PicturesDummy 頁面使用
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string OrderTypeID { get; set; }
        public string SeasonID { get; set; }

    }
}
