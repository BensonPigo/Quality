using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.SampleRFT
{
    public class InspectionDefectList_ViewModel : ResultModelBase<InspectionDefectList_Result>
    {

        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string SampleStage { get; set; }
        public string SeasonID { get; set; }
    }

    public class InspectionDefectList_Result
    {


        //這四個不需要顯示在畫面上
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string SampleStage { get; set; }
        public string SeasonID { get; set; }

        public string Article { get; set; }
        public string Size { get; set; }
        public bool HasImage { get; set; }
        public DateTime? InspectionDate { get; set; }
        public string DefectType { get; set; }
        public string DefectCode{ get; set; }
        public string AreaCodes { get; set; }
        
        public byte[] DefectPicture { get; set; }
    }
}
