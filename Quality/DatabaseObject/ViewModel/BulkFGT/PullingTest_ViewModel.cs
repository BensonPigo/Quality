using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class PullingTest_ViewModel
    {
        // 搜尋列
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }


        public List<string> ReportNo_Source { get; set; }

        public List<SelectListItem> Result_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }

        public List<SelectListItem> TestItem_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Snaps",Value="Snaps"},
                    new SelectListItem(){ Text="Buttons",Value="Buttons"},
                    new SelectListItem(){ Text="Rivet",Value="Rivet"},
                    new SelectListItem(){ Text="Eyelet",Value="Eyelet"},
                    new SelectListItem(){ Text="Sew on buttons",Value="Sew on buttons"},
                    new SelectListItem(){ Text="Tie cord",Value="Tie cord"},
                };
            }
            set { }
        }

        public PullingTest_Result Detail { get; set; }
    }


    public class PullingTest_Result
    {
        // 搜尋列
        public string ReportNo { get; set; }
        public string POID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string SizeCode { get; set; }
        public DateTime? TestDate { get; set; }
        public string Result { get; set; }
        public string TestItem { get; set; }
        public string PullForce { get; set; }
        public string PullForceUnit { get; set; }
        public int Time { get; set; }
        public string FabricRefno { get; set; }
        public string AccRefno { get; set; }
        public string SnapOperator { get; set; }
        public string Remark { get; set; }
        public string LastEditName { get; set; }

    }
}
