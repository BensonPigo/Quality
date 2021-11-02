using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MockupWash_ViewModel : MockupWash
    {
        public bool ReturnResult { get; set; } = true;
        public bool SaveResult { get; set; } = true;
        public string ErrorMessage { get; set; }

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

        public List<SelectListItem> TestingMethod_Source { get; set; }

        [Display(Name = "T1 廠商 名稱")]
        public string T1SubconAbb { get; set; }

        [Display(Name = "T2 廠商 名稱")]
        public string T2SupplierAbb { get; set; }

        [Display(Name = "技術人員 名稱")]
        public string TechnicianName { get; set; }

        [Display(Name = "技術人員 ExtNo")]
        public string TechnicianExtNo { get; set; }

        [Display(Name = "業務 名稱")]
        public string MRName { get; set; }

        [Display(Name = "業務 ExtNo")]
        public string MRExtNo { get; set; }

        [Display(Name = "業務 信箱")]
        public string MRMail { get; set; }

        [Display(Name = "LastEditName")]
        public string LastEditName { get; set; }

        public string MethodDescription { get; set; }        

        [Display(Name = "報表電子簽章圖檔完整路徑")]
        public byte[] Signature { get; set; }

        public MockupWash_Request Request { get; set; }

        public List<MockupWash_Detail_ViewModel> MockupWash_Detail { get; set; }
    }

    public class MockupWash_Detail_ViewModel : MockupWash_Detail
    {
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

        public string ArtworkColorName { get; set; }
        public string FabricColorName { get; set; }
        public string LastUpdate { get; set; }
    }
}
