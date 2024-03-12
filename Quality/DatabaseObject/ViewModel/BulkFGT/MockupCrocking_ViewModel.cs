using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MockupCrocking_ViewModel : MockupCrocking
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

        public List<SelectListItem> Scale_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="1",Value="1"},
                    new SelectListItem(){ Text="1-2",Value="1-2"},
                    new SelectListItem(){ Text="2",Value="2"},
                    new SelectListItem(){ Text="2-3",Value="2-3"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="3-4",Value="3-4"},
                    new SelectListItem(){ Text="4",Value="4"},
                    new SelectListItem(){ Text="4-5",Value="4-5"},
                    new SelectListItem(){ Text="5",Value="5"},
                };
            }
            set { }
        }

        [Display(Name = "T1 廠商 名稱")]
        public string T1SubconAbb { get; set; }

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

        [Display(Name = "報表電子簽章圖檔完整路徑")]
        public byte[] Signature { get; set; }
        public string MailSubject { get; set; }

        public MockupCrocking_Request Request { get; set; }

        public List<MockupCrocking_Detail_ViewModel> MockupCrocking_Detail { get; set; }
    }

    public class MockupCrocking_Detail_ViewModel: MockupCrocking_Detail
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
