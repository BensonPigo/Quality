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
        public string ReportNo_Query { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }

        // Result = Fail時，要收信的人
        public string ToAddress { get; set; }
        public string CcAddress { get; set; }

        public bool Result { get; set; }
        public string ErrorMessage { get; set; }

        public List<SelectListItem> ReportNo_Source { get; set; }

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

        public List<SelectListItem> Gender_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Male",Value="Male"},
                    new SelectListItem(){ Text="Female",Value="Female"},
                };
            }
            set { }
        }

        public List<SelectListItem> PullForceUnit_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Newton",Value="Newton"},
                    new SelectListItem(){ Text="IB",Value="IB"},
                };
            }
            set { }
        }

        public List<SelectListItem> StyleType_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Adults",Value="Adults"},
                    new SelectListItem(){ Text="Youth",Value="Youth"},
                };
            }
            set { }
        }

        public PullingTest_Result Detail { get; set; }
    }

    public class PullingTest_Result
    {
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

        public string TestDateText { get; set; }

        public decimal PullForce { get; set; }
        public string PullForceUnit { get; set; }     
        public string StyleType { get; set; }
        public string Inspector { get; set; }
        public string InspectorName { get; set; }
        public int Time { get; set; }
        public string FabricRefno { get; set; }
        public string AccRefno { get; set; }
        public string SnapOperator { get; set; }
        public string Remark { get; set; }
        public string LastEditName { get; set; }

        public string TestBeforePicture_Base64 { get; set; }
        public string TestAfterPicture_Base64 { get; set; }

        /// <summary>測試前的照片</summary>
        public Byte[] TestBeforePicture { get; set; }

        /// <summary>測試後的照片</summary>
        public Byte[] TestAfterPicture { get; set; }

        public decimal PullForce_Standard { get; set; }
        public int Time_Standard { get; set; }
        public string AddName { get; set; }
        public string EditName { get; set; }
        public string Gender { get; set; }
        public byte[] Signature { get; set; }
        public string MailSubject { get; set; }
        public string PullingTestApprover { get; set; }
        public string PullingTestApproverName { get; set; }
        public DateTime? ReceiveDate { get; set; }
        public DateTime? ReportDate { get; set; }
    }
}
