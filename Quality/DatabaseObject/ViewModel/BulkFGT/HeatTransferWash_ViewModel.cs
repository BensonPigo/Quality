using DatabaseObject.Public;
using DatabaseObject.RequestModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class HeatTransferWash_ViewModel : BaseResult
    {
        public HeatTransferWash_Request Request { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> ArtworkType_Source { get; set; }
        public List<SelectListItem> Cycles_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0",Value="0"},
                    new SelectListItem(){ Text="1",Value="1"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="5",Value="5"},
                    new SelectListItem(){ Text="10",Value="10"},
                    new SelectListItem(){ Text="15",Value="15"},
                    new SelectListItem(){ Text="25",Value="25"},
                };
            }
        }
        public List<SelectListItem> TemperatureUnit_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0",Value="0"},
                    new SelectListItem(){ Text="30",Value="30"},
                    new SelectListItem(){ Text="40",Value="40"},
                    new SelectListItem(){ Text="50",Value="50"},
                    new SelectListItem(){ Text="60",Value="60"},
                };
            }
        }
        public List<SelectListItem> PassFail_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
        }
        public HeatTransferWash_Result Main { get; set; }
        public List<HeatTransferWash_Detail_Result> Details { get; set; }
    }

    public class HeatTransferWash_Result
    {
        public string ReportNo { get; set; }
        public string OrderID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string Line { get; set; }
        public string Machine { get; set; }
        public bool Teamwear { get; set; }
        public DateTime? ReceivedDate { get; set; }
        public DateTime? ReportDate { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string Status { get; set; }
        public string ArtworkTypeID { get; set; }
        public string ArtworkTypeID_FullName { get; set; }
        //public int Temperature { get; set; }
        //public int Time { get; set; }
        //public decimal Pressure { get; set; }
        //public string PeelOff { get; set; }
        //public int Cycles { get; set; }
        //public int TemperatureUnit { get; set; }

        public string TestBeforePicture_Base64 { get; set; }
        public string TestAfterPicture_Base64 { get; set; }

        /// <summary>測試前的照片</summary>
        public byte[] TestBeforePicture { get; set; }

        /// <summary>測試後的照片</summary>
        public byte[] TestAfterPicture { get; set; }

        public byte[] Signature { get; set; }
        public string AddNameText { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }

        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public string LastEditText { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }

        public string MRHandleEmail{ get; set; }
    }
    public class HeatTransferWash_Detail_Result : CompareBase
    {
        public string ReportNo { get; set; }
        public long Ukey { get; set; }
        public string FabricRefNo { get; set; }
        public string HTRefNo { get; set; }

        public int Temperature { get; set; }
        public int Time { get; set; }
        public int SecondTime { get; set; }
        public decimal Pressure { get; set; }
        public string PeelOff { get; set; }
        public int Cycles { get; set; }
        public int TemperatureUnit { get; set; }

        public string Result { get; set; }
        public string Remark { get; set; }
        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd")}-{this.EditName}" : string.Empty;
            }
        }

    }
}
