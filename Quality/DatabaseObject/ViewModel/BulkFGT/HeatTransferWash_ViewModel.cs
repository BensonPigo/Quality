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
        public List<SelectListItem> ReportNo_Source { get; set; }
        public HeatTransferWash_Result Main { get; set; }
        public List<HeatTransferWash_Detail_Result> Details { get; set; }
    }

    public class HeatTransferWash_Result
    {
        public string ReportNo { get; set; }
        public string POID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string Line { get; set; }
        public string Machine { get; set; }
        public bool IsTeamwear { get; set; }
        public DateTime? ReportDate { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public int Temperature { get; set; }
        public int Time { get; set; }
        public decimal Pressure { get; set; }
        public string PeelOff { get; set; }
        public int Cycles { get; set; }
        public int TemperatureUnit { get; set; }

        public string TestBeforePicture_Base64 { get; set; }
        public string TestAfterPicture_Base64 { get; set; }

        /// <summary>測試前的照片</summary>
        public byte[] TestBeforePicture { get; set; }

        /// <summary>測試後的照片</summary>
        public byte[] TestAfterPicture { get; set; }

        public string AddNameText { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }

        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd")}-{this.EditNameText}" : string.Empty;
            }
        }

    }
    public class HeatTransferWash_Detail_Result
    {
        public string ReportNo { get; set; }
        public long Ukey { get; set; }
        public string FabricRefNo { get; set; }
        public string HTRefNo { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd")}-{this.EditNameText}" : string.Empty;
            }
        }

    }
}
