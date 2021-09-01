using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel.FinalInspection
{
    public class PoSelect_Result
    {
        [Display(Name = "ID")]
        public string ID { get; set; }

        [Display(Name = "POID")]
        public string POID { get; set; }
        [Display(Name = "BrandID")]
        public string BrandID { get; set; }
        [Display(Name = "StyleID")]
        public string StyleID { get; set; }
        [Display(Name = "SeasonID")]
        public string SeasonID { get; set; }
        [Display(Name = "Qty")]
        public int? Qty { get; set; }
        [Display(Name = "Selected")]
        public bool Selected { get; set; }
    }
}
