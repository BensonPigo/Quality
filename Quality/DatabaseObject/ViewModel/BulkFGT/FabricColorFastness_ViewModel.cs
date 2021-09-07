using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class FabricColorFastness_ViewModel: BaseResult
    {
        public string PoID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? CutInLine { get; set; }
        public DateTime? MinSciDelivery { get; set; }
        public DateTime? EarliestDate { get; set; }
        public DateTime? EarliestSCIDel { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal ArticlePercent { get; set; }
        public string ColorFastnessLaboratoryRemark { get; set; }

        public string CreateBy { get; set; }
        public string EditBy { get; set; }
        public string InspectorName { get; set; }

        public List<ColorFastness_Result> ColorFastness_MainList { get; set; }

        // 下拉式選單 List
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

        public List<SelectListItem> Temperature_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0",Value="0"},
                    new SelectListItem(){ Text="30",Value="30"},
                    new SelectListItem(){ Text="40",Value="40"},
                    new SelectListItem(){ Text="50",Value="50"},
                    new SelectListItem(){ Text="60",Value="65"},
                };
            }
            set { }
        }

        public List<SelectListItem> Cycle_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0",Value="0"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="5",Value="3"},
                    new SelectListItem(){ Text="10",Value="10"},
                    new SelectListItem(){ Text="15",Value="15"},
                    new SelectListItem(){ Text="25",Value="25"},
                };
            }
            set { }
        }

        public List<SelectListItem> Detergent_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Woolite",Value="Woolite"},
                    new SelectListItem(){ Text="Tide",Value="Tide"},
                    new SelectListItem(){ Text="AATCC",Value="AATCC"},
                    new SelectListItem(){ Text="ECE",Value="ECE"},
                };
            }
            set { }
        }

        public List<SelectListItem> Machine_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Top",Value="Top"},
                    new SelectListItem(){ Text="Load",Value="Load"},
                    new SelectListItem(){ Text="Front Load",Value="Front Load"},
                };
            }
            set { }
        }

        public List<SelectListItem> Drying_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Line Dry",Value="Line Dry"},
                    new SelectListItem(){ Text="Tumble Dry Low",Value="Tumble Dry Low"},
                    new SelectListItem(){ Text="Tumble Dry Mediumn",Value="Tumble Dry Mediumn"},
                    new SelectListItem(){ Text="Tumble Dry High",Value="Tumble Dry High"},
                };
            }
            set { }
        }
    }

    public class ColorFastness_Result : ColorFastness
    {
        public string LastUpdate { get; set; }

        public BaseResult baseResult { get; set; }
    }

    public class Fabric_ColorFastness_Detail_ViewModel
    {
        // 第三層表頭
        public ColorFastness Main { get; set; }

        public List<Fabric_ColorFastness_Detail_Result> Detail { get; set; }
    }

    public class Fabric_ColorFastness_Detail_Result : ColorFastness_Detail
    {
        public string Seq { get; set; }

        public string Refno { get; set; }

        public string SCIRefno { get; set; }

        public string ColorID { get; set; }

        public string LastUpdate { get; set; }
    }

    
}
