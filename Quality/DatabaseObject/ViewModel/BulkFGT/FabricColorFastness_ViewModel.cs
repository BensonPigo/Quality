using DatabaseObject.ProductionDB;
using DatabaseObject.Public;
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
                    new SelectListItem(){ Text="60",Value="60"},
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
                    new SelectListItem(){ Text="1",Value="1"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="5",Value="5"},
                    new SelectListItem(){ Text="10",Value="10"},
                    new SelectListItem(){ Text="15",Value="15"},
                    new SelectListItem(){ Text="25",Value="25"},
                };
            }
            set { }
        }
        public List<SelectListItem> CycleTime_List
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="30",Value="30"},
                    new SelectListItem(){ Text="60",Value="60"},
                    new SelectListItem(){ Text="90",Value="90"},
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
                    new SelectListItem(){ Text="Fastness Machine",Value="Fastness Machine"},
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

        public string Name { get; set; }

        public string InspectionName { get; set; }
        public string MailSubject { get; set; }

        public BaseResult baseResult { get; set; }
    }

    public class Fabric_ColorFastness_Detail_ViewModel : BaseResult
    {
        public bool sentMail { get; set; }

        public string reportPath { get; set; }

        // 第三層表頭
        public ColorFastness_Result Main { get; set; }

        public List<Fabric_ColorFastness_Detail_Result> Details { get; set; }
    }

    public class Fabric_ColorFastness_Detail_Result : CompareBase
    {
        public string ID { get; set; }

        public string ColorFastnessGroup { get; set; }

        public string Roll { get; set; }

        public string Dyelot { get; set; }

        public string Result { get; set; }

        public string changeScale { get; set; }

        public string StainingScale { get; set; }

        public string Remark { get; set; }

        public string AddName { get; set; }

        public DateTime? AddDate { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

        public DateTime? SubmitDate { get; set; }

        public string ResultChange { get; set; }

        public string ResultStain { get; set; }

        public string AcetateScale { get; set; }
        public string ResultAcetate { get; set; }

        public string CottonScale { get; set; }
        public string ResultCotton { get; set; }
        public string NylonScale { get; set; }
        public string ResultNylon { get; set; }
        public string PolyesterScale { get; set; }
        public string ResultPolyester { get; set; }
        public string AcrylicScale { get; set; }
        public string ResultAcrylic { get; set; }
        public string WoolScale { get; set; }
        public string ResultWool { get; set; }


        public string Seq { get; set; }

        public string Refno { get; set; }

        public string SCIRefno { get; set; }

        public string ColorID { get; set; }

        public string LastUpdate { get; set; }

        public string SEQ1
        {
            get
            {
                if (string.IsNullOrEmpty(this.Seq))
                {
                    return string.Empty;
                }

                if (!this.Seq.Contains("-"))
                {
                    return string.Empty;
                }

                return this.Seq.Split('-')[0].Trim();
            }
        }

        public string SEQ2
        {
            get
            {
                if (string.IsNullOrEmpty(this.Seq))
                {
                    return string.Empty;
                }

                if (!this.Seq.Contains("-"))
                {
                    return string.Empty;
                }

                return this.Seq.Split('-')[1].Trim();
            }
        }
    }


    public class ColorFastness_Excel
    {
        public string ReportNo { get; set; }
        public DateTime? SubmitDate { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string StyleID { get; set; }
        public string POID { get; set; }
        public string Article { get; set; }
        public string ColorFastnessResult { get; set; }
        public int Temperature { get; set; }
        public int Cycle { get; set; }
        public int CycleTime { get; set; }
        public string Detergent { get; set; }
        public string Machine { get; set; }
        public string Drying { get; set; }
        public string SEQ1 { get; set; }
        public string SEQ2 { get; set; }
        public string SEQ 
        { 
            get
            {
                return string.Format("{0}-{1}", this.SEQ1, this.SEQ2);
            } 
        }
        public string Roll { get; set; }
        public string Dyelot { get; set; }               
        public string SCIRefno_Color { get; set; }
        public string ChangeScale { get; set; }
        public string AcetateScale { get; set; }
        public string CottonScale { get; set; }
        public string NylonScale { get; set; }
        public string PolyesterScale { get; set; }
        public string AcrylicScale { get; set; }
        public string WoolScale { get; set; }
        public string ResultChange { get; set; }        
        public string ResultAcetate { get; set; }        
        public string ResultCotton { get; set; }        
        public string ResultNylon { get; set; }        
        public string ResultPolyester { get; set; }        
        public string ResultAcrylic { get; set; }       
        public string ResultWool { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string StainingScale { get; set; }
        public string ResultStain { get; set; }
        public Byte[] Signature { get; set; }
        public string Checkby { get; set; }
        public Byte[] TestBeforePicture { get; set; }
        public Byte[] TestAfterPicture { get; set; }
    }
}
