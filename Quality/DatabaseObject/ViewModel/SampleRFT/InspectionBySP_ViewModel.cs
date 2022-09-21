using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.SampleRFT
{
    public class InspectionBySP_ViewModel : ResultModelBase<InspectionBySP_SearchResult>
    {
        public long ID { get; set; }
        public string OrderID { get; set; }
        public string CustPONo { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string FactoryID { get; set; }
    }

    public class InspectionBySP_SearchResult
    {
        public string OrderID { get; set; }
        public string CustPONo { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string Article { get; set; }
        public string SampleStage { get; set; }
    }

    public class SampleRFTInspection
    {
        public long ID { get; set; }
        public long StyleUkey { get; set; }
        public string StyleID { get; set; }
        public string OrderID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string FactoryID { get; set; }
        public string InspectionStage { get; set; }
        public int InspectionTimes { get; set; }
        public string InspectionTimesText { get; set; }
        public string InspectionStep { get; set; }
        public string SewingLineID { get; set; }
        public string SewingLine2ndID { get; set; }
        public string QCInCharge { get; set; }
        public string AddName { get; set; }
        public string CFTNameText { get; set; }
        public string MReMail { get; set; }
        public string CCMail { get; set; }
        public string MDivisionID { get; set; }
        public string BrandFTYCode { get; set; }
        public int OrderQty { get; set; }
        public int PassQty { get; set; }
        public DateTime? InspectionDate { get; set; }
        public DateTime? SubmitDate { get; set; }

        #region AQL Plan
        public string AQLPlan { get; set; }
        public int? SampleSize { get; set; }
        public int? AcceptQty { get; set; }
        public int RejectQty { get; set; }
        public int BAQty { get; set; }
        public long AcceptableQualityLevelsUkey { get; set; }
        #endregion
        public string Result { get; set; }

        public bool CheckFabricApproval { get; set; }
        public bool CheckMetalDetection { get; set; }
        public bool CheckSealingSampleApproval { get; set; }
        public bool CheckColorShade { get; set; }
        public bool CheckHandfeel { get; set; }
        public bool CheckAppearance { get; set; }
        public bool CheckPrintEmbDecorations { get; set; }
        public bool CheckFiberContent { get; set; }
        public bool CheckCareInstructions { get; set; }
        public bool CheckDecorativeLabel { get; set; }
        //public bool CheckAdicomLabel { get; set; }
        public bool CheckCountryofOrigin { get; set; }
        public bool CheckSizeKey { get; set; }
        //public bool Check8FlagLabel { get; set; }
        public bool CheckAdditionalLabel { get; set; }
        //public bool CheckShippingMark { get; set; }
        public bool CheckPolytagMarketing { get; set; }
        //public bool CheckColorSizeQty { get; set; }
        public bool CheckCareLabel { get; set; }
        public bool CheckSecurityLabel { get; set; }
        public bool CheckOuterCarton { get; set; }
        public bool CheckPackingMode { get; set; }
        public bool CheckHangtag { get; set; }
        public bool CheckHT { get; set; }
        public bool CheckEMB { get; set; }

        public string WorkNo { get; set; }
        public string POID { get; set; }
    }


    public class InspectionBySP_Setting
    {
        public long ID { get; set; }        
        public string OrderStyleUnit { get; set; }
        #region Basic
        public string InspectionStage { get; set; } = string.Empty;
        public string InspectionTimes { get; set; } = string.Empty;
        public string SewingLineID { get; set; } = string.Empty;
        public string SewingLine2ndID { get; set; } = string.Empty;
        public string TopSewingLineID { get; set; } = string.Empty;
        public string InnerSewingLineID { get; set; } = string.Empty;
        public string BottomSewingLineID { get; set; } = string.Empty;
        public string OuterSewingLineID { get; set; } = string.Empty;
        public string QCInCharge { get; set; } = string.Empty;
        public DateTime? InspectionDate { get; set; }

        public string Dest { get; set; }
        public bool VasShas { get; set; }

        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        #endregion

        #region AQL Plan
        public string AQLPlan { get; set; }
        public int? SampleSize { get; set; }
        public int? AcceptQty { get; set; }
        public long AcceptableQualityLevelsUkey { get; set; }
        #endregion

        #region 訂單相關資訊

        public string OrderID { get; set; }
        public string CustPONo { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string Article { get; set; }
        public string ComboType { get; set; }
        
        public string FactoryID { get; set; }
        public string Model { get; set; }        
        public string SampleStage { get; set; }
        public int OrderQty { get; set; }

        #endregion

        #region 下拉選單
        public List<SelectListItem> InspectionStageList { get; set; } = new List<SelectListItem>()
            {
                new SelectListItem() { Text = "100%", Value = "100%" },
                new SelectListItem() { Text = "AQL", Value = "AQL" },

            };
        public List<SelectListItem> InspectionTimesList { get; set; } = new List<SelectListItem>()
            {
                new SelectListItem() { Text = "1/Final", Value = "1" },
                new SelectListItem() { Text = "2/Final", Value = "2" },
                new SelectListItem() { Text = "3/Final", Value = "3" },

            };
        public List<SelectListItem> SewingLineList { get; set; }
        public List<SelectListItem> QC_InChargeList { get; set; }
        public List<SelectListItem> AQLPlanList { get; set; } = new List<SelectListItem>()
            {
                new SelectListItem(){Text="",Value=""},
                new SelectListItem(){Text="1.0 Level",Value="1.0 Level"},
                new SelectListItem(){Text="1.5 Level",Value="1.5 Level"},
                new SelectListItem(){Text="2.5 Level",Value="2.5 Level"},
            };
        public List<SelectSewing> SelectedSewing { get; set; }
        public List<Select_QC_InCharge> Select_QC_InCharge { get; set; }
        public List<AcceptableQualityLevels> AcceptableQualityLevels { get; set; }
        #endregion

        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class SelectStage
    {
        public bool Selected { get; set; }
        public string StageName { get; set; }
    }
    public class SelectTimes
    {
        public bool Selected { get; set; }
        public string Times { get; set; }
    }
    public class SelectSewing
    {
        public bool Selected { get; set; }
        public string SewingLine { get; set; }
    }
    public class Select_QC_InCharge
    {
        public bool Selected { get; set; }
        public string Pass1ID { get; set; }
    }

    public class InspectionBySP_CheckList : SampleRFTInspection
    {
        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }
    public class InspectionBySP_Measurement : SampleRFTInspection
    {
        public List<SelectListItem> ListArticle { get; set; }

        public List<ArticleSize> ListSize { get; set; }

        public List<SelectListItem> ListProductType { get; set; }


        public string SelectedArticle { get; set; }
        public string SelectedSize { get; set; }
        public string SelectedProductType { get; set; }
        public string SizeUnit { get; set; }

        public List<MeasurementItem> ListMeasurementItem { get; set; }
        public List<RFT_Inspection_Measurement> ListRFTMeasurementItem { get; set; }

        public List<byte[]> ImageList { get; set; }

        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }


    public class MeasurementItem
    {
        public string Description { get; set; }
        public string SizeSpec { get; set; }
        public string Tol1 { get; set; }
        public string Tol2 { get; set; }
        public string Size { get; set; }
        public string ResultSizeSpec { get; set; } = string.Empty;
        public string Code { get; set; }
        public long MeasurementUkey { get; set; }
        public bool CanEdit { get; set; }
    }
    public class MeasurementViewItem
    {
        public string Article { get; set; }
        public string Size { get; set; }
        public string ProductType { get; set; }
        public string MeasurementDataByJson { get; set; }
    }
    public class ArticleSize
    {
        public string Article { get; set; }
        public string SizeCode { get; set; }
    }
    public class InspectionBySP_AddDefect : SampleRFTInspection
    {
        public List<SelectListItem> ResposiblityList { get; set; } = new List<SelectListItem>()
            {
                new SelectListItem() { Text = "MR", Value = "MR" },
                new SelectListItem() { Text = "Sub-corn", Value = "Sub-corn" },
                new SelectListItem() { Text = "Materials", Value = "Materials" },
                new SelectListItem() { Text = "Sewing Line", Value = "Sewing Line" },
                new SelectListItem() { Text = "Others-CCDA Issue/Design Issue", Value = "Others-CCDA Issue/Design Issue" },

            };
        public List<SampleRFTInspection_Summary> ListDefectItem { get; set; }
        public List<Area> Areas { get; set; }
        public int MaxRejectQty { get; set; }
        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class SampleRFTInspection_Summary
    {
        public Int64 RowIndex { get; set; }
        public long ID { get; set; } // SampleRFTInspection.ID
        public long UKey { get; set; } // SampleRFTInspection_Detail.UKey
        public string GarmentDefectTypeID { get; set; } // = Production.dbo.GarmentDefectType.ID
        public string DefectType { get; set; }// = Production.dbo.GarmentDefectType.Description
        public string DefectTypeDesc { get; set; } // = ID +"-" + Description

        public string GarmentDefectCodeID { get; set; }// = Production.dbo.GarmentDefectCode.ID
        public string DefectCode { get; set; } // = Production.dbo.GarmentDefectCode.Description
        public string DefectCodeDesc { get; set; }// = ID +"-" + Description
        public bool HasImage { get; set; }

        public string AreaCodes { get; set; }
        public int Qty { get; set; }
        public string Responsibility { get; set; }

        public List<SelectListItem> Images_Source { get; set; }
        public List<DefectImage> Images { get; set; }

        public bool Result { get; set; }
        public string ErrMsg { get; set; }
    }

    public class DefectImage
    {
        public string LoginToken { get; set; }
        public string GarmentDefectCodeID { get; set; }
        public long ImageUKey { get; set; }
        public long Seq { get; set; }
        public long SampleRFTInspectionDetailUKey { get; set; }
        public byte[] Image { get; set; }
        public byte[] TempImage { get; set; }
    }

    /// <summary>
    /// DB專用
    /// </summary>
    public class SampleRFTInspection_Detail
    {
        public long SampleRFTInspectionUkey { get; set; }
        public long UKey { get; set; }
        public string DefectCode { get; set; }

        public string AreaCode { get; set; }

        public string GarmentDefectTypeID { get; set; }

        public string GarmentDefectCodeID { get; set; }

        public string Responsibility { get; set; }

        public int Qty { get; set; }

        public List<byte[]> ListDefectImage { get; set; } = new List<byte[]>();

    }


    public class InspectionBySP_BA : SampleRFTInspection
    {
        public List<BACriteriaItem> ListBACriteria { get; set; }
        public int MaxBAQty { get; set; }

        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }
    public class BACriteriaItem
    {
        public long ID { get; set; } // SampleRFTInspection.ID
        public long Ukey { get; set; } // SampleRFTInspection_NonBACriteria.Ukey
        public string BACriteria { get; set; } // SampleRFTInspection_NonBACriteria.BACriteria
        public string BACriteriaDesc { get; set; }
        public int? Qty { get; set; }  // SampleRFTInspection_NonBACriteria.Qty
        public bool HasImage { get; set; }
        public List<SelectListItem> Images_Source { get; set; }
        public List<BAImage> Images { get; set; } = new List<BAImage>();
        public byte[] TempImage { get; set; }
        public string TempRemark { get; set; }
        public Int64 RowIndex { get; set; }

        public bool Result { get; set; }
        public string ErrMsg { get; set; }
    }
    public class BAImage
    {
        public string LoginToken { get; set; }
        public string BACriteria { get; set; }
        public long ImageUKey { get; set; }
        public long SampleRFTInspection_NonBACriteriaUKey { get; set; }
        public long Seq { get; set; }
        public byte[] Image { get; set; }
        public byte[] TempImage { get; set; }
    }

    public class SampleRFTInspection_NonBACriteria
    {
        public long SampleRFTInspectionUkey { get; set; }
        public long UKey { get; set; }
        public string BACriteria { get; set; }
        public int Qty { get; set; }
        public List<byte[]> ListBAImage { get; set; } = new List<byte[]>();

    }

    public class InspectionBySP_DummyFit : SampleRFTInspection
    {
        public List<SelectListItem> ArticleList { get; set; }
        public List<SelectListItem> SizeList { get; set; }
        public List<ArticleSize> ArticleSizeList { get; set; }
        public List<DummyFitImage> DetailList { get; set; }

        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }
    public class DummyFitImage
    {
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string Size { get; set; }
        public string FrontImageName 
        { 
            get 
            {
                return $"{this.OrderID}_{this.Article}_{this.Size}_Front.png";
            } 
        }
        public string LeftImageName
        {
            get
            {
                return $"{this.OrderID}_{this.Article}_{this.Size}_Left.png";
            }
        }
        public string RightImageName
        {
            get
            {
                return $"{this.OrderID}_{this.Article}_{this.Size}_Right.png";
            }
        }
        public string BackImageName
        {
            get
            {
                return $"{this.OrderID}_{this.Article}_{this.Size}_Back.png";
            }
        }
        public byte[] Front { get; set; }
        public byte[] Side { get; set; }
        public byte[] Back { get; set; }
        public byte[] Left { get; set; }
        public byte[] Right { get; set; }
    }

    public class InspectionBySP_Others : SampleRFTInspection
    {
        public string Inspector { get; set; }
        public string InspectorResult { get; set; }
        public string Action { get; set; }
        //public string OrderID { get; set; }
        //public string StyleID { get; set; }
        //public string BrandID { get; set; }
        //public string SeasonID { get; set; }
        public string SampleStage { get; set; }
        public List<CFTComments_Result> DataList { get; set; }

        public List<SelectListItem> SamePOIDList { get; set; }
        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }
    }
    public class CFTComments_Result
    {
        public string PMS_RFTCommentsID { get; set; }
        public string Description { get; set; }
        public string Comnments { get; set; }
    }
}
