using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;
using System.Data;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IFinalInspectionProvider
    {
        FinalInspection GetFinalInspection(string FinalInspectionID);

        /// <summary>
        /// 取得可用的關卡
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="CustPONO"></param>
        /// <returns></returns>
        List<FinalInspection_Step> GetAllStep(string FinalInspectionID, string CustPONO);

        /// <summary>
        /// 取得Final單子的上一關or 下一關 or 目前關卡
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        List<FinalInspection_Step> GetAllStepByAction(string FinalInspectionID, FinalInspectionSStepAction action);

        /// <summary>
        /// 將Final單子前往下一關/回到上一關
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="action">前往下一關/回到上一關</param>
        /// <returns></returns>
        void UpdateStepByAction(string FinalInspectionID, string UserID, FinalInspectionSStepAction action);


        string GetInspectionTimes(string CustPONO);

        string UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid, string NewFinalInspectionID);

        /// <summary>
        /// 部分功能Back/Next按鈕按下時，要存檔的東西(Remark之類的)
        /// </summary>
        /// <param name="finalInspection"></param>
        /// <param name="currentStep"></param>
        /// <param name="userID"></param>
        void UpdateStepInfo(FinalInspection finalInspection, string currentStep, string userID);

        /// <summary>
        /// 最後一關按下Submit按鈕
        /// </summary>
        /// <param name="finalInspection"></param>
        /// <param name="userID"></param>
        void SubmitFinalInspection(FinalInspection finalInspection, string userID);

        //IList<byte[]> GetFinalInspectionDefectImage(long FinalInspection_DetailUkey);
        IList<ImageRemark> GetFinalInspectionDetail(long FinalInspection_DetailUkey);

        Dictionary<string, byte[]> GetFinalInspectionDefectImage(string FinalInspectionID);

        Dictionary<string, byte[]> GetInlineInspectionDefectImage(string InspectionID);

        Dictionary<string, byte[]> GetEndLineInspectionDefectImage(string InspectionID);

        void UpdateFinalInspectionDetail(AddDefect addDefect, string UserID);

        IList<BACriteriaItem> GetBeautifulProductAuditForInspection(string finalInspectionID);

        void UpdateBeautifulProductAudit(BeautifulProductAudit addDefect, string UserID);

        //List<byte[]> GetBACriteriaImage(long FinalInspection_NonBACriteriaUkey);
        IList<ImageRemark> GetBA_DetailImage(long FinalInspection_NonBACriteriaUkey);


        IList<CartonItem> GetMoistureListCartonItem(string finalInspectionID);

        IList<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID);

        /// <summary>
        /// 根據品牌尋找EndlineMoisture設定
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="BrandID"></param>
        /// <returns></returns>
        IList<EndlineMoisture> GetEndlineMoistureByBrand(string FinalInspectionID, string BrandID);

        /// <summary>
        /// 取得EndlineMoisture的預設值
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="BrandID"></param>
        /// <returns></returns>
        IList<EndlineMoisture> GetEndlineMoistureDefault();
        void UpdateMoisture(MoistureResult moistureResult);

        bool CheckMoistureExists(string finalInspectionID, string article, long? finalInspection_OrderCartonUkey);

        void DeleteMoisture(long ukey);

        void UpdateMeasurement(DatabaseObject.ViewModel.FinalInspection.ServiceMeasurement measurement, string userID);

        IList<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID);
        DataTable GetMeasurement(string finalInspectionID, string article, string size, string productType);

        //void UpdateFinalInspection_OtherImage(string finalInspectionID, List<byte[]> images);
        void UpdateFinalInspection_OtherImage(string finalInspectionID, List<OtherImage> images);
        DataTable GetReportMailInfo(string finalInspectionID);

        IList<QueryFinalInspection> GetFinalinspectionQueryList_Default(QueryFinalInspection_ViewModel request);
        IList<QueryFinalInspection> GetFinalinspectionQueryList(QueryFinalInspection_ViewModel request);

        DataTable GetQueryReportInfo(string finalInspectionID);

        IList<FinalInspection_OrderCarton> GetListCartonInfo(string finalInspectionID);

        IList<SelectOrderShipSeq> GetListShipModeSeq(string finalInspectionID);

        IList<OtherImage> GetOthersImageList(string finalInspectionID);
        string GetNewFinalInspectionID(string factoryID);
        IList<DatabaseObject.ProductionDB.System> GetSystem();

        DataSet GetPivot88(string ID);

        DataSet GetEndInlinePivot88(string ID, string inspectionType);

        List<string> GetPivot88FinalInspectionID(string finalInspectionID);
        
        List<string> Get_FinalInspectionID_BrandID(string finalInspectionID);

        string Get_FinalInspectionID_Top1_OrderID(string finalInspectionID);

        List<string> GetPivot88EndLineInspectionID(string inspectionID);

        List<string> GetPivot88InlineInspectionID(string inspectionID);

        void UpdateIsExportToP88(string ID, string inspectionType);

        void ExecImp_EOLInlineInspectionReport();

        BaseResult UpdateJunk(string ID);

        List<FinalInspectionBasicGeneral> GetGeneralByBrand(string FinalInspectionID, string BrandID);
        List<FinalInspectionBasicCheckList> GetCheckListByBrand(string FinalInspectionID, string BrandID);
        void UpdateGeneral(FinalInspectionGeneral General);
        void UpdateCheckList(FinalInspectionCheckList CheckList);

        /// <summary>
        /// 取得濕度標準
        /// </summary>
        /// <param name="finalInspectionID"></param>
        /// <returns></returns>
        List<FinalInspectionMoistureStandard> GetMoistureStandardSetting(string finalInspectionID);

        List<FinalInspectionBasicGeneral> GetAllGeneral();
        List<FinalInspectionBasicCheckList> GetAllCheckList();
        int GetAvailableQty(string FinalInspectionID);
    }
}
