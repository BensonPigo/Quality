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

        string GetInspectionTimes(string CustPONO);

        string UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid, string NewFinalInspectionID);

        void UpdateFinalInspectionByStep(FinalInspection finalInspection, string currentStep, string userID);

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

        IList<EndlineMoisture> GetEndlineMoisture();

        void UpdateMoisture(MoistureResult moistureResult);

        bool CheckMoistureExists(string finalInspectionID, string article, long? finalInspection_OrderCartonUkey);

        void DeleteMoisture(long ukey);

        void UpdateMeasurement(DatabaseObject.ViewModel.FinalInspection.Measurement measurement, string userID);

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
    }
}
