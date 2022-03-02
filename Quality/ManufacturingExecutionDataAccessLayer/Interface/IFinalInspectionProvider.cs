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

        IList<byte[]> GetFinalInspectionDefectImage(long FinalInspection_DetailUkey);
        IList<ImageRemark> GetFinalInspectionDetail(long FinalInspection_DetailUkey);

        void UpdateFinalInspectionDetail(AddDefect addDefect, string UserID);

        IList<BACriteriaItem> GetBeautifulProductAuditForInspection(string finalInspectionID);

        void UpdateBeautifulProductAudit(BeautifulProductAudit addDefect, string UserID);

        List<byte[]> GetBACriteriaImage(long FinalInspection_NonBACriteriaUkey);
        IList<ImageRemark> GetBA_DetailImage(long FinalInspection_NonBACriteriaUkey);

        List<byte[]> GetOthersImage(string finalInspectionID);

        IList<CartonItem> GetMoistureListCartonItem(string finalInspectionID);

        IList<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID);

        IList<EndlineMoisture> GetEndlineMoisture();

        void UpdateMoisture(MoistureResult moistureResult);

        bool CheckMoistureExists(string finalInspectionID, string article, long? finalInspection_OrderCartonUkey);

        void DeleteMoisture(long ukey);

        void UpdateMeasurement(DatabaseObject.ViewModel.FinalInspection.Measurement measurement, string userID);

        IList<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID);
        DataTable GetMeasurement(string finalInspectionID, string article, string size, string productType);

        void UpdateFinalInspection_OtherImage(string finalInspectionID, List<byte[]> images);

        DataTable GetReportMailInfo(string finalInspectionID);

        IList<QueryFinalInspection> GetFinalinspectionQueryList(QueryFinalInspection_ViewModel request);

        DataTable GetQueryReportInfo(string finalInspectionID);

        IList<FinalInspection_OrderCarton> GetListCartonInfo(string finalInspectionID);

        IList<SelectOrderShipSeq> GetListShipModeSeq(string finalInspectionID);

        IList<OtherImage> GetOthersImageList(string finalInspectionID);
        string GetNewFinalInspectionID(string factoryID);
        IList<DatabaseObject.ProductionDB.System> GetSystem();

        DataSet GetPivot88(string ID);

        List<string> GetPivot88FinalInspectionID(string finalInspectionID);

        void UpdateIsExportToP88(string ID);
    }
}
