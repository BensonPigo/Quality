using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFinalInspFromPMSProvider
    {
        IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForSetting();
        IList<AcceptableQualityLevelsProList> GetAcceptableQualityLevelsProListForSetting(string BrandID ,long ProUkey);
        IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForMeasurement();
        IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID);
        IList<SelectedPO> GetSelectedPOForInspection(string finalInspectionID);
        IList<SelectCarton> GetSelectedCartonForSetting(List<string> listOrderID);
        IList<SelectCarton> GetSelectedCartonForSetting(string finalInspectionID);
        IList<FinalInspectionDefectItem> GetFinalInspectionDefectItems(string finalInspectionID);
        IList<FinalInspection_DefectDetail> GetFinalInspection_DefectDetails(string finalInspectionID, long ProUkey);
        IList<SelectSewing> GetSelectedSewingLineFromEndline(List<string> listOrderID);
        IList<SelectSewing> GetSelectedSewingLine(string FactoryID);
        IList<SelectSewingTeam> GetSelectedSewingTeam();
        IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(List<string> listOrderID);
        IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(string finalInspectionID);
        List<string> GetArticleList(string finalInspectionID);
        IList<ArticleSize> GetArticleSizeList(string finalInspectionID);
        List<string> GetProductTypeList(string finalInspectionID);

        IList<SelectListItem> GetActionSelectListItem();
        IList<DatabaseObject.ProductionDB.System> GetSystem();
        int GetMeasurementAmount(string finalInspectionID);
        void UpdateOrderQtyShip(string finalInspectionID);
    }
}
