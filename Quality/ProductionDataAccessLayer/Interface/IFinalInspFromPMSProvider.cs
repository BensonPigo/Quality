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
        IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID);
        IList<SelectedPO> GetSelectedPOForInspection(string finalInspectionID);
        IList<SelectCarton> GetSelectedCartonForSetting(List<string> listOrderID);
        IList<SelectCarton> GetSelectedCartonForSetting(string finalInspectionID);
        IList<FinalInspectionDefectItem> GetFinalInspectionDefectItems(string finalInspectionID);

        IList<SelectSewing> GetSelectedSewingLine(string FactoryID);
        IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(List<string> listOrderID);
        IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(string finalInspectionID);
        List<string> GetArticleList(string finalInspectionID);
        IList<ArticleSize> GetArticleSizeList(string finalInspectionID);
        List<string> GetProductTypeList(string finalInspectionID);

        IList<SelectListItem> GetActionSelectListItem();
        IList<DatabaseObject.ProductionDB.System> GetSystem();
    }
}
