using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFinalInspFromPMSProvider
    {
        IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForSetting();
        IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID);
        IList<SelectedPO> GetSelectedPOForInspection(string finalInspectionID);
        IList<SelectCarton> GetSelectedCartonForSetting(List<string> listOrderID);
        IList<SelectCarton> GetSelectedCartonForSetting(string finalInspectionID);
    }
}
