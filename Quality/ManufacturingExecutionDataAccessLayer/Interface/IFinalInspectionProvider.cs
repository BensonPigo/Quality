using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IFinalInspectionProvider
    {
        FinalInspection GetFinalInspection(string FinalInspectionID);

        string GetInspectionTimes(string POID);

        string UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid);

        void UpdateFinalInspectionByStep(FinalInspection finalInspection, string currentStep, string userID);

        IList<byte[]> GetFinalInspectionDefectImage(long FinalInspection_DetailUkey);

        void UpdateFinalInspectionDetail(AddDefect addDefect, string UserID);
    }
}
