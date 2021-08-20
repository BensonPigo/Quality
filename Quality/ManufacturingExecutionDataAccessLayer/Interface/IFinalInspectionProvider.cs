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

        void UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid);
    }
}
