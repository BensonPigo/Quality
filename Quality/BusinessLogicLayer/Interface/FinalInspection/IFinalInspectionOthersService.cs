using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionOthersService
    {
        Others GetOthersForInspection(string finalInspectionID);

        BaseResult UpdateOthersBack(Others others, string UserID);

        BaseResult UpdateOthersSubmit(Others others, string UserID);

        List<OtherImage> GetOthersImage(string finalInspectionID);

        BaseResult UpdateOthersSubmitPMS(Others others);
    }
}
