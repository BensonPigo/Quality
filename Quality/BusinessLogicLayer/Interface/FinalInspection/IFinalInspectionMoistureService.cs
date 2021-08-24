using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionMoistureService
    {
        Moisture GetMoistureForInspection(string finalInspectionID);

        BaseResult UpdateMoisture(MoistureResult moistureResult);

        BaseResult DeleteMoisture(long ukey);

        List<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID);
    }
}
