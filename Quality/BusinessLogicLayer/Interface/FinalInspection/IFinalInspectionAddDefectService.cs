using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionAddDefectService
    {
        /// <summary>
        /// 從InspectionQueryReport進入時呼叫
        /// </summary>
        /// <param name="finalInspectionID"></param>
        /// <returns>Setting</returns>
        AddDefect GetDefectForInspection(string finalInspectionID);

        BaseResult UpdateFinalInspectionDetail(AddDefect addDefect, string UserID);

        List<byte[]> GetDefectImage(long FinalInspection_DetailUkey);
    }
}
