using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ResultModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionService
    {
        IList<Orders> GetOrderForInspection(FinalInspection_Request request);

        IList<PoSelect_Result> GetOrderForInspection_ByModel(FinalInspection_Request request);

        FinalInspection GetFinalInspection(string FinalInspectionID);

        /// <summary>
        /// UpdateFinalInspectionByStep
        /// </summary>
        /// <param name="finalInspection">如果是按Back或Next觸發finalInspection.InspectionStep要放入目標Page對應的Step讓後端更新
        /// 例如在General畫面[Back To Setting]則Setting；[Next]則Insp-General</param>
        /// <param name="userID"></param>
        /// <returns>BaseResult</returns>
        BaseResult UpdateFinalInspectionByStep(FinalInspection finalInspection, string currentStep, string userID);

        string GetAQLPlanDesc(long ukey);

        object GetPivot88Json(string ID);

        List<SentPivot88Result> SentPivot88(PivotTransferRequest pivotTransferRequest);
    }
}
