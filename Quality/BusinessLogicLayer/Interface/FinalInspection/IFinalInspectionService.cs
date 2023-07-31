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
        /// 部分功能Back/Next按鈕按下時，要存檔的東西(Remark之類的)
        /// </summary>
        /// <param name="finalInspection">如果是按Back或Next觸發finalInspection.InspectionStep要放入目標Page對應的Step讓後端更新
        /// 例如在General畫面[Back To Setting]則Setting；[Next]則Insp-General</param>
        /// <param name="userID"></param>
        /// <returns>BaseResult</returns>
        BaseResult UpdateStepInfo(FinalInspection finalInspection, string currentStep, string userID);

        string GetAQLPlanDesc(long ukey);

        object GetPivot88Json(string ID);
        object GetEndInlinePivot88Json(string ID, string inspectionType);

        List<SentPivot88Result> SentPivot88(PivotTransferRequest pivotTransferRequest);

        void ExecImp_EOLInlineInspectionReport();
        List<FinalInspection_Step> GetAllStep(string FinalInspectionID, string CustPONO);
        BaseResult CheckStep(string FinalInspectionID);
        BaseResult UpdateStepByAction(string FinalInspectionID, string UserID, FinalInspectionSStepAction action);

        List<FinalInspectionBasicGeneral> GetGeneralByBrand(string FinalInspectionID, string BrandID);
        List<FinalInspectionBasicCheckList> GetCheckListByBrand(string FinalInspectionID, string BrandID);

        List<FinalInspectionBasicGeneral> GetAllGeneral();
        List<FinalInspectionBasicCheckList> GetAllCheckList();
    }
}
