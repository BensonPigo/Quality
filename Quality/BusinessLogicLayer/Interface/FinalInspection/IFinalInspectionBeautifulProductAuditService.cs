using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionBeautifulProductAuditService
    {
        /// <summary>
        /// 從InspectionQueryReport進入時呼叫
        /// </summary>
        /// <param name="finalInspectionID"></param>
        /// <returns>Setting</returns>
        BeautifulProductAudit GetBeautifulProductAuditForInspection(string finalInspectionID);

        BaseResult UpdateBeautifulProductAudit(BeautifulProductAudit beautifulProductAudit, string UserID);

        List<byte[]> GetBACriteriaImage(long FinalInspection_DetailUkey);
    }
}
