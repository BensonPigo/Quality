using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionSettingService
    {
        /// <summary>
        /// 從InspectionQueryReport進入時呼叫
        /// </summary>
        /// <param name="finalInspectionID"></param>
        /// <returns>Setting</returns>
        Setting GetSettingForInspection(string finalInspectionID);

        /// <summary>
        /// 新建單子時由FindSP進入時呼叫
        /// </summary>
        /// <param name="POID">POID</param>
        /// <param name="listOrderID">listOrderID</param>
        /// <param name="FactoryID">FactoryID</param>
        /// <returns>Setting</returns>
        Setting GetSettingForInspection(string POID, List<string> listOrderID, string FactoryID, string UserID);

        /// <summary>
        /// UpdateFinalInspection
        /// </summary>
        /// <param name="setting">setting</param>
        /// <param name="UserID">UserID</param>
        /// <param name="factoryID">factoryID</param>
        /// <param name="MDivisionid">MDivisionid</param>
        /// <returns>回傳新增或異動的FinalInspectionID</returns>
        BaseResult UpdateFinalInspection(Setting setting, string UserID, string factoryID, string MDivisionid, out string finalInspectionID);
    }
}
