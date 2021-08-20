using System;
using System.Collections.Generic;
using System.Linq;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ToolKit;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionSettingService : IFinalInspectionSettingService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IOrdersProvider _OrdersProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public Setting GetSettingForInspection(string finalInspectionID)
        {
            Setting result = new Setting();
            DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);

                finalInspection = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                if (!finalInspection)
                {
                    result.Result = false;
                    result.ErrorMessage = finalInspection.ErrorMessage;
                    return result;
                }

                result.FinalInspectionID = finalInspectionID;
                result.InspectionStage = finalInspection.InspectionStage;
                result.AuditDate = finalInspection.AuditDate;
                result.SewingLineID = finalInspection.SewingLineID;
                result.InspectionTimes = finalInspection.InspectionTimes;
                result.AcceptableQualityLevelsUkey = finalInspection.AcceptableQualityLevelsUkey;
                result.SampleSize = finalInspection.SampleSize;
                result.AcceptQty = finalInspection.AcceptQty;

                result.SelectedPO = _FinalInspFromPMSProvider.GetSelectedPOForInspection(finalInspectionID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(finalInspectionID).ToList();

                foreach (SelectedPO selectedPOItem in result.SelectedPO)
                {
                    var selectedCartons = result.SelectCarton.Where(s => s.OrderID == selectedPOItem.OrderID);
                    if (!selectedCartons.Any())
                    {
                        continue;
                    }

                    selectedPOItem.Cartons = selectedCartons.Select(s => s.CTNNo).JoinToString(",");
                }

                result.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();
                result.ListSewingLine = _FinalInspFromPMSProvider.GetSewingLineForSetting(finalInspection.FactoryID).ToList();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public Setting GetSettingForInspection(string POID, List<string> listOrderID, string FactoryID)
        {
            Setting result = new Setting();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);

                result.InspectionTimes = _FinalInspectionProvider.GetInspectionTimes(POID);
                result.SelectedPO = _FinalInspFromPMSProvider.GetSelectedPOForInspection(listOrderID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(listOrderID).ToList();
                result.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();
                result.ListSewingLine = _FinalInspFromPMSProvider.GetSewingLineForSetting(FactoryID).ToList();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult UpdateFinalInspection(Setting setting, string UserID, string factoryID, string MDivisionid)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateFinalInspection(setting, UserID, factoryID, MDivisionid);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
    }
}
