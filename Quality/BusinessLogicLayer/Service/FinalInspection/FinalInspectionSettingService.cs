using System;
using System.Collections.Generic;
using System.Linq;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
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

                FinalInspectionService f = new FinalInspectionService();

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
                result.InspectionTimes = finalInspection.InspectionTimes.ToString();

                result.AcceptableQualityLevelsUkey = finalInspection.AcceptableQualityLevelsUkey.ToString();
                result.AQLPlan = f.GetAQLPlanDesc(finalInspection.AcceptableQualityLevelsUkey);

                result.SampleSize = finalInspection.SampleSize;
                result.AcceptQty = finalInspection.AcceptQty;

                result.SelectedPO = _FinalInspFromPMSProvider.GetSelectedPOForInspection(finalInspectionID).ToList();
                result.SelectOrderShipSeq = _FinalInspFromPMSProvider.GetSelectOrderShipSeqForSetting(finalInspectionID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(finalInspectionID).ToList();

                foreach (SelectedPO selectedPOItem in result.SelectedPO)
                {
                    var selectedOrderShipSeq = result.SelectOrderShipSeq.Where(s => s.Selected && s.OrderID == selectedPOItem.OrderID);
                    if (!selectedOrderShipSeq.Any())
                    {
                        continue;
                    }

                    selectedPOItem.Qty = selectedOrderShipSeq.Sum(s => s.Qty);
                    selectedPOItem.Article = selectedOrderShipSeq.Select(s => s.Article).JoinToString(",").Split(',').Distinct().JoinToString(",");

                    var selectedCartons = result.SelectCarton.Where(s => s.Selected && s.OrderID == selectedPOItem.OrderID);
                    if (!selectedCartons.Any())
                    {
                        continue;
                    }

                    selectedPOItem.Cartons = selectedCartons.Select(s => s.CTNNo).JoinToString(",");
                }

                result.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();
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
                result.SelectOrderShipSeq = _FinalInspFromPMSProvider.GetSelectOrderShipSeqForSetting(listOrderID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(listOrderID).ToList();
                result.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult UpdateFinalInspection(Setting setting, string UserID, string factoryID, string MDivisionid, out string finalInspectionID)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            finalInspectionID = string.Empty;
            try
            {

                #region 檢查AQLPlan 重抓Sample Plan Qty
                if (setting.InspectionStage == "Final" ||
                    setting.InspectionStage == "3rd Party")
                {
                    int totalAvailableQty = setting.SelectedPO.Sum(s => s.AvailableQty);
                    setting.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();
                    var AQLResult = setting.AcceptableQualityLevels.AsEnumerable();

                    switch (setting.AQLPlan)
                    {
                        case "":
                            AQLResult = new List<AcceptableQualityLevels>() {
                                new AcceptableQualityLevels(){
                                    AcceptedQty = setting.AcceptQty,
                                    SampleSize = totalAvailableQty,
                                    Ukey = 0
                                }
                            };
                            break;
                        case null:
                            AQLResult = new List<AcceptableQualityLevels>() {
                                new AcceptableQualityLevels(){
                                    AcceptedQty = setting.AcceptQty,
                                    SampleSize = totalAvailableQty,
                                    Ukey = 0
                                }
                            };
                            break;
                        case "1.0 Level I":
                            AQLResult = setting.AcceptableQualityLevels
                                .Where(s => s.AQLType == 1 &&
                                       s.InspectionLevels == "1" &&
                                       totalAvailableQty >= s.LotSize_Start && totalAvailableQty <= s.LotSize_End);
                            break;
                        case "1.0 Level II":
                            AQLResult = setting.AcceptableQualityLevels
                                .Where(s => s.AQLType == 1 &&
                                       s.InspectionLevels == "2" &&
                                       totalAvailableQty >= s.LotSize_Start && totalAvailableQty <= s.LotSize_End);
                            break;
                        case "1.5 Level I":
                            AQLResult = setting.AcceptableQualityLevels
                                .Where(s => s.AQLType == (decimal)1.5 &&
                                       s.InspectionLevels == "1" &&
                                       totalAvailableQty >= s.LotSize_Start && totalAvailableQty <= s.LotSize_End);
                            break;
                        case "2.5 Level I":
                            AQLResult = setting.AcceptableQualityLevels
                                .Where(s => s.AQLType == (decimal)2.5 &&
                                       s.InspectionLevels == "1" &&
                                       totalAvailableQty >= s.LotSize_Start && totalAvailableQty <= s.LotSize_End);
                            break;
                        case "100% Inspection":
                            AQLResult = new List<AcceptableQualityLevels>() {
                                new AcceptableQualityLevels(){
                                    AcceptedQty = setting.AcceptQty,
                                    SampleSize = totalAvailableQty,
                                    Ukey = -1
                                }
                            };
                            break;
                        default:
                            result.Result = false;
                            result.ErrorMessage = "When Inspection Stage is [Final] or [3rd Party], AQL Plan can not empty";
                            return result;
                    }

                    if (!AQLResult.Any())
                    {
                        result.Result = false;
                        result.ErrorMessage = $"No matching SampleSize setting found";
                        return result;
                    }

                    setting.SampleSize = AQLResult.First().SampleSize;
                    setting.AcceptQty = AQLResult.First().AcceptedQty;
                    setting.AcceptableQualityLevelsUkey = AQLResult.First().Ukey.ToString();
                }
                #endregion

                if (setting.AcceptQty > setting.SampleSize)
                {
                    result.Result = false;
                    result.ErrorMessage = $"[Accepted Qty] cannot be more than [Sample Plan Qty] ";
                    return result;
                }

                var selecedCarton = setting.SelectCarton.Where(s => s.Selected);
                setting.SelectCarton = selecedCarton.Any() ? selecedCarton.ToList() : new List<SelectCarton>();

                var selectOrderShipSeq = setting.SelectOrderShipSeq.Where(s => s.Selected);
                setting.SelectOrderShipSeq = selectOrderShipSeq.Any() ? selectOrderShipSeq.ToList() : new List<SelectOrderShipSeq>();

                finalInspectionID = _FinalInspectionProvider.UpdateFinalInspection(setting, UserID, factoryID, MDivisionid);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError('{ex.Message}');
";
            }
            return result;
        }
    }
}
