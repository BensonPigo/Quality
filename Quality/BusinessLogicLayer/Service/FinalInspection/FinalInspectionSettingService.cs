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
        public IQualityPass1Provider _QualityPass1 { get; set; }

        public Setting GetSettingForInspection(string finalInspectionID)
        {
            Setting result = new Setting();
            DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                FinalInspectionService f = new FinalInspectionService();

                finalInspection = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                if (!finalInspection)
                {
                    result.Result = false;
                    result.ErrorMessage = finalInspection.ErrorMessage;
                    return result;
                }

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);

                result.FinalInspectionID = finalInspectionID;
                result.InspectionStage = finalInspection.InspectionStage;
                result.AuditDate = finalInspection.AuditDate;
                result.SewingLineID = finalInspection.SewingLineID;
                result.Shift = finalInspection.Shift;
                result.Team = finalInspection.Team;
                result.InspectionTimes = finalInspection.InspectionTimes.ToString();

                result.AcceptableQualityLevelsUkey = finalInspection.AcceptableQualityLevelsUkey.ToString();
                result.AQLPlan = f.GetAQLPlanDesc(finalInspection.AcceptableQualityLevelsUkey);

                result.SampleSize = finalInspection.SampleSize;
                result.AcceptQty = finalInspection.AcceptQty;

                result.SelectedSewing = _FinalInspFromPMSProvider.GetSelectedSewingLine(finalInspection.FactoryID).ToList();
                result.SelectedSewingTeam = _FinalInspFromPMSProvider.GetSelectedSewingTeam().ToList();
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
                    selectedPOItem.Seq = selectedOrderShipSeq.Select(s => s.Seq).JoinToString(",");

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

        public Setting GetSettingForInspection(string CustPONO, List<string> listOrderID, string FactoryID, string UserID)
        {
            Setting result = new Setting();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _QualityPass1 = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);

                var tmp = _QualityPass1.Get(new Quality_Pass1() { ID = UserID });

                result.InspectionTimes = _FinalInspectionProvider.GetInspectionTimes(CustPONO);

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                result.SelectedSewing = _FinalInspFromPMSProvider.GetSelectedSewingLine(FactoryID).ToList();
                result.SelectedSewingTeam = _FinalInspFromPMSProvider.GetSelectedSewingTeam().ToList();
                result.SelectedPO = _FinalInspFromPMSProvider.GetSelectedPOForInspection(listOrderID).ToList();
                result.SelectOrderShipSeq = _FinalInspFromPMSProvider.GetSelectOrderShipSeqForSetting(listOrderID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(listOrderID).ToList();
                result.AcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();

                if (result.SelectedPO.Any(o=>o.BrandID.ToUpper() == "ADIDAS" || o.BrandID.ToUpper() == "REEBOK"))
                {
                    Quality_Pass1 Pass1 = tmp.Any() ? tmp.ToList().FirstOrDefault() : new Quality_Pass1();
                    if (string.IsNullOrEmpty(Pass1.Pivot88UserName) && UserID.ToUpper() != "SCIMIS")
                    {
                        result.Result = false;
                        result.ErrorMessage = $@"msg.WithError(""No Pivot88 account, please contact to local IT."")";
                        return result;
                    }
                }

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
            _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            finalInspectionID = string.Empty;
            try
            {
                // 每個OrderID至少要選擇一個 OrderShipSeq
                var listNotSelectShipSeq = setting.SelectOrderShipSeq.GroupBy(s => s.OrderID).Where(groupItem => !groupItem.Any(s => s.Selected));
                if (listNotSelectShipSeq.Any())
                {
                    result.Result = false;
                    result.ErrorMessage = "The following SP has not yet selected Shipmode Seq" + Environment.NewLine + listNotSelectShipSeq.Select(s => s.Key).JoinToString(",");
                    return result;
                }

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

                var selecedCarton = setting.SelectCarton != null ? setting.SelectCarton.Where(s => s.Selected) : new List<SelectCarton>();
                setting.SelectCarton = selecedCarton.Any() ? selecedCarton.ToList() : new List<SelectCarton>();

                var selectOrderShipSeq = setting.SelectOrderShipSeq.Where(s => s.Selected);
                setting.SelectOrderShipSeq = selectOrderShipSeq.Any() ? selectOrderShipSeq.ToList() : new List<SelectOrderShipSeq>();

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                string newfinalInspectionID = _FinalInspectionProvider.GetNewFinalInspectionID(factoryID);
                finalInspectionID = _FinalInspectionProvider.UpdateFinalInspection(setting, UserID, factoryID, MDivisionid, newfinalInspectionID);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"
msg.WithError(''{ex.Message}'');
";
            }
            return result;
        }
    }
}
