using System;
using System.Collections.Generic;
using System.Linq;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel.EtoEFlowChart;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ToolKit;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionSettingService //: IFinalInspectionSettingService
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
                result.BrandID = finalInspection.BrandID;

                result.AcceptableQualityLevelsUkey = finalInspection.AcceptableQualityLevelsUkey.ToString();
                result.AQLPlan = f.GetAQLPlanDesc(finalInspection.AcceptableQualityLevelsUkey);

                result.SampleSize = finalInspection.SampleSize;
                result.AcceptQty = finalInspection.AcceptQty;

                result.SelectedSewing = _FinalInspFromPMSProvider.GetSelectedSewingLine(finalInspection.FactoryID).ToList();
                result.SelectedSewingTeam = _FinalInspFromPMSProvider.GetSelectedSewingTeam().ToList();

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);
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

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                var tmpAcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().Where(o => o.BrandID == string.Empty || o.BrandID == "AllBrand" || o.BrandID == finalInspection.BrandID);


                // 判斷該品牌有沒有特別設定，有的話就用特別設定；沒有的話用預設
                if (tmpAcceptableQualityLevels.Where(o=>o.BrandID != string.Empty && o.BrandID != "AllBrand").Any())
                {
                    // 有的話就用特別設定
                    result.AcceptableQualityLevels = tmpAcceptableQualityLevels.Where(o => o.BrandID == finalInspection.BrandID || o.BrandID == "AllBrand").OrderBy(o => o.AQLType).ToList();
                }
                else
                {
                    // 沒有的話用預設
                    result.AcceptableQualityLevels = tmpAcceptableQualityLevels.Where(o => o.BrandID == string.Empty || o.BrandID == "AllBrand").OrderBy(o=>o.AQLType).ToList();
                }


            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public Setting GetSettingForInspection(string CustPONO, List<string> listOrderID, string FactoryID, string UserID, string BrandID)
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
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SelectSewing> endlineSewing = _FinalInspFromPMSProvider.GetSelectedSewingLineFromEndline(listOrderID).ToList();

                foreach (var item in result.SelectedSewing)
                {
                    if (endlineSewing.Where(o=>o.SewingLine == item.SewingLine).Any())
                    {
                        item.Selected = true;
                    }
                }
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                result.SewingLineID = string.Join(",", result.SelectedSewing.Where(o => o.Selected).Select(o => o.SewingLine));

                result.SelectedSewingTeam = _FinalInspFromPMSProvider.GetSelectedSewingTeam().ToList();
                result.SelectedPO = _FinalInspFromPMSProvider.GetSelectedPOForInspection(listOrderID).ToList();
                result.SelectOrderShipSeq = _FinalInspFromPMSProvider.GetSelectOrderShipSeqForSetting(listOrderID).ToList();
                result.SelectCarton = _FinalInspFromPMSProvider.GetSelectedCartonForSetting(listOrderID).ToList();

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

                var tmpAcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().Where(o => o.BrandID == string.Empty || o.BrandID == "AllBrand" || o.BrandID == BrandID);


                // 判斷該品牌有沒有特別設定，有的話就用特別設定；沒有的話用預設
                if (tmpAcceptableQualityLevels.Where(o => o.BrandID != string.Empty && o.BrandID != "AllBrand").Any())
                {
                    // 有的話就用特別設定
                    result.AcceptableQualityLevels = tmpAcceptableQualityLevels.Where(o => o.BrandID == BrandID || o.BrandID == "AllBrand").OrderBy(o => o.AQLType).ToList();
                }
                else
                {
                    // 沒有的話用預設
                    result.AcceptableQualityLevels = tmpAcceptableQualityLevels.Where(o => o.BrandID == string.Empty || o.BrandID == "AllBrand").OrderBy(o => o.AQLType).ToList();
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

                string brandID = setting.BrandID;

                #region 檢查AQLPlan 重抓Sample Plan Qty
                if (setting.InspectionStage == "Final" ||
                    setting.InspectionStage == "3rd Party")
                {
                    int totalAvailableQty = setting.SelectedPO.Sum(s => s.AvailableQty);

                    var tmpAcceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().Where(o => o.BrandID == string.Empty || o.BrandID == "AllBrand" || o.BrandID == brandID);
                    setting.AcceptableQualityLevels = tmpAcceptableQualityLevels.ToList();

                    var AQLResult = setting.AcceptableQualityLevels.AsEnumerable();
                    int? maxStart = 0;

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
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1").Max(o => o.LotSize_Start);

                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            break;
                        case "1.0 Level II":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "2").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "2").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "2" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            break;
                        case "1.0 Level S-4":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "S-4").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "S-4").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "S-4" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            AQLResult.First().SampleSize = AQLResult.First().SampleSize == 0 ? totalAvailableQty : AQLResult.First().SampleSize;
                            break;
                        case "1.5 Level I":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)1.5 && o.InspectionLevels == "1").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)1.5 && o.InspectionLevels == "1").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)1.5 && o.InspectionLevels == "1" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            break;
                        case "2.5 Level I":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "1").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "1").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "1" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            break;
                        case "2.5 Level II":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "2").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "2").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)2.5 && o.InspectionLevels == "2" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
                            break;
                        case "4.0 Level I":
                            maxStart = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)4.0 && o.InspectionLevels == "1").Max(o => o.LotSize_Start);
                            if (totalAvailableQty > maxStart)
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)4.0 && o.InspectionLevels == "1").OrderByDescending(o => o.LotSize_Start);
                            }
                            else
                            {
                                AQLResult = setting.AcceptableQualityLevels.Where(o => o.AQLType == (decimal)4.0 && o.InspectionLevels == "1" && o.LotSize_Start <= totalAvailableQty && totalAvailableQty <= o.LotSize_End);
                            }
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
                else
                {
                    setting.AcceptableQualityLevelsUkey = null;
                }
                #endregion

                if (setting.AcceptQty > setting.SampleSize)
                {
                    result.Result = false;
                    result.ErrorMessage = $"Accepted Qty cannot be more than Sample Plan Qty ";
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
