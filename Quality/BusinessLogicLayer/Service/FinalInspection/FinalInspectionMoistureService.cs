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
    public class FinalInspectionMoistureService : IFinalInspectionMoistureService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public Moisture GetMoistureForInspection(string finalInspectionID)
        {
            Moisture moisture = new Moisture();

            try
            {
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);

                moisture.FinalInspectionID = finalInspectionID;
                moisture.ListArticle = _FinalInspFromPMSProvider.GetArticleList(finalInspectionID);
                moisture.ActionSelectListItem = _FinalInspFromPMSProvider.GetActionSelectListItem().ToList();

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);


                List<FinalInspectionMoistureStandard> MoistureStandardSetting = _FinalInspectionProvider.GetMoistureStandardSetting(finalInspectionID);

                moisture.FinalInspection_CTNMoistureStandard = MoistureStandardSetting.Where(o => o.Category == "CTNMoisture").FirstOrDefault().Standard;
                moisture.FinalInspection_CTNMoistureStandardBM = MoistureStandardSetting.Where(o => o.Category == "CTNMoistureBM").FirstOrDefault().Standard;
                if (MoistureStandardSetting.Where(o => o.Category == "CTNMoistureLandtek").Any())
                {
                    moisture.FinalInspection_CTNMoistureStandardLandtek = MoistureStandardSetting.Where(o => o.Category == "CTNMoistureLandtek").FirstOrDefault().Standard;
                }
                else
                {
                    moisture.FinalInspection_CTNMoistureStandardLandtek = 0;
                }
                

                // 取得EndlineMoisture個別品牌設定，如果沒特別的就用預設值
                var ListEndlineMoisture = _FinalInspectionProvider.GetEndlineMoistureByBrand(finalInspectionID, string.Empty).ToList();
                var ListEndlineMoistureDefault = _FinalInspectionProvider.GetEndlineMoistureDefault().ToList();

                moisture.ListEndlineMoisture = ListEndlineMoisture.Any() ? ListEndlineMoisture : ListEndlineMoistureDefault;
                moisture.BrandID = moisture.ListEndlineMoisture.FirstOrDefault()?.BrandID;
                moisture.ListCartonItem = _FinalInspectionProvider.GetMoistureListCartonItem(finalInspectionID).ToList();
            }
            catch (Exception ex)
            {
                moisture.Result = false;
                moisture.ErrorMessage = ex.ToString();
            }

            return moisture;
        }

        public List<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID)
        {
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<ViewMoistureResult> listViewMoistureResult =
                    _FinalInspectionProvider.GetViewMoistureResult(finalInspectionID).ToList();

                return listViewMoistureResult;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult UpdateMoistureBySave(MoistureResult moistureResult)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                bool isMoistureExists = _FinalInspectionProvider.CheckMoistureExists(moistureResult.FinalInspectionID, moistureResult.Article, moistureResult.FinalInspection_OrderCartonUkey);

                if (isMoistureExists)
                {
                    result.Result = false;
                    result.ErrorMessage = "Moisture data already exists, please use View to delete first";
                    return result;
                }

                string newResult = string.Empty;

                result = this.CheckResult(moistureResult, out newResult);
                if (!result)
                {
                    return result;
                }

                moistureResult.Result = newResult;

                _FinalInspectionProvider.UpdateMoisture(moistureResult);

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult UpdateMoistureByNext(MoistureResult moistureResult)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = _FinalInspectionProvider.GetFinalInspection(moistureResult.FinalInspectionID);
                List<string> BrandIDs = _FinalInspectionProvider.Get_FinalInspectionID_BrandID(moistureResult.FinalInspectionID);

                // inline或3rd Party則可以跳過Moisture檢查 所以如果Instrument或Fabrication沒輸入可以直接跳過
                if ((finalInspection.InspectionStage == "Inline" || finalInspection.InspectionStage == "3rd Party") &&
                    (string.IsNullOrEmpty(moistureResult.Instrument) || string.IsNullOrEmpty(moistureResult.Fabrication))
                    )
                {
                    return result;
                }

                if (moistureResult.Instrument != null && moistureResult.Fabrication != null && (moistureResult.GarmentBottom == null || moistureResult.GarmentMiddle == null || moistureResult.GarmentTop == null))
                {
                    result.Result = false;
                    result.ErrorMessage = "Garment Moisture can not be empty.";
                    return result;
                }

                bool isMoistureExists = true; // _FinalInspectionProvider.CheckMoistureExists(moistureResult.FinalInspectionID, moistureResult.Article, moistureResult.FinalInspection_OrderCartonUkey);

                // ISP20230647 拔掉按下Next Back的存檔
                //if (!isMoistureExists &&
                //    !string.IsNullOrEmpty(moistureResult.Instrument) &&
                //    !string.IsNullOrEmpty(moistureResult.Fabrication))
                //{
                //    string newResult = string.Empty;

                //    result = this.CheckResult(moistureResult, out newResult);
                //    if (!result)
                //    {
                //        return result;
                //    }

                //    moistureResult.Result = newResult;

                //    _FinalInspectionProvider.UpdateMoisture(moistureResult);
                //}

                if ((finalInspection.InspectionStage == "Stagger" || finalInspection.InspectionStage == "Final") && BrandIDs.Where(o => o.ToUpper() == "ADIDAS").Any())
                {
                    isMoistureExists = _FinalInspectionProvider.CheckMoistureExists(moistureResult.FinalInspectionID, string.Empty, null);
                    if (!isMoistureExists)
                    {
                        result.Result = false;
                        result.ErrorMessage = "Please input the moisture fields if ＜Inspection Stage＞ is Stagger or Final.";
                        return result;
                    }
                }

                if (finalInspection.InspectionStage == "Final" && BrandIDs.Where(o => o.ToUpper() == "LLL").Any())
                {
                    isMoistureExists = _FinalInspectionProvider.CheckMoistureExists(moistureResult.FinalInspectionID, string.Empty, null);
                    if (!isMoistureExists)
                    {
                        result.Result = false;
                        result.ErrorMessage = "Must be inspected 3 garments..";
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

        private BaseResult CheckResult(MoistureResult moistureResult, out string result)
        {
            BaseResult baseResult = new BaseResult();
            result = string.Empty;

            if (string.IsNullOrEmpty(moistureResult.Instrument) ||
                string.IsNullOrEmpty(moistureResult.Fabrication))
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = "Please maintain [Detection Instrument] and [Fabrication] first";
                return baseResult;
            }

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<EndlineMoisture> listEndlineMoisture = _FinalInspectionProvider.GetEndlineMoistureByBrand(moistureResult.FinalInspectionID, string.Empty).ToList();
                List<EndlineMoisture> listDefaultEndlineMoisture = _FinalInspectionProvider.GetEndlineMoistureDefault().ToList();

                List<FinalInspectionMoistureStandard> MoistureStandardSetting = _FinalInspectionProvider.GetMoistureStandardSetting(moistureResult.FinalInspectionID);

                decimal CTNMoisureStandard = MoistureStandardSetting.Where(o => o.Category == "CTNMoisture").FirstOrDefault().Standard;

                // 取得EndlineMoisture個別品牌設定，如果沒特別的就用預設值
                decimal GMTMoisureStandard = listEndlineMoisture.Any() ? listEndlineMoisture
                                            .Where(s => s.Instrument == moistureResult.Instrument &&
                                                        s.Fabrication == moistureResult.Fabrication).First().Standard :
                                                        listDefaultEndlineMoisture
                                            .Where(s => s.Instrument == moistureResult.Instrument &&
                                                        s.Fabrication == moistureResult.Fabrication).First().Standard;

                if (moistureResult.FinalInspection_OrderCartonUkey != null)
                {
                    result = moistureResult.GarmentTop > GMTMoisureStandard ||
                        moistureResult.GarmentMiddle > GMTMoisureStandard ||
                        moistureResult.GarmentBottom > GMTMoisureStandard ||
                        moistureResult.CTNInside > CTNMoisureStandard ||
                        moistureResult.CTNOutside > CTNMoisureStandard ? "F" : "P";
                }
                else
                {
                    result = moistureResult.GarmentTop > GMTMoisureStandard ||
                            moistureResult.GarmentMiddle > GMTMoisureStandard ||
                            moistureResult.GarmentBottom > GMTMoisureStandard ? "F" : "P";
                }

                if (result == "F" && string.IsNullOrEmpty(moistureResult.Action))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "[Action] can not be empty";
                    return baseResult;
                }

            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult DeleteMoisture(long ukey)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.DeleteMoisture(ukey);

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
