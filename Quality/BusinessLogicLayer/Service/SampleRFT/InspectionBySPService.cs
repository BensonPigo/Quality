using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Newtonsoft.Json;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class InspectionBySPService
    {
        private InspectionBySPProvider _Provider;
        private IMailToProvider _IMailToProvider;

        public InspectionBySP_ViewModel Get_SearchResults(InspectionBySP_ViewModel Req)
        {
            InspectionBySP_ViewModel Result = new InspectionBySP_ViewModel()
            {
                OrderID = Req.OrderID,
                CustPONo = Req.CustPONo,
                StyleID = Req.StyleID,
                SeasonID = Req.SeasonID,
            };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                Result.DataList = _Provider.Get_SearchResults(Req).ToList();
                Result.Result = true;
            }
            catch (Exception ex)
            {
                Result.Result = false;
                Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return Result;
        }

        /// <summary>
        /// 取得產生新的Setting以供填寫
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public InspectionBySP_Setting GetNewSetting(string OrderID, string FactoryID)
        {
            InspectionBySP_Setting Result = new InspectionBySP_Setting();
            string InspectionTimes = string.Empty;
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                Orders orderInfo = _Provider.GetOrders(OrderID).FirstOrDefault();

                Result.OrderID = OrderID;
                Result.CustPONo = orderInfo.CustPONO;
                Result.StyleID = orderInfo.StyleID;
                Result.SeasonID = orderInfo.SeasonID;
                Result.BrandID = orderInfo.BrandID;
                Result.Article = orderInfo.Article;
                Result.FactoryID = orderInfo.FactoryID;
                Result.Model = orderInfo.Model;
                Result.SampleStage = orderInfo.OrderTypeID;
                Result.OrderQty = orderInfo.Qty.Value;
                Result.ComboType = orderInfo.ComboType;
                Result.TopSewingLineID = orderInfo.TopSewingLineID;
                Result.BottomSewingLineID = orderInfo.BottomSewingLineID;
                Result.InnerSewingLineID = orderInfo.InnerSewingLineID;
                Result.OuterSewingLineID = orderInfo.OuterSewingLineID;

                // 確認現在要進行第幾次檢驗
                List<SampleRFTInspection> list = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { OrderID = OrderID }).ToList();
                if (list.Count() == 0)
                {
                    // = 0 代表第一次驗
                    InspectionTimes = "1";
                }
                else
                {
                    // 其餘則是Fail過，因此直接+1
                    InspectionTimes = (list.Max(o => Convert.ToInt32(o.InspectionTimes)) + 1).ToString();
                }

                // InspectionTimes 不可編輯，因此直接帶出即可
                Result.InspectionTimes = InspectionTimes;
                Result.InspectionTimesList.Where(o => o.Value == InspectionTimes).FirstOrDefault().Selected = true;

                Result.SelectedSewing = _Provider.Get_SewingLineList(FactoryID, OrderID).ToList();
                Result.Select_QC_InCharge = _Provider.Get_QC_InChargeList().ToList();

                List<SelectListItem> SewingTeam = new List<SelectListItem>();
                foreach (var item in Result.SelectedSewing)
                {
                    SewingTeam.Add(new SelectListItem() { Text = item.SewingLine, Value = item.SewingLine, Selected = item.Selected });
                }
                Result.SewingLineList = SewingTeam;

                List<SelectListItem> QcList = new List<SelectListItem>();
                foreach (var item in Result.Select_QC_InCharge)
                {
                    QcList.Add(new SelectListItem() { Text = item.Pass1ID, Value = item.Pass1ID });
                }
                Result.QC_InChargeList = QcList;

                // BrandID =’ADIDAS’ or BrandID =’Reebok 才可以有AQL選項
                if (Result.BrandID.ToUpper() != "ADIDAS" && Result.BrandID.ToUpper() != "REEBOK")
                {
                    Result.InspectionStageList.RemoveAt(1);
                }

                Result.AcceptableQualityLevels = _Provider.GetAcceptableQualityLevelsForSetting(OrderID).ToList();
                Result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                Result.ExecuteResult = false;
                Result.ErrorMessage = ex.Message;
            }
            return Result;

        }

        /// <summary>
        /// 取得現有Setting紀錄
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="FactoryID"></param>
        /// <returns></returns>
        public InspectionBySP_Setting GetExistedSetting(InspectionBySP_ViewModel Req)
        {
            InspectionBySP_Setting Result = new InspectionBySP_Setting();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                SampleRFTInspection existedData = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = Req.ID }).FirstOrDefault();
                Orders orderInfo = _Provider.GetOrders(existedData.OrderID).FirstOrDefault();

                Result.OrderID = Req.OrderID;
                Result.CustPONo = orderInfo.CustPONO;
                Result.StyleID = orderInfo.StyleID;
                Result.SeasonID = orderInfo.SeasonID;
                Result.BrandID = orderInfo.BrandID;
                Result.Article = orderInfo.Article;
                Result.FactoryID = orderInfo.FactoryID;
                Result.Model = orderInfo.Model;
                Result.SampleStage = orderInfo.OrderTypeID;
                Result.OrderQty = orderInfo.Qty.Value;
                Result.ComboType = orderInfo.ComboType;
                Result.TopSewingLineID = orderInfo.TopSewingLineID;
                Result.BottomSewingLineID = orderInfo.BottomSewingLineID;
                Result.InnerSewingLineID = orderInfo.InnerSewingLineID;
                Result.OuterSewingLineID = orderInfo.OuterSewingLineID;
                Result.Dest = orderInfo.Dest;
                Result.VasShas = orderInfo.VasShas;

                // 取得現有Setting資料
                Result.ID = existedData.ID;
                Result.InspectionStage = existedData.InspectionStage;
                Result.InspectionTimes = existedData.InspectionTimes.ToString();
                Result.SewingLineID = existedData.SewingLineID;
                Result.SewingLine2ndID = existedData.SewingLine2ndID;
                Result.InspectionDate = existedData.InspectionDate;
                Result.QCInCharge = existedData.QCInCharge;
                Result.AQLPlan = existedData.AQLPlan;
                Result.SampleSize = existedData.SampleSize;
                Result.AcceptQty = existedData.AcceptQty;
                Result.AcceptableQualityLevelsUkey = existedData.AcceptableQualityLevelsUkey;

                #region 下拉選單資料來源
                Result.SelectedSewing = _Provider.Get_SewingLineList(Result.FactoryID, Result.OrderID).ToList();
                Result.Select_QC_InCharge = _Provider.Get_QC_InChargeList().ToList();

                List<SelectListItem> SewingTeam = new List<SelectListItem>();
                foreach (var item in Result.SelectedSewing)
                {
                    SewingTeam.Add(new SelectListItem() { Text = item.SewingLine, Value = item.SewingLine, Selected = item.Selected });
                }

                List<SelectListItem> QcList = new List<SelectListItem>();
                foreach (var item in Result.Select_QC_InCharge)
                {
                    QcList.Add(new SelectListItem() { Text = item.Pass1ID, Value = item.Pass1ID });
                }


                // BrandID =’ADIDAS’ or BrandID =’Reebok 才可以有AQL選項
                if (Result.BrandID.ToUpper() != "ADIDAS" && Result.BrandID.ToUpper() != "REEBOK")
                {
                    Result.InspectionStageList.RemoveAt(1);
                }

                Result.QC_InChargeList = QcList;
                Result.SewingLineList = SewingTeam;

                var AcceptableQualityLevelsList = _Provider.GetAcceptableQualityLevelsForSetting(existedData.OrderID);
                Result.AcceptableQualityLevels = AcceptableQualityLevelsList.ToList();
                Result.AQLPlan = AcceptableQualityLevelsList.Where(o => o.Ukey == Result.AcceptableQualityLevelsUkey).Select(o => $@"{o.AQLType} Level").FirstOrDefault();

                // InspectionTimes 不可編輯，因此直接帶出即可
                Result.InspectionTimesList.Where(o => o.Value == Result.InspectionTimes).FirstOrDefault().Selected = true;
                #endregion 
                Result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                Result.ExecuteResult = false;
                Result.ErrorMessage = ex.Message;
            }
            return Result;

        }

        public List<SampleRFTInspection> GetSampleRFTInspections(InspectionBySP_ViewModel Req)
        {
            List<SampleRFTInspection> list = new List<SampleRFTInspection>();
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                list = _Provider.Get_SampleRFTInspection(Req).ToList();

            }
            catch (Exception ex)
            {

                throw ex;
            }


            return list;
        }

        /// <summary>
        /// 確認該SP#的檢驗狀況
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public string CheckInspection(InspectionBySP_ViewModel Req)
        {
            string InspectionProgress = string.Empty;
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SampleRFTInspection> list = _Provider.Get_SampleRFTInspection(Req).ToList();

                if (list.Count() == 0)
                {
                    // 沒有檢驗過
                    InspectionProgress = "NonInspect";
                }
                else if (list.Where(o => o.Result.ToUpper() == "PASS").Any())
                {
                    // 已經有檢驗Pass紀錄
                    InspectionProgress = "Pass";
                }
                else if (list.Where(o => o.Result.ToUpper() == "FAIL").Any() && !list.Where(o => string.IsNullOrEmpty(o.Result)).Any())
                {
                    // 所有紀錄都已經完成檢驗，且只有Fail
                    InspectionProgress = "Failure";
                }
                else if (list.Where(o => string.IsNullOrEmpty(o.Result)).Any())
                {
                    // 沒有Pass，且還有檢驗未完成
                    InspectionProgress = "InProcess";
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }


            return InspectionProgress;
        }

        public string CheckOrderStyleUnit(InspectionBySP_ViewModel Req)
        {
            string OrderStyleUnit = string.Empty;
            try
            {
                _Provider = new InspectionBySPProvider(Common.ProductionDataAccessLayer);
                OrderStyleUnit = _Provider.GetOrderStyleUnit(Req.OrderID);

            }
            catch (Exception ex)
            {

                throw ex;
            }


            return OrderStyleUnit;
        }
        public BaseResult UpdateSampleRFTInspectionByStep(SampleRFTInspection SampleRFTInspection, string currentStep, string userID)
        {
            BaseResult result = new BaseResult();
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                _Provider.UpdateSampleRFTInspectionnByStep(SampleRFTInspection, currentStep, userID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        /// <summary>
        /// 儲存Setting的異動，並傳出新增的ID
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public InspectionBySP_Setting SettingProcess(InspectionBySP_Setting Req)
        {
            InspectionBySP_Setting result = new InspectionBySP_Setting();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                SampleRFTInspection data = _Provider.SettingSampleRFTInspection(Req).FirstOrDefault();
                result.ID = data.ID;


                result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                result.ExecuteResult = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }

        public InspectionBySP_Measurement GetMeasurement(long ID, string UserID)
        {
            InspectionBySP_Measurement measurement = new InspectionBySP_Measurement();
            try
            {

                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                SampleRFTInspection data = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();

                measurement.OrderID = data.OrderID;
                measurement.FactoryID = data.FactoryID;
                measurement.StyleUkey = data.StyleUkey;
                measurement.SewingLineID = data.SewingLineID;
                measurement.ListArticle = _Provider.GetArticleList(data.OrderID)
                            .Select(s => new SelectListItem()
                            {
                                Text = s,
                                Value = s,
                            }).ToList();
                measurement.ListSize = _Provider.GetArticleSizeList(data.OrderID).ToList();
                measurement.ListProductType = _Provider.GetProductTypeList(data.OrderID)
                            .Select(s => new SelectListItem()
                            {
                                Text = s,
                                Value = s,
                            }).ToList();

                measurement.SizeUnit = _Provider.GetSizeUnitByCustPONO(data.OrderID);
                List<DatabaseObject.ManufacturingExecutionDB.Measurement> baseMeasurementItems = _Provider.GetMeasurementsByPOID(data.OrderID, UserID).ToList();
                measurement.ListMeasurementItem = baseMeasurementItems.Select(s =>
                       new MeasurementItem()
                       {
                           Description = s.Description,
                           SizeSpec = s.SizeSpec,
                           Tol1 = s.Tol1,
                           Tol2 = s.Tol2,
                           Size = s.SizeCode,
                           Code = s.Code,
                           CanEdit = !s.IsPatternMeas,
                           MeasurementUkey = s.Ukey
                       }
                    ).ToList();

                measurement.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                measurement.ExecuteResult = false;
                measurement.ErrorMessage = ex.ToString();
            }

            return measurement;

        }

        public List<MeasurementViewItem> GetMeasurementViewItem(string OrderID)
        {
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<MeasurementViewItem> listMeasurementViewItem = _Provider.GetMeasurementViewItem(OrderID).ToList();
                //SampleRFTInspection sampleRFTInspection = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = inputID }).FirstOrDefault();

                foreach (MeasurementViewItem measurementViewItem in listMeasurementViewItem)
                {
                    DataTable dtMeasurementData = _Provider.GetMeasurement(OrderID, measurementViewItem.Article, measurementViewItem.Size, measurementViewItem.ProductType);
                    measurementViewItem.MeasurementDataByJson = JsonConvert.SerializeObject(dtMeasurementData);
                }

                return listMeasurementViewItem;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Measurement_ResultModel GetMeasurementImageList(string OrderID)
        {
            Measurement_ResultModel measurement_Result = new Measurement_ResultModel();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                // 取得圖片下拉選單
                List<SelectListItem> imageSourceList = _Provider.Get_MeasurementImageSource(OrderID).ToList();
                List<RFT_Inspection_Measurement_Image> imageList = _Provider.Get_MeasurementImageList(OrderID).ToList();
                measurement_Result = new Measurement_ResultModel()
                {
                    Result = true,
                    Images_Source = (imageSourceList.Any() ? imageSourceList.ToList() : new List<SelectListItem>()),
                    Images = (imageList.Any() ? imageList.ToList() : new List<RFT_Inspection_Measurement_Image>()),
                };
            }
            catch (Exception ex)
            {
                measurement_Result.Result = false;
                measurement_Result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            return measurement_Result;
        }

        public InspectionBySP_Measurement InsertMeasurement(InspectionBySP_Measurement measurement)
        {
            InspectionBySP_Measurement result = new InspectionBySP_Measurement();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);


                //var needUpdMeasurement = measurement.ListRFTMeasurementItem.Where(s => !string.IsNullOrEmpty(s.ResultSizeSpec));
                List<RFT_Inspection_Measurement> rft = new List<RFT_Inspection_Measurement>();

                if (measurement.ListMeasurementItem != null)
                {
                    foreach (var item in measurement.ListMeasurementItem)
                    {
                        if (string.IsNullOrEmpty(item.ResultSizeSpec))
                        {
                            continue;
                        }
                        RFT_Inspection_Measurement t = new RFT_Inspection_Measurement()
                        {
                            MeasurementUkey = item.MeasurementUkey,
                            Code = item.Code,
                            SizeCode = measurement.SelectedSize,
                            ResultSizeSpec = item.ResultSizeSpec,
                            Article = measurement.SelectedArticle,
                            Location = measurement.SelectedProductType,
                        };
                        rft.Add(t);
                    }
                }              
                measurement.ListRFTMeasurementItem = rft;

                _Provider.InsertRFT_Inspection_Measurement(measurement);
                result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                result.ExecuteResult = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public InspectionBySP_AddDefect GetAddDefectBody(long ID)
        {
            InspectionBySP_AddDefect result = new InspectionBySP_AddDefect() { ID = ID };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                SampleRFTInspection existedData = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();

                if (existedData.InspectionStage == "100%")
                {
                    result.MaxRejectQty = existedData.OrderQty;
                }
                else if (existedData.InspectionStage == "AQL")
                {
                    result.MaxRejectQty = existedData.SampleSize.Value;
                }

                var summary = _Provider.GetDefectDefaultBody(ID);
                result.RejectQty = existedData.RejectQty;
                result.ListDefectItem = summary.ToList();
                result.Areas = _Provider.GetArea().ToList();

                result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                result.ExecuteResult = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"")";
            }

            return result;
        }

        public SampleRFTInspection_Summary GetDefectImageList(long ID, long SampleRFTInspectionDetailUKey)
        {
            SampleRFTInspection_Summary model = new SampleRFTInspection_Summary();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                // 取得圖片下拉選單
                List<SelectListItem> imageSourceList = _Provider.Get_DefectImageSource(ID, SampleRFTInspectionDetailUKey).ToList();
                List<DefectImage> imageList = _Provider.Get_DefectImageList(ID, SampleRFTInspectionDetailUKey).ToList();
                model = new SampleRFTInspection_Summary()
                {
                    Result = true,
                    Images_Source = (imageSourceList.Any() ? imageSourceList.ToList() : new List<SelectListItem>()),
                    Images = (imageList.Any() ? imageList.ToList() : new List<DefectImage>()),
                };
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            return model;
        }

        /// <summary>
        /// 包含三個步驟，依序為：1.刪除圖片  2.異動Detail(同時更新受影響的PMSFile圖片Detail UKey欄位)  3.新增圖片
        /// </summary>
        /// <param name="addDefct"></param>
        /// <param name="DeleteImg"></param>
        public InspectionBySP_AddDefect AddDefectProcess(InspectionBySP_AddDefect addDefct, List<DefectImage> DeleteImg)
        {
            InspectionBySP_AddDefect model = new InspectionBySP_AddDefect();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                _Provider = new InspectionBySPProvider(_ISQLDataTransaction);

                List<SampleRFTInspection_Detail> detail = new List<SampleRFTInspection_Detail>();
                foreach (SampleRFTInspection_Summary item in addDefct.ListDefectItem)
                {
                    long SampleRFTInspectionUkey = item.ID;
                    long DetailUkey = item.UKey;
                    string DefectCode = item.DefectCode;
                    string GarmentDefectTypeID = item.GarmentDefectTypeID;
                    string GarmentDefectCodeID = item.GarmentDefectCodeID;
                    string Responsibility = item.Responsibility;
                    List<string> AreaCodes = item.AreaCodes != null ? item.AreaCodes.Split(',').ToList() : new List<string>();

                    if (AreaCodes.Count != 0)
                    {
                        int Qty = 1;
                        foreach (var area in AreaCodes)
                        {
                            SampleRFTInspection_Detail d = new SampleRFTInspection_Detail()
                            {
                                SampleRFTInspectionUkey = SampleRFTInspectionUkey,
                                //UKey = DetailUkey,
                                DefectCode = DefectCode,
                                GarmentDefectTypeID = GarmentDefectTypeID,
                                GarmentDefectCodeID = GarmentDefectCodeID,
                                Responsibility = Responsibility,
                                Qty = Qty,
                                AreaCode = area,
                            };
                            detail.Add(d);
                        }
                    }
                    else
                    {
                        int Qty = 0;
                        SampleRFTInspection_Detail d = new SampleRFTInspection_Detail()
                        {
                            SampleRFTInspectionUkey = SampleRFTInspectionUkey,
                            //UKey = DetailUkey,
                            DefectCode = DefectCode,
                            GarmentDefectTypeID = GarmentDefectTypeID,
                            GarmentDefectCodeID = GarmentDefectCodeID,
                            Responsibility = Responsibility,
                            Qty = Qty,
                            AreaCode = string.Empty,
                        };
                        detail.Add(d);
                    }
                }

                // 傳入Detail UKey明確的圖片，直接刪除
                _Provider.AddDefect_Delete_Image(DeleteImg);

                // 開始異動表身，有些Detail可能會刪掉，但圖片所屬是看GarmentDefectCodeID，因此圖片不能跟著刪，要把Detail UKey改成其他相同GarmentDefectCodeID的Detail UKey
                if (detail.Count > 0)
                {
                    _Provider.AddDefect_Detail_Process(detail);
                }

                // 新增圖片，由於前面更新過了Detail UKey，因此直接抓GarmentDefectCodeID對應的其中一個Detail Ukey，寫入PMSFile即可
                foreach (var defectItem in addDefct.ListDefectItem)
                {
                    _Provider.AddDefect_Add_Image(defectItem);
                }

                // 更新表頭
                _Provider.AddDefect_Update_Head(new SampleRFTInspection()
                {
                    ID = addDefct.ID,
                    RejectQty = addDefct.RejectQty
                });

                _ISQLDataTransaction.Commit();
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return model;
        }

        public InspectionBySP_BA GetBA_Body(long ID)
        {
            InspectionBySP_BA result = new InspectionBySP_BA() { ID = ID };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                SampleRFTInspection existedData = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();

                if (existedData.InspectionStage == "100%")
                {
                    result.MaxBAQty = existedData.OrderQty;
                }
                else if (existedData.InspectionStage == "AQL")
                {
                    result.MaxBAQty = existedData.SampleSize.Value;
                }

                var body = _Provider.GetBeautifulProductAudit(ID);
                result.BAQty = existedData.BAQty;

                result.ListBACriteria = body.ToList();

                result.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                result.ExecuteResult = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"")";
            }

            return result;
        }
        public BACriteriaItem GetBAImageList(long ID, long SampleRFTInspectionDetailUKey)
        {
            BACriteriaItem model = new BACriteriaItem();

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                // 取得圖片下拉選單
                List<SelectListItem> imageSourceList = _Provider.Get_BAImageSource(ID, SampleRFTInspectionDetailUKey).ToList();
                List<BAImage> imageList = _Provider.Get_BAImageList(ID, SampleRFTInspectionDetailUKey).ToList();
                model = new BACriteriaItem()
                {
                    Result = true,
                    Images_Source = (imageSourceList.Any() ? imageSourceList.ToList() : new List<SelectListItem>()),
                    Images = (imageList.Any() ? imageList.ToList() : new List<BAImage>()),
                };
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            return model;
        }
        public InspectionBySP_BA BAProcess(InspectionBySP_BA ba, List<BAImage> DeleteImg)
        {
            InspectionBySP_BA model = new InspectionBySP_BA();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                _Provider = new InspectionBySPProvider(_ISQLDataTransaction);

                List<BACriteriaItem> detail = new List<BACriteriaItem>();

                // 傳入Detail UKey明確的圖片，直接刪除
                _Provider.BA_Delete_Image(DeleteImg);

                // 開始異動表身，有些Detail可能會刪掉，但圖片所屬是看GarmentDefectCodeID，因此圖片不能跟著刪，要把Detail UKey改成其他相同GarmentDefectCodeID的Detail UKey
                _Provider.BA_Detail_Process(ba.ListBACriteria);

                // 新增圖片，由於前面更新過了Detail UKey，因此直接抓GarmentDefectCodeID對應的其中一個Detail Ukey，寫入PMSFile即可
                foreach (var defectItem in ba.ListBACriteria)
                {
                    _Provider.BA_Add_Image(defectItem);
                }

                // 更新表頭
                _Provider.BA_Update_Head(new SampleRFTInspection()
                {
                    ID = ba.ID,
                    BAQty = ba.BAQty
                });

                _ISQLDataTransaction.Commit();
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return model;
        }

        public InspectionBySP_DummyFit GetDummyFit(long ID)
        {
            InspectionBySP_DummyFit model = new InspectionBySP_DummyFit()
            {
                DetailList=new List<DummyFitImage>()
            };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                SampleRFTInspection existedData = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();
                model.OrderID = existedData.OrderID;
                model.StyleID = existedData.StyleID;

                model.ArticleList = _Provider.GetArticleList(existedData.OrderID)
                                .Select(s => new SelectListItem()
                                {
                                    Text = s,
                                    Value = s,
                                }).ToList(); ;

                model.ArticleSizeList = _Provider.GetArticleSizeList(existedData.OrderID).ToList();
                List<string> listSize = model.ArticleSizeList.Select(O => O.SizeCode).Distinct().ToList();
                model.SizeList = listSize.Select(s => new SelectListItem()
                {
                    Text = s,
                    Value = s,
                }).ToList(); ;


                model.DetailList = _Provider.GetDummyFitImages(existedData.OrderID).ToList();
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            return model;
        }

        public InspectionBySP_DummyFit DummyFitProcess(InspectionBySP_DummyFit req)
        {
            InspectionBySP_DummyFit model = new InspectionBySP_DummyFit();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                _Provider = new InspectionBySPProvider(_ISQLDataTransaction);

                List<BACriteriaItem> detail = new List<BACriteriaItem>();

                _Provider.DummyFitUpdate(req);

                _ISQLDataTransaction.Commit();
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return model;
        }

        public InspectionBySP_Others Get_Others(long ID)
        {
            InspectionBySP_Others model = new InspectionBySP_Others()
            {
                DataList = new List<CFTComments_Result>(),
                SamePOIDList = new List<SelectListItem>() { new SelectListItem() }
            };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                SampleRFTInspection existedData = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();
                model.OrderID = existedData.OrderID;
                model.StyleID = existedData.StyleID;
                model.MReMail = existedData.MReMail;
                model.CCMail = existedData.MDivisionID.ToUpper() == "PM1" || existedData.MDivisionID.ToUpper() == "PM2" || existedData.MDivisionID.ToUpper() == "PM9" ? "arvin.goles@sportscity.com.ph" : string.Empty;

                List<string> samePOIDList = _Provider.GetSamePOIDList(ID).ToList();
                var RFTInspection_Result = _Provider.Get_RFTInspection_Result(ID).ToList();
                var res = _Provider.Get_CFT_OrderComments(model.OrderID).ToList();

                //foreach (var item in res)
                //{
                //    item.Comnments = item.Comnments == null ? string.Empty : item.Comnments.Replace("*", Environment.NewLine);
                //}
                foreach (var item in samePOIDList)
                {

                    model.SamePOIDList.Add(new SelectListItem() { Text = item, Value = item });
                }

                model.InspectorResult = RFTInspection_Result.FirstOrDefault().InspectorResult;
                model.DataList = res;
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public InspectionBySP_Others Get_CFT_OrderComments(string OrderID)
        {
            InspectionBySP_Others model = new InspectionBySP_Others()
            {
                DataList = new List<CFTComments_Result>(),
                SamePOIDList = new List<SelectListItem>() { new SelectListItem() }
            };

            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                var res = _Provider.Get_CFT_OrderComments(OrderID).ToList();

                foreach (var item in res)
                {
                    item.Comnments = item.Comnments == null ? string.Empty : item.Comnments.Replace("*", "<br>");
                }

                model.DataList = res;
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public InspectionBySP_Others OthersProcess(InspectionBySP_Others Req)
        {
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            InspectionBySP_Others model = new InspectionBySP_Others();
            try
            {
                _Provider = new InspectionBySPProvider(_ISQLDataTransaction);

                SampleRFTInspection existseddata = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = Req.ID }).FirstOrDefault();

                Req.OrderID = existseddata.OrderID;
                _Provider.OthersUpdate(Req);
                if (Req.Action != "Back")
                {
                    _Provider.Others_Update_Head(Req.ID);
                }

                _ISQLDataTransaction.Commit();
                model.ExecuteResult = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.ExecuteResult = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return model;

        }

        public QueryInspectionBySP_ViewModel GetQuery(QueryInspectionBySP_ViewModel Req)
        {
            QueryInspectionBySP_ViewModel model = new QueryInspectionBySP_ViewModel() { DataList = new List<QueryInspectionBySP>() };
            try
            {
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);

                var details = _Provider.Get_QueryResults(Req);

                model.DataList = details.Any() ? details.ToList() : new List<QueryInspectionBySP>();
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return model;

        }

        public QueryReport GetQueryDetail(long inputID, string UserID)
        {
            _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
            QueryReport model = new QueryReport();

            try
            {
                SampleRFTInspection sampleRFTInspection = _Provider.Get_SampleRFTInspection(new InspectionBySP_ViewModel() { ID = inputID }).FirstOrDefault();

                sampleRFTInspection.CCMail = sampleRFTInspection.MDivisionID.ToUpper() == "PM1" || sampleRFTInspection.MDivisionID.ToUpper() == "PM2" || sampleRFTInspection.MDivisionID.ToUpper() == "PM9" ? "arvin.goles@sportscity.com.ph" : string.Empty;

                InspectionBySP_Setting setting = GetExistedSetting(new InspectionBySP_ViewModel()
                {
                    ID = sampleRFTInspection.ID,
                    OrderID = sampleRFTInspection.OrderID
                });

                _Provider = new InspectionBySPProvider(Common.ProductionDataAccessLayer);
                string orderStyleUnit = _Provider.GetOrderStyleUnit(sampleRFTInspection.OrderID);
                _Provider = new InspectionBySPProvider(Common.ManufacturingExecutionDataAccessLayer);
                setting.OrderStyleUnit = orderStyleUnit;

                InspectionBySP_Measurement Measurement = GetMeasurement(inputID, UserID);
                InspectionBySP_AddDefect AddDefect = GetAddDefectBody(inputID);
                InspectionBySP_BA BA = GetBA_Body(inputID);
                InspectionBySP_DummyFit DummyFit = GetDummyFit(inputID);
                InspectionBySP_Others Others = Get_Others(inputID);

                model.sampleRFTInspection = sampleRFTInspection;
                model.Setting = setting;

                model.ListMeasurementViewItem = _Provider.GetMeasurementViewItem(sampleRFTInspection.OrderID).ToList();
                model.Measurement = Measurement;

                model.GoOnInspectURL = this.GetCurrentAction(sampleRFTInspection.InspectionStep);

                foreach (MeasurementViewItem measurementViewItem in model.ListMeasurementViewItem)
                {
                    DataTable dtMeasurementData = _Provider.GetMeasurement(sampleRFTInspection.OrderID, measurementViewItem.Article, measurementViewItem.Size, measurementViewItem.ProductType);
                    measurementViewItem.MeasurementDataByJson = JsonConvert.SerializeObject(dtMeasurementData);
                }

                model.AddDefect = AddDefect;
                model.BA = BA;
                model.DummyFit = DummyFit;
                model.Others = Others;


                model.ExcuteResult = true;
            }
            catch (Exception ex)
            {
                model.ExcuteResult = false;
                model.ErrorMessage =  $@"msg.WithError(""{ex.Message}"")";
            }

            return model;
        }

        private string GetCurrentAction(string InspectionStep)
        {
            string ActionName = string.Empty;

            switch (InspectionStep)
            {
                case "Insp-Setting":
                    ActionName = "Setting";
                    break;
                case "Insp-CheckList":
                    ActionName = "CheckList";
                    break;
                case "Insp-AddDefect":
                    ActionName = "AddDefect";
                    break;
                case "Insp-BA":
                    ActionName = "BeautifulProductAudit";
                    break;
                case "Insp-Measurement":
                    ActionName = "Measurement";
                    break;
                case "Insp-DummyFit":
                    ActionName = "DummyFitting";
                    break;
                case "Insp-Others":
                    ActionName = "Others";
                    break;
                default:
                    break;
            }

            return ActionName;
        }

    }
}
