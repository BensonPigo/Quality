using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Newtonsoft.Json;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ToolKit;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionMeasurementService : IFinalInspectionMeasurementService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IStyleProvider _StyleProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }
        public IMeasurementProvider _MeasurementProvider { get; set; }

        public ServiceMeasurement GetMeasurementForInspection(string finalInspectionID, string userID)
        {
            ServiceMeasurement measurement = new ServiceMeasurement();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                string OrderID = string.Empty;
                if (string.IsNullOrEmpty(finalInspection.CustPONO) || finalInspection.CustPONO　== "null")
                {
                    OrderID = _FinalInspectionProvider.Get_FinalInspectionID_Top1_OrderID(finalInspectionID);
                }

                measurement.FinalInspectionID = finalInspectionID;
                measurement.BrandID = finalInspection.BrandID;
                measurement.InspectionStage = finalInspection.InspectionStage;
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);
                measurement.ListArticle = _FinalInspFromPMSProvider.GetArticleList(finalInspectionID)
                            .Select(s => new SelectListItem() { 
                                Text = s,
                                Value = s,
                            }).ToList();
                measurement.ListSize = _FinalInspFromPMSProvider.GetArticleSizeList(finalInspectionID).ToList();
                measurement.ListProductType = _FinalInspFromPMSProvider.GetProductTypeList(finalInspectionID)
                            .Select(s => new SelectListItem()
                            {
                                Text = s,
                                Value = s,
                            }).ToList();

                measurement.MeasurementRemainingAmount = _FinalInspFromPMSProvider.GetMeasurementAmount(finalInspectionID);

                _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
                measurement.SizeUnit = _StyleProvider.GetSizeUnitByCustPONO(finalInspection.CustPONO, OrderID);

                _MeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<DatabaseObject.ManufacturingExecutionDB.Measurement> baseMeasurementItems = _MeasurementProvider.GetMeasurementsByPOID(finalInspection.CustPONO, OrderID, userID).ToList();
                
                measurement.ListMeasurementItem = baseMeasurementItems
                    .Where(o => measurement.ListSize.Any(x=>x.SizeCode == o.SizeCode))
                    .Select( s =>
                        new MeasurementItem() {
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
            }
            catch (Exception ex)
            {
                measurement.Result = false;
                measurement.ErrorMessage = ex.ToString();
            }

            return measurement;
        }
        public int GetMeasurementAmount(string finalInspectionID)
        {
            try
            {
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);
                int MeasurementRemainingAmount = _FinalInspFromPMSProvider.GetMeasurementAmount(finalInspectionID);
                return MeasurementRemainingAmount;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID)
        {
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<MeasurementViewItem> listMeasurementViewItem = _FinalInspectionProvider.GetMeasurementViewItem(finalInspectionID).ToList();

                foreach (MeasurementViewItem measurementViewItem in listMeasurementViewItem)
                {
                    DataTable dtMeasurementData = _FinalInspectionProvider.GetMeasurement(finalInspectionID, measurementViewItem.Article, measurementViewItem.Size, measurementViewItem.ProductType);
                    measurementViewItem.MeasurementDataByJson = JsonConvert.SerializeObject(dtMeasurementData);
                }

                return listMeasurementViewItem;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult InsertMeasurement(ServiceMeasurement measurement, string userID)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(measurement.FinalInspectionID);

                var needUpdMeasurement = measurement.ListMeasurementItem.Where(s => !string.IsNullOrEmpty(s.ResultSizeSpec));

                if (!needUpdMeasurement.Any() && 
                    (finalInspection.InspectionStage == "Final" || finalInspection.InspectionStage == "3rd Party")
                    )
                {
                    result.Result = false;
                    result.ErrorMessage = "Please input the measurement data if <Inspection Stage> is Final or 3rd Party";
                    return result;
                }

                measurement.ListMeasurementItem = needUpdMeasurement.ToList();

                _FinalInspectionProvider.InsertMeasurement(measurement, userID);

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult DeleteMeasurement(ServiceMeasurement measurement, DateTime AddDate)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.DeleteMeasurement(measurement, AddDate);

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
        public BaseResult UpdateMeasurement(ServiceMeasurement model, string userID)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateMeasurement(model, userID);

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
        public BaseResult Single_UpdateMeasurement(ServiceMeasurement Head, MeasurementItem Body, string userID)
        {
            BaseResult result = new BaseResult();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(Head.FinalInspectionID);


                if (finalInspection.InspectionStage == "Final" || finalInspection.InspectionStage == "3rd Party")                    
                {
                    result.Result = false;
                    result.ErrorMessage = "Please input the measurement data if <Inspection Stage> is Final or 3rd Party";
                    return result;
                }

                Head.ListMeasurementItem = new List<MeasurementItem>() { Body };

                _FinalInspectionProvider.InsertMeasurement(Head, userID);

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
