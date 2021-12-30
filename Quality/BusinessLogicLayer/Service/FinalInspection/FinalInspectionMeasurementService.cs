using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using BusinessLogicLayer.Interface;
using DatabaseObject;
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

        public Measurement GetMeasurementForInspection(string finalInspectionID, string userID)
        {
            Measurement measurement = new Measurement();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                

                measurement.FinalInspectionID = finalInspectionID;

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
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

                _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
                measurement.SizeUnit = _StyleProvider.GetSizeUnitByCustPONO(finalInspection.CustPONO);

                _MeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<DatabaseObject.ManufacturingExecutionDB.Measurement> baseMeasurementItems = _MeasurementProvider.GetMeasurementsByPOID(finalInspection.CustPONO, userID).ToList();
                measurement.ListMeasurementItem = baseMeasurementItems.Select( s =>
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

        public BaseResult UpdateMeasurement(Measurement measurement, string userID)
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

                _FinalInspectionProvider.UpdateMeasurement(measurement, userID);

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult Single_UpdateMeasurement(Measurement Head, MeasurementItem Body, string userID)
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

                _FinalInspectionProvider.UpdateMeasurement(Head, userID);

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
