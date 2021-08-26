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
        public ISystemProvider _SystemProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public Moisture GetMoistureForInspection(string finalInspectionID)
        {
            Moisture moisture = new Moisture();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                _SystemProvider = new SystemProvider(Common.ProductionDataAccessLayer);

                moisture.FinalInspectionID = finalInspectionID;
                moisture.FinalInspection_CTNMoisureStandard = _SystemProvider.Get()[0].FinalInspection_CTNMoisureStandard;
                moisture.ListArticle = _FinalInspFromPMSProvider.GetArticleList(finalInspectionID);
                moisture.ListCartonItem = _FinalInspectionProvider.GetMoistureListCartonItem(finalInspectionID).ToList();
                moisture.ListEndlineMoisture = _FinalInspectionProvider.GetEndlineMoisture().ToList();
                moisture.ActionSelectListItem = _FinalInspFromPMSProvider.GetActionSelectListItem().ToList();
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

                bool isMoistureExists = _FinalInspectionProvider.CheckMoistureExists(moistureResult.FinalInspectionID, moistureResult.Article, moistureResult.FinalInspection_OrderCartonUkey);

                if (!isMoistureExists)
                {
                    _FinalInspectionProvider.UpdateMoisture(moistureResult);
                }

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
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
