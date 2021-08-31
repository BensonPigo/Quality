using System;
using System.Collections.Generic;
using System.Linq;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionService : IFinalInspectionService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IOrdersProvider _OrdersProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public string GetAQLPlanDesc(long ukey)
        {
            switch (ukey)
            {
                case -1:
                    return "100% Inspection";
                case 0:
                    return string.Empty;
                default:
                    _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ProductionDataAccessLayer);
                    List<AcceptableQualityLevels> acceptableQualityLevels = _FinalInspFromPMSProvider.GetAcceptableQualityLevelsForSetting().ToList();

                    if (!acceptableQualityLevels.Any(s => s.Ukey == ukey))
                    {
                        return string.Empty;
                    }

                    return acceptableQualityLevels.Where(s => s.Ukey == ukey)
                        .Select(s => $"{s.AQLType.ToString("0.0")} Level {s.InspectionLevels.Replace("1", "I").Replace("2", "II")}")
                        .First();
            }
        }

        public DatabaseObject.ManufacturingExecutionDB.FinalInspection GetFinalInspection(string finalInspectionID)
        {
            DatabaseObject.ManufacturingExecutionDB.FinalInspection result = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                result = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public IList<Orders> GetOrderForInspection(FinalInspection_Request request)
        {
            try
            {
                _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

                IList<Orders> result = _OrdersProvider.GetOrderForInspection(request);

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public IList<PoSelect_Result> GetOrderForInspection_ByModel(FinalInspection_Request request)
        {
            try
            {
                _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

                IList<PoSelect_Result> result = _OrdersProvider.GetOrderForInspection_ByModel(request);

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public BaseResult UpdateFinalInspectionByStep(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection, string currentStep, string userID)
        {
            BaseResult result = new BaseResult();
            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                _FinalInspectionProvider.UpdateFinalInspectionByStep(finalInspection, currentStep, userID);
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
