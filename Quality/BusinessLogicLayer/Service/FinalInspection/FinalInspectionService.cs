using System;
using System.Collections.Generic;
using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
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

        public FinalInspection GetFinalInspection(string finalInspectionID)
        {
            FinalInspection result = new FinalInspection();

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
    }
}
