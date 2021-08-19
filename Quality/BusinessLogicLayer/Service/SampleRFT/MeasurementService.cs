using BusinessLogicLayer.Interface.SampleRFT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class MeasurementService : IMeasurementService
    {
        private IMeasurementProvider _IMeasurementProvider;
        public Measurement_ResultModel MeasurementGet(Measurement_Request measurement)
        {
            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_ResultModel measurement_Result = new Measurement_ResultModel() { Result = true };
            try
            {
                measurement_Result.TotalQty = _IMeasurementProvider.Get_Total_Measured_Qty();
                measurement_Result.MeasuredQty = _IMeasurementProvider.Get_Measured_Qty(measurement);

                DataTable dt = _IMeasurementProvider.Get_Measured_Detail(measurement);

                string jsonBody = JsonConvert.SerializeObject(dt);
                measurement_Result.JsonBody = jsonBody;
            }
            catch (Exception ex)
            {
                measurement_Result.Result = false;
                measurement_Result.ErrMsg = ex.ToString();
                throw ex;
            }

            return measurement_Result;
        }

        public Measurement_Request MeasurementGetPara(string OrderID)
        {
            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_Request measurement_Request = new Measurement_Request() { Result = true, ErrMsg = string.Empty};
            try
            {
                measurement_Request = _IMeasurementProvider.Get_OrdersPara(OrderID);
                measurement_Request.Articles = _IMeasurementProvider.GetAtricle(OrderID).Select(x => x.Article).ToList();
                measurement_Request.Article = _IMeasurementProvider.GetAtricle(OrderID).Select(x => x.Article).FirstOrDefault();
                measurement_Request.Result = true;
                measurement_Request.ErrMsg = string.Empty;
            }
            catch (Exception ex)
            {
                measurement_Request.ErrMsg = ex.ToString();
                measurement_Request.Result = false;
                throw ex;
            }

            return measurement_Request;
        }

        public void MeasurementToExcel(Measurement_Request measurement)
        {
            
            throw new NotImplementedException();
        }
    }
}
