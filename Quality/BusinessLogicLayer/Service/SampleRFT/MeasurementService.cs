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
using Newtonsoft.Json.Linq;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class MeasurementService : IMeasurementService
    {
        private IMeasurementProvider _IMeasurementProvider;
        public Measurement_ResultModel MeasurementGet(Measurement_Request measurement)
        {
            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_ResultModel measurement_Result = new Measurement_ResultModel();
            try
            {
                Measurement_Request measurement_Request = MeasurementGetPara(measurement.OrderID, measurement.Factory);
                measurement_Result = new Measurement_ResultModel() 
                { 
                    Result = true,
                    Factory = measurement_Request.Factory,
                    OrderID = measurement_Request.OrderID,
                    OrderTypeID = measurement_Request.OrderTypeID,
                    StyleID = measurement_Request.StyleID,
                    SeasonID = measurement_Request.SeasonID,
                    Unit = measurement_Request.Unit,
                    Article = measurement_Request.Article,
                    Articles = measurement_Request.Articles,
                    TotalQty = _IMeasurementProvider.Get_Total_Measured_Qty(),
                    MeasuredQty = _IMeasurementProvider.Get_Measured_Qty(measurement_Request),
                    JsonBody = JsonConvert.SerializeObject(_IMeasurementProvider.Get_Measured_Detail(measurement_Request)),
                };
            }
            catch (Exception ex)
            {
                measurement_Result.Result = false;
                measurement_Result.ErrMsg = ex.Message.ToString();
            }

            return measurement_Result;
        }

        public Measurement_Request MeasurementGetPara(string OrderID, string FactoryID)
        {
            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_Request measurement_Request = new Measurement_Request() { Result = true, ErrMsg = string.Empty};
            try
            {
                var query = _IMeasurementProvider.Get_OrdersPara(OrderID, FactoryID);
                if (!query.Any() || query.Count() == 0)
                {
                    throw new Exception("no data found");
                }

                measurement_Request = query.FirstOrDefault();
                measurement_Request.Articles = _IMeasurementProvider.GetAtricle(OrderID).Select(x => x.Article).ToList();
                measurement_Request.Article = measurement_Request.Articles.FirstOrDefault();
                measurement_Request.Result = true;
                measurement_Request.ErrMsg = string.Empty;
            }
            catch (Exception ex)
            {
                measurement_Request.ErrMsg = ex.Message.ToString();
                measurement_Request.Result = false;
            }

            return measurement_Request;
        }

        public void MeasurementToExcel(Measurement_Request measurement)
        {
            
            throw new NotImplementedException();
        }
    }
}
