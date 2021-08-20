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
                #region 處理OOT Qty

                if (dt != null && dt.Rows.Count > 0)
                {
                    bool bolCal;
                    List<string> columnListsp = new List<string>(); // 用來記錄幾筆有問題

                    List<string> diffArry = new List<string>();
                    foreach (var itemCol in dt.Columns)
                    {
                        if (itemCol.ToString().Contains("diff"))
                        {
                            diffArry.Add(itemCol.ToString());
                        }
                    }

                    foreach (DataRow dr in dt.Rows)
                    {
                        foreach (var item in diffArry)
                        {
                            if (measurement.Unit.ToString().ToUpper() == "INCH")
                            {
                                string num;
                                if (dr[item.ToString()].ToString().Contains("-"))
                                {
                                    string d = dr[item.ToString()].ToString().Replace("-", string.Empty);
                                    num = _IMeasurementProvider.Get_CalculateSizeSpec(d, dr["Tol(-)"].ToString()).Rows[0]["Vaule"].ToString();
                                }
                                else
                                {
                                    string d = dr[item.ToString()].ToString();
                                    num = _IMeasurementProvider.Get_CalculateSizeSpec(d, dr["Tol(+)"].ToString()).Rows[0]["Vaule"].ToString();
                                }

                                bolCal = num.Contains("-") && !string.IsNullOrEmpty(dr[item.ToString()].ToString());
                            }
                            else
                            {
                                double d = Convert.ToDouble(dr[item.ToString()]);
                                double num;
                                if (d < 0)
                                {
                                    num = Math.Abs(d) - Convert.ToDouble(dr["Tol(-)"]);
                                }
                                else
                                {
                                    num = d - Convert.ToDouble(dr["Tol(+)"]);
                                }

                                bolCal = num > 0;
                            }

                            if (bolCal)
                            {
                                if (!columnListsp.Contains(item.ToString()))
                                {
                                    columnListsp.Add(item.ToString());
                                }
                            }
                        }
                    }

                    measurement_Result.OOTQty = columnListsp.Count;
                }
                #endregion

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
