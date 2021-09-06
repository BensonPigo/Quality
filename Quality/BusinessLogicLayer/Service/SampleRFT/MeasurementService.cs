using BusinessLogicLayer.Interface.SampleRFT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Newtonsoft.Json;

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

                measurement_Result.TotalQty = _IMeasurementProvider.Get_Total_Measured_Qty();
                measurement_Result.MeasuredQty = _IMeasurementProvider.Get_Measured_Qty(measurement);

                DataTable dt = _IMeasurementProvider.Get_Measured_Detail(measurement);
                #region 處理OOT Qty

                List<string> columnListsp = new List<string>(); // 用來記錄幾筆有問題
                if (dt != null && dt.Rows.Count > 0)
                {
                    bool bolCal;

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
                                if (dr[item.ToString()] == DBNull.Value)
                                {
                                    continue;
                                }

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
                                if (dr[item.ToString()] == DBNull.Value)
                                {
                                    continue;
                                }
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
                }
                #endregion

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
                    OOTQty = columnListsp.Count,
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

        public Measurement_Request MeasurementToExcel(string OrderID)
        {
            Measurement_Request result = new Measurement_Request() {
                Result = false,
                ErrMsg = "eff",
            };
            return result;
        }
    }
}
