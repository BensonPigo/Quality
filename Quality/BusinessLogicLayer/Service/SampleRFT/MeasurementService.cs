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
using Sci;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Web.Mvc;

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

                // 取得圖片下拉選單
                List<SelectListItem> imageSourceList = _IMeasurementProvider.Get_ImageSource(measurement.OrderID).ToList();
                List<RFT_Inspection_Measurement_Image> imageList = _IMeasurementProvider.Get_ImageList(measurement.OrderID).ToList();

                DataTable dt = _IMeasurementProvider.Get_Measured_Detail(measurement);
                #region 處理OOT Qty

                List<string> columnListsp = new List<string>(); // 用來記錄幾筆有問題
                if (dt != null && dt.Rows.Count > 0)
                {
                    // 超出值變紅底
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
                            if (dr[item.ToString()] == DBNull.Value || string.IsNullOrEmpty(dr[item.ToString()].ToString()))
                            {
                                continue;
                            }

                            if (string.IsNullOrEmpty(dr["Tol(-)"].ToString()) || string.IsNullOrEmpty(dr["Tol(+)"].ToString()))
                            {
                                continue;
                            }

                            if (measurement.Unit.ToString().ToUpper() == "INCH")
                            {
                                string num;
                                if (dr[item.ToString()].ToString().Contains("-"))
                                {
                                    string d = dr[item.ToString()].ToString().Replace("-", string.Empty);
                                    num = _IMeasurementProvider.Get_CalculateSizeSpec(d, dr["Tol(-)"].ToString()).Rows[0]["value"].ToString();
                                }
                                else
                                {
                                    string d = dr[item.ToString()].ToString();
                                    num = _IMeasurementProvider.Get_CalculateSizeSpec(d, dr["Tol(+)"].ToString()).Rows[0]["value"].ToString();
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
                    Images_Source = (imageSourceList.Any() ? imageSourceList.ToList() : new List<SelectListItem>()),
                    Images = (imageList.Any() ? imageList.ToList() : new List<RFT_Inspection_Measurement_Image>()),
                    OOTQty = columnListsp.Count,                    
                    JsonBody = JsonConvert.SerializeObject(_IMeasurementProvider.Get_Measured_Detail(measurement_Request)),
                };
            }
            catch (Exception ex)
            {
                measurement_Result.Result = false;
                measurement_Result.ErrMsg = ex.Message.Replace("'", string.Empty);
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
                measurement_Request.ErrMsg = ex.Message.Replace("'", string.Empty);
                measurement_Request.Result = false;
            }

            return measurement_Request;
        }

        public Measurement_Request MeasurementToExcel(string OrderID, string FactoryID, bool test = false)
        {
            Measurement_Request result = new Measurement_Request()
            {
                Result = false,
                ErrMsg = "eff",
            };

            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_Request measurement_Request = MeasurementGetPara(OrderID, FactoryID);
            DataTable dt = _IMeasurementProvider.Get_Measured_Detail(measurement_Request);

            if (dt == null || dt.Rows.Count <= 0)
            {
                result.Result = false;
                result.ErrMsg = "Data not found!";
                return result;
            }

            string basefileName = "Measurement";
            string openfilepath;
            if (test)
            {
                openfilepath = "C:\\Willy_Repository\\Quality_KPI\\Quality\\Quality\\bin\\XLT\\Measurement.xlsx";
            }
            else
            {
                openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xlsx";
            }

            Microsoft.Office.Interop.Excel.Application objApp = MyUtility.Excel.ConnectExcel(openfilepath);

            objApp.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
            Microsoft.Office.Interop.Excel.Worksheet worksheet = objApp.ActiveWorkbook.Worksheets[1]; // 取得工作表

            int ColumnIndex = 1;

            // 新增header
            foreach (DataColumn dc in dt.Columns)
            {
                string column = dc.ColumnName;
                if (dc.ColumnName.IndexOf("_aa") > -1)
                {
                    column = dc.ColumnName.Replace(dc.ColumnName.Substring(dc.ColumnName.IndexOf("_aa"), dc.ColumnName.Length - dc.ColumnName.IndexOf("_aa")), "");
                }

                if (dc.ColumnName.IndexOf("diff") > -1)
                {
                    var index = dc.ColumnName.Replace("diff", "");
                    column = dc.ColumnName.Replace(index, "");
                }

                worksheet.Cells[1, ColumnIndex] = column;
                ColumnIndex++;
            }

            int ttlcolumn = dt.Columns.Count;
            int rowCnt = 2;

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 1; i <= ttlcolumn; i++)
                {
                    worksheet.Cells[rowCnt, i] = dr[i - 1];
                }

                rowCnt++;
            }

            worksheet.Columns.AutoFit();

            // Save Excel
            string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
            string filepath;
            if (test)
            {
                filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
            }
            else
            {
                filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);
            }

            Excel.Workbook workbook = objApp.ActiveWorkbook;
            workbook.SaveAs(filepath);
            workbook.Close();

            MyUtility.Excel.KillExcelProcess(objApp);            

            result.Result = true;
            result.FileName = fileName;

            return result;
        }

        public Measurement_Request DeleteMeasurementImage(long  ID)
        {
            _IMeasurementProvider = new MeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            Measurement_Request measurement_Request = new Measurement_Request() { Result = true, ErrMsg = string.Empty };
            try
            {
                var r = _IMeasurementProvider.DeleteMeasurementImgae(ID);

                measurement_Request.Result = true;
                measurement_Request.ErrMsg = string.Empty;
            }
            catch (Exception ex)
            {
                measurement_Request.ErrMsg = ex.Message.Replace("'", string.Empty);
                measurement_Request.Result = false;
            }

            return measurement_Request;
        }
    }
}
