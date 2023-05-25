using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class DailyMoistureService
    {
        public DailyMoistureProvider _Provider;
        private IOrdersProvider _OrdersProvider;

        public DailyMoisture_ViewModel GetDailyMoisture(DailyMoisture_Request Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Main = new DailyMoisture_Result(),
                Details = new List<DailyMoisture_Detail_Result>(),
                Request = Req,
            };

            try
            {
                _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SelectListItem> ReportNoList = _Provider.GetReportNo_Source(Req);

                if (!string.IsNullOrEmpty(Req.ReportNo) & ReportNoList.Any())
                {
                    model.Main = _Provider.GetMainData(Req);
                    if (model.Main.ReportNo == null)
                    {
                        model.Main = _Provider.GetMainData(new DailyMoisture_Request()
                        {
                            ReportNo = ReportNoList.FirstOrDefault().Value
                        });
                        //Req.ReportNo = model.Main.ReportNo;
                        //Req.BrandID = model.Main.BrandID;
                        //Req.SeasonID = model.Main.SeasonID;
                        //Req.StyleID = model.Main.StyleID;
                        //Req.OrderID = model.Main.OrderID;
                    }
                    model.Details = _Provider.GetDetailData(Req.ReportNo).ToList();
                }
                else if (string.IsNullOrEmpty(Req.ReportNo) && ReportNoList.Any())
                {
                    model.Main = _Provider.GetMainData(new DailyMoisture_Request()
                    {
                        ReportNo = ReportNoList.FirstOrDefault().Value,
                    });
                    model.Details = _Provider.GetDetailData(ReportNoList.FirstOrDefault().Value).ToList();
                }
                model.ReportNo_Source = ReportNoList;
                model.Request.ReportNo = model.Main.ReportNo;
                model.Request.BrandID = model.Main.BrandID;
                model.Request.SeasonID = model.Main.SeasonID;
                model.Request.StyleID = model.Main.StyleID;
                model.Request.OrderID = model.Main.OrderID;

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }
        public BaseResult Create(DailyMoisture_ViewModel model, string MDivision, string userid, out string NewReportNo)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new DailyMoistureProvider(_ISQLDataTransaction);
            NewReportNo = string.Empty;
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<DailyMoisture_Detail_Result>() : model.Details;

                foreach (var item in model.Details)
                {
                    item.Result = item.ReSetResult(model.Main.Standard);
                }

                ctn = _Provider.Insert_DailyMoisture(model.Main, MDivision, userid, out NewReportNo);

                if (ctn == 0)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrorMessage = "Create data fail.";
                    return result;
                }
                model.Main.ReportNo = NewReportNo;

                foreach (var detail in model.Details)
                {
                    detail.ReportNo = NewReportNo;
                    detail.EditName = model.Main.AddName;
                    ctn = _Provider.Insert_DailyMoisture_Detail(detail);
                    if (ctn == 0)
                    {
                        _ISQLDataTransaction.RollBack();
                        result.Result = false;
                        result.ErrorMessage = "Create detail Fail.";
                        return result;
                    }
                }

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult Update(DailyMoisture_ViewModel model, string userid)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new DailyMoistureProvider(_ISQLDataTransaction);
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<DailyMoisture_Detail_Result>() : model.Details;

                foreach (var item in model.Details)
                {
                    item.Result = item.ReSetResult(model.Main.Standard);
                }

                model.Main.EditName = userid;
                ctn = _Provider.Update_DailyMoisture(model);

                if (ctn == 0)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrorMessage = "update data fail.";
                    return result;
                }

                foreach (var detail in model.Details)
                {
                    detail.ReportNo = model.Main.ReportNo;
                    detail.EditName = userid;
                }
                _Provider.Update_DailyMoisture_Detail(model);

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult Delete(DailyMoisture_ViewModel model)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new DailyMoistureProvider(_ISQLDataTransaction);

            try
            {
                _Provider.Delete_DailyMoisture(model.Request.ReportNo);

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult EncodeAmend(DailyMoisture_ViewModel model)
        {
            BaseResult result = new BaseResult();

            try
            {
                _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);


                _Provider.Confirm_DailyMoisture(model);

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }
        public List<EndlineMoisture> GetEndlineMoisture()
        {
            _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<EndlineMoisture>  list = _Provider.GetEndlineMoisture().ToList();

            return list;
        }
        public List<SelectListItem> GetAction()
        {
            _Provider = new DailyMoistureProvider(Common.ProductionDataAccessLayer);
            List<SelectListItem> list = _Provider.GetAction().ToList();

            return list;
        }
        public List<Orders> GetOrders(Orders orders)
        {
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            try
            {
                orders.Category = "B";
                return _OrdersProvider.Get(orders).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }


        public BaseResult ToReport(string ReportNo, bool IsPDF, out string FinalFilenmae)
        {

            BaseResult result = new BaseResult();
            _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;

            try
            {
                DailyMoisture_Result head = _Provider.GetMainData(new DailyMoisture_Request() { ReportNo = ReportNo });
                List<DailyMoisture_Detail_Result> body = _Provider.GetDetailData(ReportNo).ToList();


                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string strXltName = baseFilePath + "\\XLT\\DailyMoisture.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }
                excel.Visible = false;
                excel.DisplayAlerts = false;

                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];


                // 表頭填入

                worksheet.Cells[1, 2] = head.ReportDate.HasValue ? head.ReportDate.Value.ToString("yyyy-MM-dd") : string.Empty;
                worksheet.Cells[1, 4] = head.BrandID;
                worksheet.Cells[1, 6] = head.OrderID;
                worksheet.Cells[1, 8] = head.SeasonID;

                worksheet.Cells[2, 2] = head.StyleID;
                worksheet.Cells[2, 4] = head.Instrument;
                worksheet.Cells[2, 6] = head.Fabrication;
                worksheet.Cells[2, 8] = (head.Standard * (decimal)0.01);

                worksheet.Cells[3, 2] = head.Line;
                // 表身筆數處理
                if (!body.Any())
                {
                    worksheet.get_Range("A6").EntireRow.Delete();
                }
                else
                {
                    int copyCount = body.Count - 1;

                    for (int i = 0; i < copyCount; i++)
                    {
                        Excel.Range paste1 = worksheet.get_Range($"A{i + 6}", Type.Missing);
                        Excel.Range copyRow = worksheet.get_Range("A6").EntireRow;
                        paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }


                // 表身填入
                int bodyStart = 6;
                foreach (var item in body)
                {
                    worksheet.Cells[bodyStart, 1] = item.Area;
                    worksheet.Cells[bodyStart, 2] = item.Fabric;
                    worksheet.Cells[bodyStart, 3] = item.Point1;
                    worksheet.Cells[bodyStart, 4] = item.Point2;
                    worksheet.Cells[bodyStart, 5] = item.Point3;
                    worksheet.Cells[bodyStart, 6] = item.Point4;
                    worksheet.Cells[bodyStart, 7] = item.Point5;
                    worksheet.Cells[bodyStart, 8] = item.Point6;
                    worksheet.Cells[bodyStart, 9] = item.Result;
                    bodyStart++;
                }

                string excelFileName = $"Daily Bulk Moisture Test_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string pdfFileName = $"Daily Bulk Moisture Test_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

                if (IsPDF)
                {
                    Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                    Excel.Workbook workBook = excel.ActiveWorkbook;
                    workBook.ExportAsFixedFormat(targetType, pdfPath);
                    Marshal.ReleaseComObject(workBook);
                    FinalFilenmae = pdfFileName;
                }
                else
                {
                    excel.ActiveWorkbook.SaveAs(excelPath);
                    FinalFilenmae = excelFileName;
                }

                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);

                result.Result = true;
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
