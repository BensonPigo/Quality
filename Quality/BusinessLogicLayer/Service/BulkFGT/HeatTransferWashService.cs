using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class HeatTransferWashService
    {
        public HeatTransferWashProvider _Provider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        public BaseResult Create(HeatTransferWash_ViewModel model, string MDivision, string userid, out string NewReportNo)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);
            NewReportNo = string.Empty;
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<HeatTransferWash_Detail_Result>() : model.Details;

                if (model.Details.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                {
                    model.Main.Result = "Fail";
                }
                else
                {
                    model.Main.Result = "Pass";
                }

                ctn = _Provider.Insert_HeatTransferWash(model.Main, MDivision, userid, out NewReportNo);

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
                    ctn = _Provider.Insert_HeatTransferWash_Detail(detail);
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

        public BaseResult Update(HeatTransferWash_ViewModel model, string userid)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<HeatTransferWash_Detail_Result>() : model.Details;

                if (model.Details.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                {
                    model.Main.Result = "Fail";
                }
                else
                {
                    model.Main.Result = "Pass";
                }

                model.Main.EditName = userid;
                ctn = _Provider.Update_HeatTransferWash(model);

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
                _Provider.Update_HeatTransferWash_Detail(model);

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

        public BaseResult Delete(HeatTransferWash_ViewModel model)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);

            try
            {
                _Provider.Delete_HeatTransferWash(model.Request.ReportNo);

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

        public BaseResult EncodeAmend(HeatTransferWash_ViewModel model)
        {
            BaseResult result = new BaseResult();

            try
            {
                _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
                _Provider.Confirm_HeatTransferWash(model);

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }
        public HeatTransferWash_ViewModel GetHeatTransferWash(HeatTransferWash_Request Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                Request = Req,
            };

            try
            {
                _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SelectListItem> ReportNoList = _Provider.GetReportNo_Source(Req);

                if (!string.IsNullOrEmpty(Req.ReportNo))
                {
                    model.Main = _Provider.GetMainData(Req);
                    model.Details = _Provider.GetDetailData(Req.ReportNo).ToList();
                }
                else if (string.IsNullOrEmpty(Req.ReportNo) && ReportNoList.Any())
                {
                    model.Main = _Provider.GetMainData(new HeatTransferWash_Request()
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
                model.Request.Article = model.Main.Article;

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public BaseResult ToReport(string ReportNo ,bool IsPDF,  out string FinalFilenmae)
        {

            BaseResult result = new BaseResult();
            _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;

            try
            {
                HeatTransferWash_Result head = _Provider.GetMainData(new HeatTransferWash_Request() { ReportNo = ReportNo });
                List<HeatTransferWash_Detail_Result> body =_Provider.GetDetailData(ReportNo).ToList();


                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string strXltName = baseFilePath + "\\XLT\\HeatTransferWash.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }
                excel.Visible = false;
                excel.DisplayAlerts = false;

                Excel.Worksheet worksheet= excel.ActiveWorkbook.Worksheets[1];


                // 表頭填入

                worksheet.Cells[2, 4] = head.ReportDate.HasValue ? head.ReportDate.Value.ToString("yyyy-MM-dd") : string.Empty;
                worksheet.Cells[3, 2] = head.OrderID;
                worksheet.Cells[3, 4] = head.IsTeamwear ? "V" : string.Empty;
                worksheet.Cells[4, 2] = head.SeasonID;
                worksheet.Cells[4, 4] = head.StyleID;
                worksheet.Cells[5, 2] = head.Article;
                worksheet.Cells[5, 4] = head.Line;
                worksheet.Cells[6, 2] = head.BrandID;
                worksheet.Cells[6, 4] = head.Machine;
                worksheet.Cells[9, 3] = head.Temperature;
                worksheet.Cells[10, 3] = head.Time;
                worksheet.Cells[11, 3] = head.Pressure;
                worksheet.Cells[12, 3] = head.PeelOff;
                worksheet.Cells[14, 1] = head.Cycles;
                worksheet.Cells[14, 3] = head.TemperatureUnit;

                if (head.Result == "Pass")
                {
                    worksheet.Cells[17, 1] = "V";
                }
                else
                {
                    worksheet.Cells[17, 3] = "V";
                }
                worksheet.Cells[19, 1] = head.Remark;

                string imgPath_BeforePicture = string.Empty;
                string imgPath_AfterPicture = string.Empty;
                if (head.TestBeforePicture != null )
                {
                    byte[] beforePic = head.TestBeforePicture;
                    imgPath_BeforePicture = ToolKit.PublicClass.AddImageSignWord(beforePic, head.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    Excel.Range cell = worksheet.Cells[22, 1];
                    worksheet.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 2, cell.Top + 2, 220, 130);
                }
                if (head.TestBeforePicture != null)
                {
                    byte[] beforePic = head.TestAfterPicture;
                    imgPath_AfterPicture = ToolKit.PublicClass.AddImageSignWord(beforePic, head.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    Excel.Range cell = worksheet.Cells[22, 3];
                    worksheet.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 2, cell.Top + 2, 220, 130);
                }


                // 表身筆數處理
                if (!body.Any())
                {
                    worksheet.get_Range("A7").EntireRow.Delete();
                }
                else
                {
                    int copyCount = body.Count - 1;

                    for (int i = 0; i < copyCount; i++)
                    {
                        Excel.Range paste1 = worksheet.get_Range($"A{i + 7}", Type.Missing);
                        Excel.Range copyRow = worksheet.get_Range("A7").EntireRow;
                        paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }


                // 表身填入
                int bodyStart = 7;
                foreach (var item in body)
                {
                    worksheet.Cells[bodyStart, 2] = item.FabricRefNo;
                    worksheet.Cells[bodyStart, 4] = item.HTRefNo;
                    bodyStart++;
                }

                string excelFileName = $"Daily HT Wash Test_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string pdfFileName = $"Daily HT Wash Test_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";

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

        public List<Order_Qty> GetDistinctArticle(Order_Qty order_Qty)
        {
            _OrderQtyProvider = new OrderQtyProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _OrderQtyProvider.GetDistinctArticle(order_Qty).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
