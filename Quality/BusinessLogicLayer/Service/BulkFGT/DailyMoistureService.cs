using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using DocumentFormat.OpenXml.Spreadsheet;
using Library;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Interop.Excel;
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
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class DailyMoistureService
    {
        public DailyMoistureProvider _Provider;
        private MailToolsService _MailService;
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

                string Subject = $"Daily Moisture Test/{model.Main.OrderID}/" +
                    $"{model.Main.StyleID}/" +
                    $"Line {model.Main.Line}/" +
                    $"{model.Main.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                model.Main.MailSubject = Subject;
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


        public BaseResult ToReport2(string ReportNo, bool IsPDF, out string FinalFilenmae, string AssignedFineName = "")
        {

            BaseResult result = new BaseResult();
            _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;
            string tmpName = string.Empty;
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

                tmpName = $"Daily Moisture Test_{head.OrderID}_" +
                   $"{head.StyleID}_" +
                   $"{head.Line}_" +
                   $"{head.Result}_" +
                   $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";



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

                //string tmpName = $"Daily Bulk Moisture Test_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();
                char[] additionalChars = { '-', '+' }; // 您想要新增的字元
                char[] updatedInvalidChars = invalidChars.Concat(additionalChars).ToArray();

                foreach (char invalidChar in updatedInvalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }

                string excelFileName = $"{tmpName}.xlsx";
                string pdfFileName = $"{tmpName}.pdf";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workBook = excel.ActiveWorkbook;
                if (IsPDF)
                {
                    Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                    
                    workBook.ExportAsFixedFormat(targetType, pdfPath);
                    Marshal.ReleaseComObject(workBook);
                    FinalFilenmae = pdfFileName;
                }
                else
                {
                    workBook.SaveAs(excelPath);
                    //excel.ActiveWorkbook.SaveAs(excelPath);
                    FinalFilenmae = excelFileName;
                }

                workBook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workBook);
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
        public BaseResult ToReport(string ReportNo, bool IsPDF, out string FinalFilenmae, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;
            string tmpName = string.Empty;

            try
            {
                DailyMoisture_Result head = _Provider.GetMainData(new DailyMoisture_Request() { ReportNo = ReportNo });
                List<DailyMoisture_Detail_Result> body = _Provider.GetDetailData(ReportNo).ToList();

                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templatePath = Path.Combine(baseFilePath, "XLT", "DailyMoisture.xltx");
                string outputDirectory = Path.Combine(baseFilePath, "TMP");

                if (!File.Exists(templatePath))
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                tmpName = string.IsNullOrWhiteSpace(AssignedFineName) ?
                    $"Daily Moisture Test_{head.OrderID}_{head.StyleID}_{head.Line}_{head.Result}_{DateTime.Now:yyyyMMddHHmmss}" :
                    AssignedFineName;

                tmpName = Path.GetInvalidFileNameChars()
                    .Concat(new[] { '-', '+' })
                    .Aggregate(tmpName, (current, c) => current.Replace(c.ToString(), ""));

                string excelFileName = $"{tmpName}.xlsx";
                string excelPath = Path.Combine(outputDirectory, excelFileName);
                string pdfFileName = $"{tmpName}.pdf";
                string pdfPath = Path.Combine(outputDirectory, pdfFileName);

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填寫表頭資料
                    worksheet.Cell(1, 2).Value = head.ReportDate?.ToString("yyyy-MM-dd") ?? string.Empty;
                    worksheet.Cell(1, 4).Value = head.BrandID;
                    worksheet.Cell(1, 6).Value = head.OrderID;
                    worksheet.Cell(1, 8).Value = head.SeasonID;

                    worksheet.Cell(2, 2).Value = head.StyleID;
                    worksheet.Cell(2, 4).Value = head.Instrument;
                    worksheet.Cell(2, 6).Value = head.Fabrication;
                    worksheet.Cell(2, 8).Value = head.Standard * 0.01m;

                    worksheet.Cell(3, 2).Value = head.Line;

                    // 填寫表身資料
                    if (!body.Any())
                    {
                        worksheet.Row(6).Delete();
                    }
                    else
                    {
                        int startRow = 6;
                        for (int i = 0; i < body.Count - 1; i++)
                        {
                            worksheet.Row(startRow).InsertRowsBelow(1);
                        }

                        for (int i = 0; i < body.Count; i++)
                        { 
                            var row = startRow + i;
                            worksheet.Cell(row, 1).Value = body[i].Area;
                            worksheet.Cell(row, 2).Value = body[i].Fabric;
                            worksheet.Cell(row, 3).Value = body[i].Point1;
                            worksheet.Cell(row, 4).Value = body[i].Point2;
                            worksheet.Cell(row, 5).Value = body[i].Point3;
                            worksheet.Cell(row, 6).Value = body[i].Point4;
                            worksheet.Cell(row, 7).Value = body[i].Point5;
                            worksheet.Cell(row, 8).Value = body[i].Point6;
                            worksheet.Cell(row, 9).Value = body[i].Result;
                        }
                    }
                    workbook.SaveAs(excelPath);
                    if (IsPDF)
                    {
                        LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                        officeService.ConvertExcelToPdf(excelPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));

                        FinalFilenmae = pdfFileName;
                    }
                    else
                    {
                        FinalFilenmae = excelFileName;
                    }
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }


        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new DailyMoistureProvider(Common.ManufacturingExecutionDataAccessLayer);

            DailyMoisture_ViewModel model = this.GetDailyMoisture(new DailyMoisture_Request() { ReportNo = ReportNo });
            string name = $"Daily Moisture Test_{model.Main.OrderID}_" +
                    $"{model.Main.StyleID}_" +
                    $"Line {model.Main.Line}_" +
                    $"{model.Main.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            BaseResult report = this.ToReport(ReportNo, false,out string TempFileName, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Daily Moisture Test/{model.Main.OrderID}/" +
                    $"{model.Main.StyleID}/" +
                    $"Line {model.Main.Line}/" +
                    $"{model.Main.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",

                To = TO,
                CC = CC,
                Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                FileUploader = Files,
                IsShowAIComment = true,
                //AICommentType = "Accelerated Aging by Hydrolysis",
                StyleID = model.Main.StyleID,
                SeasonID = model.Main.SeasonID,
                BrandID = model.Main.BrandID,
            };

            if (!string.IsNullOrEmpty(Subject))
            {
                sendMail_Request.Subject = Subject;
            }

            _MailService = new MailToolsService();
            string comment = string.Empty;// _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = string.Empty;//_MailService.GetBuyReadyDate(sendMail_Request);

            sendMail_Request.Body = Body + Environment.NewLine + sendMail_Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
