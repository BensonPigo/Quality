using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Web;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PullingTestService
    {
        private PullingTestProvider _PullingTestProvider;
        private bool IsTest = bool.Parse(ConfigurationManager.AppSettings["IsTest"]);
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public PullingTest_ViewModel GetReportNoList(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel result = Req;

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.ReportNo_Source = _PullingTestProvider.GetReportNoList(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }


        public PullingTest_ViewModel GetData(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.Detail = _PullingTestProvider.GetData(ReportNo);
                System.Data.DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);
                string Subject = $"Pulling Test/{dt.Rows[0]["POID"]}/" +
                        $"{dt.Rows[0]["StyleID"]}/" +
                        $"{dt.Rows[0]["Article"]}/" +
                        $"{dt.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                result.Detail.MailSubject = Subject;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_Result CheckSP(string POID)
        {
            PullingTest_Result result = new PullingTest_Result();
            PullingTest_Result result2 = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.CheckSP(POID);

                // 透過Brand得到PullForceUnit
                result2 = _PullingTestProvider.GetPullForceUnit(result.BrandID);

                result.PullForceUnit = result2.PullForceUnit;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }


        public PullingTest_Result GetStandard(string BrandID, string TestItem, string PullForceUnit, string StyleType)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.GetStandard(BrandID, TestItem, PullForceUnit, StyleType);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_Result GetPullUnit(string BrandID)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.GetPullForceUnit(BrandID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_ViewModel Insert(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Insert(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Update(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Update(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Delete(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Delete(ReportNo);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }


        public SendMail_Result SendMail(string ReportNo, string ToAddress, string CcAddress, string Subject, string Body, List<HttpPostedFileBase> Files)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                System.Data.DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);
                string name = $"Pulling Test_{dt.Rows[0]["POID"]}_" +
                        $"{dt.Rows[0]["StyleID"]}_" +
                        $"{dt.Rows[0]["Article"]}_" +
                        $"{dt.Rows[0]["Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                Report_Result baseResult = GetPDF(ReportNo, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;

                string unit = dt.Rows[0]["PullForceUnit"].ToString();
                if (!string.IsNullOrEmpty(unit))
                {
                    dt.Columns["PullForceUnit"].ColumnName = unit;
                    dt.Rows[0][unit] = dt.Rows[0]["PullForce"].ToString();
                    dt.Columns.Remove("PullForce");
                }

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CcAddress,
                    //Subject = "Pulling Test - Test Fail",
                    Subject = $"Pulling Test/{dt.Rows[0]["POID"]}/" +
                        $"{dt.Rows[0]["StyleID"]}/" +
                        $"{dt.Rows[0]["Article"]}/" +
                        $"{dt.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Pulling test for Snap/Button/Rivet",
                    StyleID = dt.Rows[0]["StyleID"].ToString(),
                    SeasonID = dt.Rows[0]["SeasonID"].ToString(),
                    BrandID = dt.Rows[0]["BrandID"].ToString(),
                };

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }


            return result;
        }

        public Report_Result GetPDF(string ReportNo, string AssignedFineName = "")
        {
            Report_Result result = new Report_Result();

            if (string.IsNullOrEmpty(ReportNo))
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            string baseFileName = "PullingTest";
            string tmpName = string.Empty;

            try
            {
                // 建立 TMP 目錄
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string xltDir = Path.Combine(baseFilePath, "XLT");
                string tmpDir = Path.Combine(baseFilePath, "TMP");

                if (!Directory.Exists(xltDir)) Directory.CreateDirectory(xltDir);
                if (!Directory.Exists(tmpDir)) Directory.CreateDirectory(tmpDir);

                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料
                PullingTest_Result model = _PullingTestProvider.GetData(ReportNo);
                var testCode = _QualityBrandTestCodeProvider.Get(model.BrandID, "Pulling test for Snap/Button/Rivet");
                DataTable ReportTechnician = _PullingTestProvider.GetReportTechnician(ReportNo);

                string templateFilePath = Path.Combine(xltDir, $"{baseFileName}.xltx");
                if (this.IsTest)
                {
                    templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLT", $"{baseFileName}.xltx");
                }

                tmpName = $"Pulling Test_{model.POID}_{model.StyleID}_{model.Article}_{model.Result}_{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[\/:*?""<>|]", ""); // 移除非法字元

                string excelFilePath = Path.Combine(tmpDir, $"{tmpName}.xlsx");
                string pdfFilePath = Path.Combine(tmpDir, $"{tmpName}.pdf");

                // 使用 ClosedXML 打開範本
                using (var workbook = new XLWorkbook(templateFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填寫報表資料
                    worksheet.Cell(2, 2).Value = model.ReportNo;
                    worksheet.Cell(2, 4).Value = DateTime.Now.ToString("yyyy/MM/dd");
                    worksheet.Cell(3, 2).Value = model.POID;
                    worksheet.Cell(3, 4).Value = model.TestDateText;
                    worksheet.Cell(4, 2).Value = model.SeasonID;
                    worksheet.Cell(4, 4).Value = model.StyleID;
                    worksheet.Cell(5, 2).Value = model.BrandID;
                    worksheet.Cell(5, 4).Value = model.Article;
                    worksheet.Cell(6, 2).Value = model.SizeCode;
                    worksheet.Cell(6, 4).Value = model.InspectorName;
                    worksheet.Cell(7, 2).Value = model.TestItem;
                    worksheet.Cell(7, 3).Value = model.PullForceUnit;
                    worksheet.Cell(7, 4).Value = model.PullForce;
                    worksheet.Cell(8, 2).Value = model.Time;
                    worksheet.Cell(8, 4).Value = model.Gender;
                    worksheet.Cell(9, 2).Value = model.FabricRefno;
                    worksheet.Cell(9, 4).Value = model.AccRefno;
                    worksheet.Cell(10, 2).Value = model.SnapOperator;
                    worksheet.Cell(10, 4).Value = model.Result;
                    worksheet.Cell(12, 1).Value = model.Remark;

                    // 簽名圖片
                    AddImageToWorksheet(worksheet, model.Signature, 23, 4, 100, 24);

                    // Technician 資訊
                    if (ReportTechnician.Rows.Count > 0)
                    {
                        string technicianName = ReportTechnician.Rows[0]["Technician"].ToString();
                        worksheet.Cell(22, 4).Value = technicianName;

                        if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                        {
                            byte[] technicianSignature = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"];
                            AddImageToWorksheet(worksheet, technicianSignature, 23, 4, 100, 24);
                        }
                    }

                    // TestBeforePicture 和 TestAfterPicture
                    AddImageToWorksheet(worksheet, model.TestBeforePicture, 16, 1, 200, 300);
                    AddImageToWorksheet(worksheet, model.TestAfterPicture, 16, 3, 200, 300);

                    // Excel 合併 + 塞資料
                    #region Title
                    string FactoryNameEN = _PullingTestProvider.GetFactoryNameEN(model.ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:D1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    mergedCell.Style.Font.Italic = false;
                    #endregion

                    // 儲存 Excel
                    workbook.SaveAs(excelFilePath);
                }

                // 轉 PDF
                LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                officeService.ConvertExcelToPdf(excelFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                result.TempFileName = $"{tmpName}.pdf";
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }

            return result;
        }

        // 新增圖片的共用方法
        private void AddImageToWorksheet(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
        {
            if (imageData != null)
            {
                using (var stream = new MemoryStream(imageData))
                {
                    worksheet.AddPicture(stream)
                             .MoveTo(worksheet.Cell(row, col), 5, 5) // 微調位置
                             .WithSize(width, height);
                }
            }
        }

    }
}
