using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
//using Excel = Microsoft.Office.Interop.Excel;


namespace BusinessLogicLayer.Service.BulkFGT
{
    public class WaterFastnessService
    {
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];
        public WaterFastnessProvider _WaterFastnessProvider;
        public IScaleProvider _ScaleProvider;
        public IOrdersProvider _OrdersProvider;
        public IStyleProvider _StyleProvider;
        private MailToolsService _MailService;
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public BaseResult AmendWaterFastnessDetail(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                WaterFastness_Detail_Result WaterFastness_Detail_Result = _WaterFastnessProvider.GetWaterFastness_Detail(poID, TestNo);

                if (WaterFastness_Detail_Result.Main.Status != "Confirmed")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {WaterFastness_Detail_Result.Main.Status}, can not Amend";
                    return baseResult;
                }

                _WaterFastnessProvider.AmendWaterFastness(poID, TestNo);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult EncodeWaterFastnessDetail(string poID, string TestNo, out string waterFastnessResult)
        {
            BaseResult baseResult = new BaseResult();
            _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
            waterFastnessResult = string.Empty;
            try
            {
                WaterFastness_Detail_Result waterFastness_Detail_Result = _WaterFastnessProvider.GetWaterFastness_Detail(poID, TestNo);

                if (waterFastness_Detail_Result.Main.Status != "New")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {waterFastness_Detail_Result.Main.Status}, can not Encode";
                    return baseResult;
                }

                if (waterFastness_Detail_Result.Details.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Data is empty please fill-in data first.";
                    return baseResult;
                }

                if (waterFastness_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.WaterFastnessGroup)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Group cannot be empty.";
                    return baseResult;
                }

                if (waterFastness_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.SEQ)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Seq cannot be empty.";
                    return baseResult;
                }

                if (waterFastness_Detail_Result.Details.Any(s =>
                    string.IsNullOrEmpty(s.ChangeScale) ||
                    string.IsNullOrEmpty(s.AcetateScale) ||
                    string.IsNullOrEmpty(s.CottonScale) ||
                    string.IsNullOrEmpty(s.NylonScale) ||
                    string.IsNullOrEmpty(s.PolyesterScale) ||
                    string.IsNullOrEmpty(s.AcrylicScale) ||
                    string.IsNullOrEmpty(s.WoolScale) ||
                    string.IsNullOrEmpty(s.Result)
                ))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Color Change Scale, Color Staining Scale and Result cannot be empty.";
                    return baseResult;
                }

                string result = waterFastness_Detail_Result.Details.Any(s => s.Result == "Fail") ? "Fail" : "Pass";
                waterFastnessResult = result;
                _WaterFastnessProvider.EncodeWaterFastness(poID, TestNo, result);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
                return baseResult;
            }
        }

        public WaterFastness_Detail_Result GetWaterFastness_Detail_Result(string poID, string TestNo, string BrandID)
        {
            try
            {
                WaterFastness_Detail_Result waterFastness_Detail_Result = new WaterFastness_Detail_Result();
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                waterFastness_Detail_Result = _WaterFastnessProvider.GetWaterFastness_Detail(poID, TestNo, BrandID);

                waterFastness_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();

                if (!string.IsNullOrEmpty(TestNo))
                {
                    DataTable dtResult = _WaterFastnessProvider.GetFailMailContentData(poID, TestNo);
                    string Subject = $"Water Fastness Test/{poID}/" +
                            $"{dtResult.Rows[0]["Style"]}/" +
                            $"{dtResult.Rows[0]["Article"]}/" +
                            $"{dtResult.Rows[0]["Result"]}/" +
                            $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    waterFastness_Detail_Result.Main.MailSubject = Subject;
                }

                return waterFastness_Detail_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public WaterFastness_Result GetWaterFastness_Result(string POID)
        {
            WaterFastness_Result result = new WaterFastness_Result();
            try
            {
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                result = _WaterFastnessProvider.GetWaterFastness_Main(POID);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public BaseResult SaveWaterFastnessDetail(WaterFastness_Detail_Result waterFastness_Detail_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                if (string.IsNullOrEmpty(waterFastness_Detail_Result.Main.Article))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Article cannot be empty.";
                    return baseResult;
                }

                if (string.IsNullOrEmpty(waterFastness_Detail_Result.Main.Inspector))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Inspector cannot be empty.";
                    return baseResult;
                }

                if (waterFastness_Detail_Result.Main.InspDate == null)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Test Date cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = waterFastness_Detail_Result
                   .Details.GroupBy(s => new
                   {
                       s.WaterFastnessGroup,
                       s.SEQ,
                   })
                   .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[OvenGroup]{s.Key.WaterFastnessGroup}, [SEQ]{s.Key.SEQ}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }

                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);

                //再檢查一次Result
                foreach (WaterFastness_Detail_Detail waterFastness_Detail_Detail in waterFastness_Detail_Result.Details)
                {
                    if (waterFastness_Detail_Detail.ResultChange == null)
                    {
                        waterFastness_Detail_Detail.ResultChange = string.Empty;
                    }

                    if (waterFastness_Detail_Detail.ResultAcetate == null)
                    {
                        waterFastness_Detail_Detail.ResultAcetate = string.Empty;
                    }
                    if (waterFastness_Detail_Detail.ResultCotton == null)
                    {
                        waterFastness_Detail_Detail.ResultCotton = string.Empty;
                    }

                    if (waterFastness_Detail_Detail.ResultNylon == null)
                    {
                        waterFastness_Detail_Detail.ResultNylon = string.Empty;
                    }

                    if (waterFastness_Detail_Detail.ResultPolyester == null)
                    {
                        waterFastness_Detail_Detail.ResultPolyester = string.Empty;
                    }

                    if (waterFastness_Detail_Detail.ResultAcrylic == null)
                    {
                        waterFastness_Detail_Detail.ResultAcrylic = string.Empty;
                    }

                    if (waterFastness_Detail_Detail.ResultWool == null)
                    {
                        waterFastness_Detail_Detail.ResultWool = string.Empty;
                    }


                    if (MyUtility.Check.Empty(waterFastness_Detail_Detail.ResultChange
                        + waterFastness_Detail_Detail.ResultAcetate
                        + waterFastness_Detail_Detail.ResultCotton
                        + waterFastness_Detail_Detail.ResultNylon
                        + waterFastness_Detail_Detail.ResultPolyester
                        + waterFastness_Detail_Detail.ResultAcrylic
                        + waterFastness_Detail_Detail.ResultWool

                        ))
                    {
                        waterFastness_Detail_Detail.Result = string.Empty;
                        continue;
                    }

                    if (waterFastness_Detail_Detail.ResultChange.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultAcetate.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultCotton.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultNylon.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultPolyester.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultAcrylic.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultWool.ToUpper() == "FAIL" ||
                        waterFastness_Detail_Detail.ResultChange == string.Empty ||
                        waterFastness_Detail_Detail.ResultAcetate == string.Empty ||
                        waterFastness_Detail_Detail.ResultCotton == string.Empty ||
                        waterFastness_Detail_Detail.ResultNylon == string.Empty ||
                        waterFastness_Detail_Detail.ResultPolyester == string.Empty ||
                        waterFastness_Detail_Detail.ResultAcrylic == string.Empty ||
                        waterFastness_Detail_Detail.ResultWool == string.Empty
                        )
                    {
                        waterFastness_Detail_Detail.Result = "Fail";
                    }
                    else
                    {
                        waterFastness_Detail_Detail.Result = "Pass";
                    }
                }

                if (waterFastness_Detail_Result.Main.TestBeforePicture == null)
                {
                    waterFastness_Detail_Result.Main.TestBeforePicture = new byte[0];
                }

                if (waterFastness_Detail_Result.Main.TestAfterPicture == null)
                {
                    waterFastness_Detail_Result.Main.TestAfterPicture = new byte[0];
                }

                if (string.IsNullOrEmpty(waterFastness_Detail_Result.Main.TestNo))
                {
                    _WaterFastnessProvider.AddWaterFastnessDetail(waterFastness_Detail_Result, userID, out string TestNo);
                    baseResult.ErrorMessage = TestNo;
                }
                else
                {
                    _WaterFastnessProvider.EditWaterFastnessDetail(waterFastness_Detail_Result, userID);
                }


            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult SaveWaterFastnessMain(WaterFastness_Main waterFastness_Main)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                _WaterFastnessProvider.SaveWaterFastnessMain(waterFastness_Main);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public SendMail_Result SendMail(string toAddress, string ccAddress, string poID, string TestNo, bool isTest, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _WaterFastnessProvider.GetFailMailContentData(poID, TestNo);
                string ID = dtResult.Rows[0]["ID"].ToString();
                dtResult.Columns.Remove("ID");
                string name = $"Water Fastness Test_{poID}_" +
                        $"{dtResult.Rows[0]["Style"]}_" +
                        $"{dtResult.Rows[0]["Article"]}_" +
                        $"{dtResult.Rows[0]["Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport(ID, out string PDFFileName, true, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", PDFFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = $"Water Fastness Test/{poID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Article"]}/" +
                        $"{dtResult.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Water Fastness Test",
                    OrderID = poID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out System.Net.Mail.AlternateView plainView);
                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.ToString();
            }

            return result;
        }

        private void AddImageToWorksheet(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
        {
            if (imageData != null)
            {
                using (var stream = new MemoryStream(imageData))
                {
                    worksheet.AddPicture(stream)
                             .MoveTo(worksheet.Cell(row, col), 5, 5)
                             .WithSize(width, height);
                }
            }
        }
        public BaseResult ToReport(string ID, out string FileName, bool isPDF, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<WaterFastness_Excel> dataList;

            string tmpName = string.Empty;
            FileName = string.Empty;

            try
            {
                dataList = _WaterFastnessProvider.GetExcel(ID).ToList();

                if (!dataList.Any())
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                tmpName = $"Water Fastness Test_{dataList.FirstOrDefault()?.POID}_" +
                          $"{dataList.FirstOrDefault()?.StyleID}_" +
                          $"{dataList.FirstOrDefault()?.Article}_" +
                          $"{dataList.FirstOrDefault()?.AllResult}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[/:?""<>|*%]", string.Empty);
                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                string baseFileName = "WaterFastness_ToExcel";
                string baseFilePath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "XLT", $"{baseFileName}.xltx");

                string outputDirectory = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP");
                string filePath = Path.Combine(outputDirectory, $"{tmpName}.xlsx");
                string pdfPath = Path.Combine(outputDirectory, $"{tmpName}.pdf");

                using (var workbook = new XLWorkbook(baseFilePath))
                {
                    var templateSheet = workbook.Worksheet(1);

                    // 複製分頁：每筆資料一個分頁
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        var data = dataList[i];
                        var currentSheet = i == 0 ? templateSheet : templateSheet.CopyTo($"Sheet_{i + 1}");

                        var testCode = _QualityBrandTestCodeProvider.Get(data.BrandID, "Water Fastness Test");
                        if (testCode.Any())
                        {
                            currentSheet.Cell("A1").Value = $"Water Fastness Test Report({testCode.FirstOrDefault()?.TestCode})";
                        }

                        // 填寫資料
                        currentSheet.Cell("C2").Value = data.ReportNo;
                        currentSheet.Cell("C3").Value = data.SubmitDate?.ToString("yyyy/MM/dd");
                        currentSheet.Cell("H3").Value = DateTime.Now.ToString("yyyy/MM/dd");

                        currentSheet.Cell("C4").Value = data.SeasonID;
                        currentSheet.Cell("H4").Value = data.BrandID;

                        currentSheet.Cell("C5").Value = data.StyleID;
                        currentSheet.Cell("H5").Value = data.POID;

                        currentSheet.Cell("C6").Value = data.Roll;
                        currentSheet.Cell("H6").Value = data.Dyelot;

                        currentSheet.Cell("C7").Value = data.SCIRefno_Color;
                        currentSheet.Cell("C9").Value = data.Temperature;
                        currentSheet.Cell("H9").Value = data.Time;

                        currentSheet.Cell("B13").Value = data.ChangeScale;
                        currentSheet.Cell("C13").Value = data.AcetateScale;
                        currentSheet.Cell("D13").Value = data.CottonScale;
                        currentSheet.Cell("E13").Value = data.NylonScale;
                        currentSheet.Cell("F13").Value = data.PolyesterScale;
                        currentSheet.Cell("G13").Value = data.AcrylicScale;
                        currentSheet.Cell("H13").Value = data.WoolScale;

                        currentSheet.Cell("B14").Value = data.ResultChange;
                        currentSheet.Cell("C14").Value = data.ResultAcetate;
                        currentSheet.Cell("D14").Value = data.ResultCotton;
                        currentSheet.Cell("E14").Value = data.ResultNylon;
                        currentSheet.Cell("F14").Value = data.ResultPolyester;
                        currentSheet.Cell("G14").Value = data.ResultAcrylic;
                        currentSheet.Cell("H14").Value = data.ResultWool;

                        currentSheet.Cell("B15").Value = data.Remark;
                        currentSheet.Cell("C71").Value = data.Inspector;
                        currentSheet.Cell("G71").Value = data.Inspector;

                        // 添加圖片
                        AddImageToWorksheet(currentSheet, data.TestBeforePicture, 46, 1, 380, 300);
                        AddImageToWorksheet(currentSheet, data.TestAfterPicture, 46, 5, 380, 300);
                    }

                    var worksheet = workbook.Worksheet(1);
                    // Excel 合併 + 塞資料 + 重設列印範圍
                    #region Title
                    string FactoryNameEN = _WaterFastnessProvider.GetFactoryNameEN(ID, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);
                    // 2. 複製格式到新插入的列
                    worksheet.Range("A1:H1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    // 設置活動儲存格（指標位置）
                    worksheet.Cell("A1").SetActive();
                    #endregion

                    // 儲存 Excel 檔案
                    workbook.SaveAs(filePath);
                }

                // PDF 轉換
                if (isPDF)
                {
                    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    //officeService.ConvertExcelToPdf(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    ConvertToPDF.ExcelToPDF(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", pdfPath));
                    FileName = Path.GetFileName(pdfPath);
                }
                else
                {
                    FileName = Path.GetFileName(filePath);
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public BaseResult DeleteWaterFastness(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                _WaterFastnessProvider.DeleteWaterFastness(poID, TestNo);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }
    }
}
