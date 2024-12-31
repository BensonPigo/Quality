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
using System.Net.Mail;
using System.Web;


namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PerspirationFastnessService
    {
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];
        public PerspirationFastnessProvider _PerspirationFastnessProvider;
        public IScaleProvider _ScaleProvider;
        public IOrdersProvider _OrdersProvider;
        public IStyleProvider _StyleProvider;
        private MailToolsService _MailService;
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public BaseResult AmendPerspirationFastnessDetail(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = _PerspirationFastnessProvider.GetPerspirationFastness_Detail(poID, TestNo, string.Empty);

                if (PerspirationFastness_Detail_Result.Main.Status != "Confirmed")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {PerspirationFastness_Detail_Result.Main.Status}, can not Amend";
                    return baseResult;
                }

                _PerspirationFastnessProvider.AmendPerspirationFastness(poID, TestNo);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult EncodePerspirationFastnessDetail(string poID, string TestNo, out string PerspirationFastnessResult)
        {
            BaseResult baseResult = new BaseResult();
            _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
            PerspirationFastnessResult = string.Empty;
            try
            {
                PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = _PerspirationFastnessProvider.GetPerspirationFastness_Detail(poID, TestNo);

                if (PerspirationFastness_Detail_Result.Main.Status != "New")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {PerspirationFastness_Detail_Result.Main.Status}, can not Encode";
                    return baseResult;
                }

                if (PerspirationFastness_Detail_Result.Details.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Data is empty please fill-in data first.";
                    return baseResult;
                }

                if (PerspirationFastness_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.PerspirationFastnessGroup)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Group cannot be empty.";
                    return baseResult;
                }

                if (PerspirationFastness_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.SEQ)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Seq cannot be empty.";
                    return baseResult;
                }

                if (PerspirationFastness_Detail_Result.Details.Any(s =>
                    string.IsNullOrEmpty(s.AlkalineChangeScale) ||
                    string.IsNullOrEmpty(s.AlkalineAcetateScale) ||
                    string.IsNullOrEmpty(s.AlkalineCottonScale) ||
                    string.IsNullOrEmpty(s.AlkalineNylonScale) ||
                    string.IsNullOrEmpty(s.AlkalinePolyesterScale) ||
                    string.IsNullOrEmpty(s.AlkalineAcrylicScale) ||
                    string.IsNullOrEmpty(s.AlkalineWoolScale) ||
                    string.IsNullOrEmpty(s.AcidChangeScale) ||
                    string.IsNullOrEmpty(s.AcidAcetateScale) ||
                    string.IsNullOrEmpty(s.AcidCottonScale) ||
                    string.IsNullOrEmpty(s.AcidNylonScale) ||
                    string.IsNullOrEmpty(s.AcidPolyesterScale) ||
                    string.IsNullOrEmpty(s.AcidAcrylicScale) ||
                    string.IsNullOrEmpty(s.AcidWoolScale) ||
                    string.IsNullOrEmpty(s.Result)
                ))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Color Change Scale, Color Staining Scale and Result cannot be empty.";
                    return baseResult;
                }

                string result = PerspirationFastness_Detail_Result.Details.Any(s => s.Result == "Fail") ? "Fail" : "Pass";
                PerspirationFastnessResult = result;
                _PerspirationFastnessProvider.EncodePerspirationFastness(poID, TestNo, result);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
                return baseResult;
            }
        }

        public PerspirationFastness_Detail_Result GetPerspirationFastness_Detail_Result(string poID, string TestNo, string BrandID)
        {
            try
            {
                PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = new PerspirationFastness_Detail_Result();
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                PerspirationFastness_Detail_Result = _PerspirationFastnessProvider.GetPerspirationFastness_Detail(poID, TestNo, BrandID);

                PerspirationFastness_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();

                if (!string.IsNullOrEmpty(TestNo))
                {
                    DataTable dtResult = _PerspirationFastnessProvider.GetFailMailContentData(poID, TestNo);
                    string Subject = $"Perspiration Fastness Test/{dtResult.Rows[0]["SP#"]}/" +
                        $"{dtResult.Rows[0]["Article"]}/" +
                        $"{dtResult.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    PerspirationFastness_Detail_Result.Main.MailSubject = Subject;
                }

                return PerspirationFastness_Detail_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public PerspirationFastness_Result GetPerspirationFastness_Result(string POID)
        {
            PerspirationFastness_Result result = new PerspirationFastness_Result();
            try
            {
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                result = _PerspirationFastnessProvider.GetPerspirationFastness_Main(POID);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public BaseResult SavePerspirationFastnessDetail(PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                if (string.IsNullOrEmpty(PerspirationFastness_Detail_Result.Main.Article))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Article cannot be empty.";
                    return baseResult;
                }

                if (string.IsNullOrEmpty(PerspirationFastness_Detail_Result.Main.Inspector))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Inspector cannot be empty.";
                    return baseResult;
                }

                if (PerspirationFastness_Detail_Result.Main.InspDate == null)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Test Date cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = PerspirationFastness_Detail_Result
                   .Details.GroupBy(s => new
                   {
                       s.PerspirationFastnessGroup,
                       s.SEQ,
                   })
                   .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[OvenGroup]{s.Key.PerspirationFastnessGroup}, [SEQ]{s.Key.SEQ}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }

                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);

                //再檢查一次Result
                foreach (PerspirationFastness_Detail_Detail PerspirationFastness_Detail_Detail in PerspirationFastness_Detail_Result.Details)
                {
                    if (PerspirationFastness_Detail_Detail.AlkalineResultChange == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultChange = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AlkalineResultAcetate == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultAcetate = string.Empty;
                    }
                    if (PerspirationFastness_Detail_Detail.AlkalineResultCotton == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultCotton = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AlkalineResultNylon == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultNylon = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AlkalineResultPolyester == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultPolyester = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AlkalineResultAcrylic == null)
                    {
                        PerspirationFastness_Detail_Detail.AlkalineResultAcrylic = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultWool == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultWool = string.Empty;
                    }
                    if (PerspirationFastness_Detail_Detail.AcidResultChange == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultChange = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultAcetate == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultAcetate = string.Empty;
                    }
                    if (PerspirationFastness_Detail_Detail.AcidResultCotton == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultCotton = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultNylon == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultNylon = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultPolyester == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultPolyester = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultAcrylic == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultAcrylic = string.Empty;
                    }

                    if (PerspirationFastness_Detail_Detail.AcidResultWool == null)
                    {
                        PerspirationFastness_Detail_Detail.AcidResultWool = string.Empty;
                    }


                    if (MyUtility.Check.Empty(PerspirationFastness_Detail_Detail.AlkalineResultChange
                        + PerspirationFastness_Detail_Detail.AlkalineResultAcetate
                        + PerspirationFastness_Detail_Detail.AlkalineResultCotton
                        + PerspirationFastness_Detail_Detail.AlkalineResultNylon
                        + PerspirationFastness_Detail_Detail.AlkalineResultPolyester
                        + PerspirationFastness_Detail_Detail.AlkalineResultAcrylic
                        + PerspirationFastness_Detail_Detail.AlkalineResultWool
                        + PerspirationFastness_Detail_Detail.AcidResultChange
                        + PerspirationFastness_Detail_Detail.AcidResultAcetate
                        + PerspirationFastness_Detail_Detail.AcidResultCotton
                        + PerspirationFastness_Detail_Detail.AcidResultNylon
                        + PerspirationFastness_Detail_Detail.AcidResultPolyester
                        + PerspirationFastness_Detail_Detail.AcidResultAcrylic
                        + PerspirationFastness_Detail_Detail.AcidResultWool

                        ))
                    {
                        PerspirationFastness_Detail_Detail.Result = string.Empty;
                        continue;
                    }

                    if (PerspirationFastness_Detail_Detail.AlkalineResultChange.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultAcetate.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultCotton.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultNylon.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultPolyester.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultAcrylic.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultWool.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AlkalineResultChange == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultAcetate == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultCotton == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultNylon == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultPolyester == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultAcrylic == string.Empty ||
                        PerspirationFastness_Detail_Detail.AlkalineResultWool == string.Empty ||

                        PerspirationFastness_Detail_Detail.AcidResultChange.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultAcetate.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultCotton.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultNylon.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultPolyester.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultAcrylic.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultWool.ToUpper() == "FAIL" ||
                        PerspirationFastness_Detail_Detail.AcidResultChange == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultAcetate == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultCotton == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultNylon == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultPolyester == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultAcrylic == string.Empty ||
                        PerspirationFastness_Detail_Detail.AcidResultWool == string.Empty
                        )
                    {
                        PerspirationFastness_Detail_Detail.Result = "Fail";
                    }
                    else
                    {
                        PerspirationFastness_Detail_Detail.Result = "Pass";
                    }
                }

                if (PerspirationFastness_Detail_Result.Main.TestBeforePicture == null)
                {
                    PerspirationFastness_Detail_Result.Main.TestBeforePicture = new byte[0];
                }

                if (PerspirationFastness_Detail_Result.Main.TestAfterPicture == null)
                {
                    PerspirationFastness_Detail_Result.Main.TestAfterPicture = new byte[0];
                }

                if (string.IsNullOrEmpty(PerspirationFastness_Detail_Result.Main.TestNo))
                {
                    _PerspirationFastnessProvider.AddPerspirationFastnessDetail(PerspirationFastness_Detail_Result, userID, out string TestNo);
                    baseResult.ErrorMessage = TestNo;
                }
                else
                {
                    _PerspirationFastnessProvider.EditPerspirationFastnessDetail(PerspirationFastness_Detail_Result, userID);
                }


            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult SavePerspirationFastnessMain(PerspirationFastness_Main PerspirationFastness_Main)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                _PerspirationFastnessProvider.SavePerspirationFastnessMain(PerspirationFastness_Main);
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
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _PerspirationFastnessProvider.GetFailMailContentData(poID, TestNo);
                string ID = dtResult.Rows[0]["ID"].ToString();
                dtResult.Columns.Remove("ID");

                string name = $"Perspiration Fastness Test_{dtResult.Rows[0]["SP#"]}_" +
                        $"{dtResult.Rows[0]["Article"]}_" +
                        $"{dtResult.Rows[0]["Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport(ID, out string PDFFileName, true, isTest, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", PDFFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,

                    Subject = $"Perspiration Fastness Test/{dtResult.Rows[0]["SP#"]}/" +
                        $"{dtResult.Rows[0]["Article"]}/" +
                        $"{dtResult.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Perspiration Fastness Test",
                    OrderID = poID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }


                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out AlternateView plainView);

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

        public BaseResult ToReport(string ID, out string FileName, bool isPDF, bool isTest = false, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<PerspirationFastness_Excel> dataList = new List<PerspirationFastness_Excel>();

            string tmpName = string.Empty;
            FileName = string.Empty;

            try
            {
                dataList = _PerspirationFastnessProvider.GetExcel(ID).ToList();

                if (!dataList.Any())
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                tmpName = $"Perspiration Fastness Test_{dataList.FirstOrDefault().POID}_{dataList.FirstOrDefault().StyleID}_{dataList.FirstOrDefault().Article}_{dataList.FirstOrDefault().AllResult}_{DateTime.Now:yyyyMMddHHmmss}";

                string basePath = isTest ? AppDomain.CurrentDomain.BaseDirectory : System.Web.HttpContext.Current.Server.MapPath("~/");
                string xltPath = Path.Combine(basePath, "XLT", "PerspirationFastness_ToExcel.xltx");
                string tmpPath = Path.Combine(basePath, "TMP");

                if (!Directory.Exists(tmpPath)) Directory.CreateDirectory(tmpPath);

                if (!File.Exists(xltPath)) throw new FileNotFoundException("Template not found", xltPath);

                using (var workbook = new XLWorkbook(xltPath))
                {
                    var worksheetTemplate = workbook.Worksheet(1);

                    for (int i = 1; i < dataList.Count; i++)
                    {
                        var copiedSheet = worksheetTemplate.CopyTo($"Sheet{i + 1}");
                        //var newSheet = workbook.AddWorksheet($"Sheet{i + 1}");
                        //worksheetTemplate.CopyTo(newSheet.Name);
                    }

                    for (int i = 0; i < dataList.Count; i++)
                    {
                        var currentSheet = workbook.Worksheet(i + 1);
                        var currentData = dataList[i];

                        var testCode = _QualityBrandTestCodeProvider.Get(currentData.BrandID, "Perspiration Fastness Test");
                        if (testCode.Any())
                        {
                            currentSheet.Cell(1, 1).Value = $"Perspiration Fastness Test Report({testCode.FirstOrDefault().TestCode})";
                        }

                        currentSheet.Cell(2, 3).Value = currentData.ReportNo;
                        currentSheet.Cell(3, 3).Value = currentData.SubmitDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                        currentSheet.Cell(3, 8).Value = DateTime.Now.ToString("yyyy/MM/dd");
                        currentSheet.Cell(4, 3).Value = currentData.SeasonID;
                        currentSheet.Cell(4, 8).Value = currentData.BrandID;
                        currentSheet.Cell(5, 3).Value = currentData.StyleID;
                        currentSheet.Cell(5, 8).Value = currentData.POID;
                        currentSheet.Cell(6, 3).Value = currentData.Roll;
                        currentSheet.Cell(6, 8).Value = currentData.Dyelot;
                        currentSheet.Cell(7, 3).Value = currentData.SCIRefno_Color;
                        currentSheet.Cell(8, 3).Value = currentData.MetalContent;
                        currentSheet.Cell(10, 3).Value = currentData.Temperature;
                        currentSheet.Cell(10, 8).Value = currentData.Time;

                        currentSheet.Cell(14, 2).Value = currentData.AlkalineChangeScale;
                        currentSheet.Cell(14, 3).Value = currentData.AlkalineAcetateScale;
                        currentSheet.Cell(14, 4).Value = currentData.AlkalineCottonScale;
                        currentSheet.Cell(14, 5).Value = currentData.AlkalineNylonScale;
                        currentSheet.Cell(14, 6).Value = currentData.AlkalinePolyesterScale;
                        currentSheet.Cell(14, 7).Value = currentData.AlkalineAcrylicScale;
                        currentSheet.Cell(14, 8).Value = currentData.AlkalineWoolScale;

                        currentSheet.Cell(15, 2).Value = currentData.AlkalineResultChange;
                        currentSheet.Cell(15, 3).Value = currentData.AlkalineResultAcetate;
                        currentSheet.Cell(15, 4).Value = currentData.AlkalineResultCotton;
                        currentSheet.Cell(15, 5).Value = currentData.AlkalineResultNylon;
                        currentSheet.Cell(15, 6).Value = currentData.AlkalineResultPolyester;
                        currentSheet.Cell(15, 7).Value = currentData.AlkalineResultAcrylic;
                        currentSheet.Cell(15, 8).Value = currentData.AlkalineResultWool;

                        currentSheet.Cell(19, 2).Value = currentData.AcidChangeScale;
                        currentSheet.Cell(19, 3).Value = currentData.AcidAcetateScale;
                        currentSheet.Cell(19, 4).Value = currentData.AcidCottonScale;
                        currentSheet.Cell(19, 5).Value = currentData.AcidNylonScale;
                        currentSheet.Cell(19, 6).Value = currentData.AcidPolyesterScale;
                        currentSheet.Cell(19, 7).Value = currentData.AcidAcrylicScale;
                        currentSheet.Cell(19, 8).Value = currentData.AcidWoolScale;

                        currentSheet.Cell(20, 2).Value = currentData.AcidResultChange;
                        currentSheet.Cell(20, 3).Value = currentData.AcidResultAcetate;
                        currentSheet.Cell(20, 4).Value = currentData.AcidResultCotton;
                        currentSheet.Cell(20, 5).Value = currentData.AcidResultNylon;
                        currentSheet.Cell(20, 6).Value = currentData.AcidResultPolyester;
                        currentSheet.Cell(20, 7).Value = currentData.AcidResultAcrylic;
                        currentSheet.Cell(20, 8).Value = currentData.AcidResultWool;

                        currentSheet.Cell(21, 2).Value = currentData.Remark;
                        currentSheet.Cell(77, 3).Value = currentData.Inspector;
                        currentSheet.Cell(77, 7).Value = currentData.Inspector;

                        AddImageToWorksheet(currentSheet, currentData.TestBeforePicture, 52, 1, 380, 300);
                        AddImageToWorksheet(currentSheet, currentData.TestAfterPicture, 52, 5, 380, 300);
                    }

                    tmpName = RemoveInvalidFileNameChars(tmpName);

                    string filePath = Path.Combine(tmpPath, $"{tmpName}.xlsx");
                    string pdfPath = Path.Combine(tmpPath, $"{tmpName}.pdf");

                    workbook.SaveAs(filePath);

                    FileName = $"{tmpName}.xlsx";
                    result.Result = true;

                    if (isPDF)
                    {
                        LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                        officeService.ConvertExcelToPdf(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                        FileName = $"{tmpName}.pdf";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
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

        private string RemoveInvalidFileNameChars(string input)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                input = input.Replace(c.ToString(), "");
            }
            return input;
        }

        public BaseResult DeletePerspirationFastness(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                _PerspirationFastnessProvider.DeletePerspirationFastness(poID, TestNo);
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
