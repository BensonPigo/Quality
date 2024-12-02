using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
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
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;


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

                tmpName = $"Perspiration Fastness Test_{dataList.FirstOrDefault().POID}_" +
                       $"{dataList.FirstOrDefault().StyleID}_" +
                       $"{dataList.FirstOrDefault().Article}_" +
                       $"{dataList.FirstOrDefault().AllResult}_" +
                       $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                string basefileName = "PerspirationFastness_ToExcel";
                string openfilepath;

                if (isTest)
                {
                    openfilepath = $"C:\\Willy_Repository\\Quality_KPI\\Quality\\Quality\\bin\\XLT\\{basefileName}.xltx";
                }
                else
                {
                    openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
                }

                Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);
                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                Excel.Worksheet worksheetn;
                // 複製分頁：表身幾筆，就幾個sheet
                for (int j = 1; j < dataList.Count; j++)
                {
                    //Excel.Worksheet worksheetFirst = excel.Worksheets[1];
                    worksheetn = (Excel.Worksheet)excel.ActiveWorkbook.Worksheets[j];

                    worksheet.Copy(worksheetn);
                }
                //開始填資料
                for (int j = 1; j <= dataList.Count; j++)
                {
                    Excel.Worksheet currenSheet = excel.ActiveWorkbook.Worksheets[j];
                    currenSheet.Name = j.ToString();
                    PerspirationFastness_Excel currenData = dataList[j - 1];
                    var testCode = _QualityBrandTestCodeProvider.Get(currenData.BrandID, "Perspiration Fastness Test");
                    if (testCode.Any())
                    {
                        currenSheet.Cells[1, 1] = $@"Perspiration Fastness Test Report({testCode.FirstOrDefault().TestCode})";
                    }
                    currenSheet.Cells[2, 3] = currenData.ReportNo;
                    currenSheet.Cells[3, 3] = currenData.SubmitDate.HasValue ? currenData.SubmitDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                    currenSheet.Cells[3, 8] = DateTime.Now.ToString("yyyy/MM/dd");

                    currenSheet.Cells[4, 3] = currenData.SeasonID;
                    currenSheet.Cells[4, 8] = currenData.BrandID;

                    currenSheet.Cells[5, 3] = currenData.StyleID;
                    currenSheet.Cells[5, 8] = currenData.POID;

                    currenSheet.Cells[6, 3] = currenData.Roll;
                    currenSheet.Cells[6, 8] = currenData.Dyelot;

                    currenSheet.Cells[7, 3] = currenData.SCIRefno_Color;
                    currenSheet.Cells[8, 3] = currenData.MetalContent;

                    // Test Request
                    currenSheet.Cells[10, 3] = currenData.Temperature;
                    currenSheet.Cells[10, 8] = currenData.Time;

                    currenSheet.Cells[14, 2] = currenData.AlkalineChangeScale;
                    currenSheet.Cells[14, 3] = currenData.AlkalineAcetateScale;
                    currenSheet.Cells[14, 4] = currenData.AlkalineCottonScale;
                    currenSheet.Cells[14, 5] = currenData.AlkalineNylonScale;
                    currenSheet.Cells[14, 6] = currenData.AlkalinePolyesterScale;
                    currenSheet.Cells[14, 7] = currenData.AlkalineAcrylicScale;
                    currenSheet.Cells[14, 8] = currenData.AlkalineWoolScale;

                    currenSheet.Cells[15, 2] = currenData.AlkalineResultChange;
                    currenSheet.Cells[15, 3] = currenData.AlkalineResultAcetate;
                    currenSheet.Cells[15, 4] = currenData.AlkalineResultCotton;
                    currenSheet.Cells[15, 5] = currenData.AlkalineResultNylon;
                    currenSheet.Cells[15, 6] = currenData.AlkalineResultPolyester;
                    currenSheet.Cells[15, 7] = currenData.AlkalineResultAcrylic;
                    currenSheet.Cells[15, 8] = currenData.AlkalineResultWool;

                    currenSheet.Cells[19, 2] = currenData.AcidChangeScale;
                    currenSheet.Cells[19, 3] = currenData.AcidAcetateScale;
                    currenSheet.Cells[19, 4] = currenData.AcidCottonScale;
                    currenSheet.Cells[19, 5] = currenData.AcidNylonScale;
                    currenSheet.Cells[19, 6] = currenData.AcidPolyesterScale;
                    currenSheet.Cells[19, 7] = currenData.AcidAcrylicScale;
                    currenSheet.Cells[19, 8] = currenData.AcidWoolScale;

                    currenSheet.Cells[20, 2] = currenData.AcidResultChange;
                    currenSheet.Cells[20, 3] = currenData.AcidResultAcetate;
                    currenSheet.Cells[20, 4] = currenData.AcidResultCotton;
                    currenSheet.Cells[20, 5] = currenData.AcidResultNylon;
                    currenSheet.Cells[20, 6] = currenData.AcidResultPolyester;
                    currenSheet.Cells[20, 7] = currenData.AcidResultAcrylic;
                    currenSheet.Cells[20, 8] = currenData.AcidResultWool;

                    currenSheet.Cells[21, 2] = currenData.Remark;
                    currenSheet.Cells[77, 3] = currenData.Inspector;
                    currenSheet.Cells[77, 7] = currenData.Inspector;


                    #region 添加圖片
                    Excel.Range cellBeforePicture = currenSheet.Cells[52, 1];
                    if (currenData.TestBeforePicture != null)
                    {
                        string imgPath = ToolKit.PublicClass.AddImageSignWord(currenData.TestBeforePicture, currenData.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: isTest);
                        currenSheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 380, 300);
                    }

                    Excel.Range cellAfterPicture = currenSheet.Cells[52, 5];
                    if (currenData.TestAfterPicture != null)
                    {
                        string imgPath = ToolKit.PublicClass.AddImageSignWord(currenData.TestAfterPicture, currenData.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: isTest);
                        currenSheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 380, 300);
                    }
                    #endregion

                }

                #region Save & Show Excel

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

                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepath;
                string filepathpdf;
                if (isTest)
                {
                    filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx);
                    filepathpdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF);
                }
                else
                {
                    filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);
                    filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                }

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);

                FileName = filexlsx;
                result.Result = true;

                if (isPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                    {
                        FileName = fileNamePDF;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }
                }

                #endregion

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
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
