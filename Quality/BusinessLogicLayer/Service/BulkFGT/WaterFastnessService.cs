using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
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
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


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

        public SendMail_Result SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _WaterFastnessProvider.GetFailMailContentData(poID, TestNo);
                string ID = dtResult.Rows[0]["ID"].ToString();
                dtResult.Columns.Remove("ID");
                BaseResult baseResult = ToReport(ID, out string PDFFileName, true, isTest);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", PDFFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Oven Test - Test Fail",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    IsShowAIComment = true,
                    AICommentType = "Water Fastness Test",
                    OrderID = poID,
                };

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, out System.Net.Mail.AlternateView plainView);
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

        private void SetDetailData(Excel.Worksheet worksheet, int setRow, DataRow dr)
        {
            worksheet.Cells[setRow, 2] = dr["Refno"];
            worksheet.Cells[setRow, 3] = dr["Colorid"];
            worksheet.Cells[setRow, 4] = dr["Dyelot"];
            worksheet.Cells[setRow, 6] = dr["Roll"];
            worksheet.Cells[setRow, 7] = dr["Changescale"];
            worksheet.Cells[setRow, 8] = dr["ResultChange"];
            worksheet.Cells[setRow, 9] = dr["Acetatescale"];
            worksheet.Cells[setRow, 10] = dr["ResultAcetate"];
            worksheet.Cells[setRow, 11] = dr["Cottonscale"];
            worksheet.Cells[setRow, 12] = dr["ResultCotton"];
            worksheet.Cells[setRow, 13] = dr["Nylonscale"];
            worksheet.Cells[setRow, 14] = dr["ResultNylon"];
            worksheet.Cells[setRow, 15] = dr["Polyesterscale"];
            worksheet.Cells[setRow, 16] = dr["ResultPolyester"];
            worksheet.Cells[setRow, 17] = dr["Acrylicscale"];
            worksheet.Cells[setRow, 18] = dr["ResultAcrylic"];
            worksheet.Cells[setRow, 19] = dr["Woolscale"];
            worksheet.Cells[setRow, 20] = dr["ResultWool"];
            worksheet.Cells[setRow, 21] = MyUtility.Convert.GetString(dr["Temperature"]) + "˚C";
            worksheet.Cells[setRow, 22] = MyUtility.Convert.GetString(dr["Time"]) + " hrs";
            worksheet.Cells[setRow, 23] = dr["Remark"];
        }

        public BaseResult ToReport(string ID, out string FileName, bool isPDF, bool isTest = false)
        {
            BaseResult result = new BaseResult();
            _WaterFastnessProvider = new WaterFastnessProvider(Common.ProductionDataAccessLayer);
            List<WaterFastness_Excel> dataList = new List<WaterFastness_Excel>();

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

                string basefileName = "WaterFastness_ToExcel";
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
                    WaterFastness_Excel currenData = dataList[j - 1];
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

                    // Test Request
                    currenSheet.Cells[9, 3] = currenData.Temperature;
                    currenSheet.Cells[9, 8] = currenData.Time;

                    currenSheet.Cells[13, 2] = currenData.ChangeScale;
                    currenSheet.Cells[13, 3] = currenData.AcetateScale;
                    currenSheet.Cells[13, 4] = currenData.CottonScale;
                    currenSheet.Cells[13, 5] = currenData.NylonScale;
                    currenSheet.Cells[13, 6] = currenData.PolyesterScale;
                    currenSheet.Cells[13, 7] = currenData.AcrylicScale;
                    currenSheet.Cells[13, 8] = currenData.WoolScale;

                    currenSheet.Cells[14, 2] = currenData.ResultChange;
                    currenSheet.Cells[14, 3] = currenData.ResultAcetate;
                    currenSheet.Cells[14, 4] = currenData.ResultCotton;
                    currenSheet.Cells[14, 5] = currenData.ResultNylon;
                    currenSheet.Cells[14, 6] = currenData.ResultPolyester;
                    currenSheet.Cells[14, 7] = currenData.ResultAcrylic;
                    currenSheet.Cells[14, 8] = currenData.ResultWool;

                    currenSheet.Cells[15, 2] = currenData.Remark;
                    currenSheet.Cells[71, 3] = currenData.Inspector;
                    currenSheet.Cells[71, 7] = currenData.Inspector;


                    #region 添加圖片
                    Excel.Range cellBeforePicture = currenSheet.Cells[46, 1];
                    if (currenData.TestBeforePicture != null)
                    {
                        string imgPath = ToolKit.PublicClass.AddImageSignWord(currenData.TestBeforePicture, currenData.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: isTest);
                        currenSheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 380, 300);
                    }

                    Excel.Range cellAfterPicture = currenSheet.Cells[46, 5];
                    if (currenData.TestAfterPicture != null)
                    {
                        string imgPath = ToolKit.PublicClass.AddImageSignWord(currenData.TestAfterPicture, currenData.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: isTest);
                        currenSheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 380, 300);
                    }
                    #endregion

                }

                #region Save & Show Excel

                string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                string filexlsx = fileName + ".xlsx";
                string fileNamePDF = fileName + ".pdf";

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
