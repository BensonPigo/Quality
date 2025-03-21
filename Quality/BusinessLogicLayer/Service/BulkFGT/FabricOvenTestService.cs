using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Ocsp;
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
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class FabricOvenTestService : IFabricOvenTestService
    {
        IFabricOvenTestProvider _FabricOvenTestProvider;
        IScaleProvider _ScaleProvider;
        IOrdersProvider _OrdersProvider;
        IStyleProvider _StyleProvider;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;

        private string IsTest = ConfigurationManager.AppSettings["IsTest"];

        public BaseResult AmendFabricOvenTestDetail(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                if (fabricOvenTest_Detail_Result.Main.Status != "Confirmed")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {fabricOvenTest_Detail_Result.Main.Status}, can not Amend";
                    return baseResult;
                }

                _FabricOvenTestProvider.AmendFabricOven(poID, TestNo);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult EncodeFabricOvenTestDetail(string poID, string TestNo, out string ovenTestResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            ovenTestResult = string.Empty;
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                if (fabricOvenTest_Detail_Result.Main.Status != "New")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {fabricOvenTest_Detail_Result.Main.Status}, can not Encode";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Data is empty please fill-in data first.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.OvenGroup)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Group cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.SEQ)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Seq cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s =>
                    string.IsNullOrEmpty(s.ChangeScale) ||
                    string.IsNullOrEmpty(s.StainingScale) ||
                    string.IsNullOrEmpty(s.Result)
                ))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Color Change Scale, Color Staining Scale and Result cannot be empty.";
                    return baseResult;
                }

                string result = fabricOvenTest_Detail_Result.Details.Any(s => s.Result == "Fail") ? "Fail" : "Pass";
                ovenTestResult = result;
                _FabricOvenTestProvider.EncodeFabricOven(poID, TestNo, result);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
                return baseResult;
            }
        }

        public FabricOvenTest_Detail_Result GetFabricOvenTest_Detail_Result(string poID, string TestNo, string BrandID)
        {
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = new FabricOvenTest_Detail_Result();
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo, BrandID);

                fabricOvenTest_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();


                if (!string.IsNullOrEmpty(TestNo))
                {
                    DataTable dtResult = _FabricOvenTestProvider.GetFailMailContentData(poID, TestNo);
                    string Subject = $"Fabric Oven Test/{poID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Article"]}/" +
                        $"{dtResult.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    fabricOvenTest_Detail_Result.Main.MailSubject = Subject;
                }


                return fabricOvenTest_Detail_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public FabricOvenTest_Result GetFabricOvenTest_Result(string POID)
        {
            FabricOvenTest_Result result = new FabricOvenTest_Result();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                result = _FabricOvenTestProvider.GetFabricOvenTest_Main(POID);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }

        public BaseResult SaveFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.Article))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Article cannot be empty.";
                    return baseResult;
                }

                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.Inspector))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Inspector cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Main.InspDate == null)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Test Date cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricOvenTest_Detail_Result
                   .Details.GroupBy(s => new
                   {
                       s.OvenGroup,
                       s.SEQ,
                   })
                   .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[OvenGroup]{s.Key.OvenGroup}, [SEQ]{s.Key.SEQ}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }

                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);

                //再檢查一次Result
                foreach (FabricOvenTest_Detail_Detail fabricOvenTest_Detail_Detail in fabricOvenTest_Detail_Result.Details)
                {
                    if (fabricOvenTest_Detail_Detail.ResultChange == null)
                    {
                        fabricOvenTest_Detail_Detail.ResultChange = string.Empty;
                    }

                    if (fabricOvenTest_Detail_Detail.ResultStain == null)
                    {
                        fabricOvenTest_Detail_Detail.ResultStain = string.Empty;
                    }

                    if (MyUtility.Check.Empty(fabricOvenTest_Detail_Detail.ResultChange + fabricOvenTest_Detail_Detail.ResultStain))
                    {
                        fabricOvenTest_Detail_Detail.Result = string.Empty;
                        continue;
                    }

                    if (fabricOvenTest_Detail_Detail.ResultChange.ToUpper() == "FAIL" ||
                        fabricOvenTest_Detail_Detail.ResultStain.ToUpper() == "FAIL" ||
                        fabricOvenTest_Detail_Detail.ResultChange == string.Empty ||
                        fabricOvenTest_Detail_Detail.ResultStain == string.Empty
                        )
                    {
                        fabricOvenTest_Detail_Detail.Result = "Fail";
                    }
                    else
                    {
                        fabricOvenTest_Detail_Detail.Result = "Pass";
                    }
                }

                if (fabricOvenTest_Detail_Result.Main.TestBeforePicture == null)
                {
                    fabricOvenTest_Detail_Result.Main.TestBeforePicture = new byte[0];
                }

                if (fabricOvenTest_Detail_Result.Main.TestAfterPicture == null)
                {
                    fabricOvenTest_Detail_Result.Main.TestAfterPicture = new byte[0];
                }

                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.TestNo))
                {
                    _FabricOvenTestProvider.AddFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID, out string TestNo);
                    baseResult.ErrorMessage = TestNo;
                }
                else
                {
                    _FabricOvenTestProvider.EditFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID);
                }


            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _FabricOvenTestProvider.SaveFabricOvenTestMain(fabricOvenTest_Main);
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
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricOvenTestProvider.GetFailMailContentData(poID, TestNo);

                string name = $"Fabric Oven Test_{poID}_" +
                    $"{dtResult.Rows[0]["Style"]}_" +
                    $"{dtResult.Rows[0]["Article"]}_" +
                    $"{dtResult.Rows[0]["Result"]}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";


                BaseResult baseResult = ToPdfFabricOvenTestDetail(poID, TestNo, out string pdfFileName, isTest, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", pdfFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    //Subject = "Fabric Oven Test - Test Fail",
                    Subject = $"Fabric Oven Test/{poID}/" +
                    $"{dtResult.Rows[0]["Style"]}/" +
                    $"{dtResult.Rows[0]["Article"]}/" +
                    $"{dtResult.Rows[0]["Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Oven Test",
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

        public BaseResult ToExcelFabricOvenTestDetail2(string poID, string TestNo, out string excelFileName, bool isTest = false)
        {
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");

                DataTable dtOvenDetail = _FabricOvenTestProvider.GetOvenDetailForExcel(poID, TestNo);
                DataTable dtOven = _FabricOvenTestProvider.GetOven(poID, TestNo);

                string[] columnNames = new string[] { "OvenGroup", "SEQ", "Roll", "Dyelot", "SCIRefno", "ColorID", "Supplier", "Changescale", "StainingScale", "Result", "Remark" };
                var ret = Array.CreateInstance(typeof(object), dtOvenDetail.Rows.Count, 11) as object[,];
                for (int i = 0; i < dtOvenDetail.Rows.Count; i++)
                {
                    DataRow row = dtOvenDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

                if (dtOvenDetail.Rows.Count == 0)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string styleID;
                string seasonID;
                string status;
                string brandID;
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = poID }).ToList();

                if (listOrders.Count == 0)
                {
                    styleID = string.Empty;
                    seasonID = string.Empty;
                    status = string.Empty;
                    brandID = string.Empty;
                }
                else
                {
                    styleID = listOrders[0].StyleID;
                    seasonID = listOrders[0].SeasonID;
                    status = dtOven.Rows[0]["status"].ToString();
                    brandID = listOrders[0].BrandID;
                }

                string strXltName = baseFilePath + "\\XLT\\FabricOvenTestDetailReport.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];

                worksheet.Cells[1, 2] = poID;
                worksheet.Cells[1, 4] = styleID;
                worksheet.Cells[1, 6] = seasonID;
                worksheet.Cells[1, 8] = dtOven.Rows[0]["Article"].ToString();
                worksheet.Cells[1, 10] = TestNo;
                worksheet.Cells[2, 2] = status;
                worksheet.Cells[2, 4] = dtOven.Rows[0]["Result"].ToString();
                worksheet.Cells[2, 6] = dtOven.Rows[0]["InspDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["InspDate"]).ToString("yyyy/MM/dd");
                worksheet.Cells[2, 8] = dtOven.Rows[0]["Inspector"].ToString();
                worksheet.Cells[2, 10] = brandID;
                worksheet.Cells[3, 2] = dtOven.Rows[0]["ReportDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["ReportDate"]).ToString("yyyy/MM/dd");

                #region 簽名檔
                string imgPath_Signature = string.Empty;
                if (dtOven.Rows[0]["Signature"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["Signature"];
                    MemoryStream ms = new MemoryStream(bytes);
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_Signature = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_Signature = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath_Signature);
                }

                string imgPath_ApvSignature = string.Empty;
                if (dtOven.Rows[0]["ApvSignature"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["ApvSignature"];
                    MemoryStream ms = new MemoryStream(bytes);
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_ApvSignature = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_ApvSignature = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath_ApvSignature);
                }

                string signature = dtOven.Rows[0]["InspectorName"].ToString();
                string apvsignature = dtOven.Rows[0]["ApproverName"].ToString();
                worksheet.Cells[24, 4] = signature;
                worksheet.Cells[24, 9] = apvsignature;

                Excel.Range cell;
                if (!string.IsNullOrEmpty(imgPath_Signature))
                {
                    cell = worksheet.Cells[26, 4];
                    worksheet.Shapes.AddPicture(imgPath_Signature, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 50, cell.Top + 4, 100, 24);
                }

                if (!string.IsNullOrEmpty(imgPath_ApvSignature))
                {
                    cell = worksheet.Cells[26, 9];
                    worksheet.Shapes.AddPicture(imgPath_ApvSignature, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 20, cell.Top + 4, 100, 24);
                }
                #endregion

                #region 圖片
                Excel.Range cellBefore = worksheet.Cells[8, 1];
                if (dtOven.Rows[0]["TestBeforePicture"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["TestBeforePicture"];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(bytes, dtOven.Rows[0]["ReportNo"].ToString(), ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left + 5, cellBefore.Top + 5, 370, 240);
                }

                Excel.Range cellAfter = worksheet.Cells[8, 7];
                if (dtOven.Rows[0]["TestAfterPicture"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["TestAfterPicture"];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(bytes, dtOven.Rows[0]["ReportNo"].ToString(), ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left + 5, cellAfter.Top + 5, 358, 240);
                }
                #endregion

                int defaultRowCount = 1;
                int detailStartIdx = 5;
                int otherCount = dtOvenDetail.Rows.Count - defaultRowCount;
                if (otherCount > 0)
                {
                    //  複製Row
                    for (int i = 0; i < otherCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = worksheet.get_Range($"A{detailStartIdx + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range($"A{detailStartIdx}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }

                int RowIdx = 0;
                foreach (DataRow dr in dtOvenDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excel.Cells[RowIdx + detailStartIdx, colIndex] = dtOvenDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                worksheet.Cells.EntireColumn.AutoFit();
                worksheet.Cells.EntireRow.AutoFit();

                worksheet.Select();

                excelFileName = $"FabricOvenTestDetailReport{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
        public BaseResult ToExcelFabricOvenTestDetail(string poID, string testNo, out string excelFileName, bool isTest = false)
        {
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {
                // 設定基底目錄
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");
                string templateFilePath = Path.Combine(baseFilePath, "XLT", "FabricOvenTestDetailReport.xltx");
                string tmpDir = Path.Combine(baseFilePath, "TMP");

                if (!Directory.Exists(tmpDir))
                {
                    Directory.CreateDirectory(tmpDir);
                }

                // 獲取資料
                DataTable dtOvenDetail = _FabricOvenTestProvider.GetOvenDetailForExcel(poID, testNo);
                DataTable dtOven = _FabricOvenTestProvider.GetOven(poID, testNo);

                if (dtOvenDetail.Rows.Count == 0)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                // 讀取訂單資訊
                var orders = _OrdersProvider.Get(new Orders { ID = poID }).FirstOrDefault();
                string styleID = orders?.StyleID ?? string.Empty;
                string seasonID = orders?.SeasonID ?? string.Empty;
                string brandID = orders?.BrandID ?? string.Empty;
                string status = dtOven.Rows.Count > 0 ? dtOven.Rows[0]["status"].ToString() : string.Empty;

                // 生成檔名
                excelFileName = $"FabricOvenTestDetailReport_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                string filePath = Path.Combine(tmpDir, excelFileName);

                // 使用 ClosedXML 開啟範本並寫入資料
                using (var workbook = new XLWorkbook(templateFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入基本資料
                    worksheet.Cell(1, 2).Value = poID;
                    worksheet.Cell(1, 4).Value = styleID;
                    worksheet.Cell(1, 6).Value = seasonID;
                    worksheet.Cell(1, 8).Value = dtOven.Rows[0]["Article"]?.ToString();
                    worksheet.Cell(1, 10).Value = testNo;
                    worksheet.Cell(2, 2).Value = status;
                    worksheet.Cell(2, 4).Value = dtOven.Rows[0]["Result"]?.ToString();
                    worksheet.Cell(2, 6).Value = dtOven.Rows[0]["InspDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["InspDate"]).ToString("yyyy/MM/dd");
                    worksheet.Cell(2, 8).Value = dtOven.Rows[0]["Inspector"]?.ToString();
                    worksheet.Cell(2, 10).Value = brandID;
                    worksheet.Cell(3, 2).Value = dtOven.Rows[0]["ReportDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["ReportDate"]).ToString("yyyy/MM/dd");

                    // 簽名檔
                    AddImageToWorksheet(worksheet, dtOven.Rows[0]["Signature"] as byte[], 26, 4, 100, 24);
                    AddImageToWorksheet(worksheet, dtOven.Rows[0]["ApvSignature"] as byte[], 26, 9, 100, 24);
                    worksheet.Cell(24, 4).Value = dtOven.Rows[0]["InspectorName"]?.ToString();
                    worksheet.Cell(24, 9).Value = dtOven.Rows[0]["ApproverName"]?.ToString();

                    // 圖片
                    AddImageToWorksheet(worksheet, dtOven.Rows[0]["TestBeforePicture"] as byte[], 8, 1, 370, 240);
                    AddImageToWorksheet(worksheet, dtOven.Rows[0]["TestAfterPicture"] as byte[], 8, 7, 358, 240);

                    // 表身資料
                    string[] columnNames = { "OvenGroup", "SEQ", "Roll", "Dyelot", "SCIRefno", "ColorID", "Supplier", "Changescale", "StainingScale", "Result", "Remark" };
                    int detailStartIdx = 5;

                    for (int i = 0; i < dtOvenDetail.Rows.Count; i++)
                    {// 第一筆資料不用複製
                        if (i > 0)
                        {
                            // 1. 複製
                            var rowToCopy = worksheet.Row(detailStartIdx);

                            // 2. 插入一列
                            worksheet.Row(detailStartIdx + i).InsertRowsAbove(1);

                            // 3. 複製格式到新插入的列
                            var newRow = worksheet.Row(detailStartIdx + i);

                            rowToCopy.CopyTo(newRow);
                        }
                        for (int j = 0; j < columnNames.Length; j++)
                        {
                            worksheet.Cell(detailStartIdx + i, j + 1).Value = dtOvenDetail.Rows[i][columnNames[j]].ToString();
                        }
                    }


                    #region Title

                    string FactoryNameEN = _FabricOvenTestProvider.GetFactoryNameEN(poID, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);
                    // 2. 合併欄位 (A1:K1)
                    worksheet.Range("A1:K1").Merge();

                    // 3. 設置文字和樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN; // 替換為你的 FactoryNameEN 變數
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    #endregion

                    // 自動調整欄寬
                    worksheet.Columns().AdjustToContents();

                    // 儲存檔案
                    workbook.SaveAs(filePath);
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

        // 圖片處理共用方法
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

        private void SetDetailData(Excel.Worksheet worksheet, int setRow, DataRow dr)
        {
            worksheet.Cells[setRow, 2] = dr["Refno"];
            worksheet.Cells[setRow, 3] = dr["Colorid"];
            worksheet.Cells[setRow, 5] = dr["Dyelot"];
            worksheet.Cells[setRow, 6] = dr["Roll"];
            worksheet.Cells[setRow, 7] = dr["Changescale"];
            worksheet.Cells[setRow, 9] = dr["ResultChange"];
            worksheet.Cells[setRow, 10] = dr["StainingScale"];
            worksheet.Cells[setRow, 11] = dr["ResultStain"];
            worksheet.Cells[setRow, 12] = MyUtility.Convert.GetString(dr["Temperature"]) + "˚C";
            worksheet.Cells[setRow, 13] = MyUtility.Convert.GetString(dr["Time"]) + " hrs";
            worksheet.Cells[setRow, 14] = dr["Remark"];
        }

        public BaseResult ToPdfFabricOvenTestDetail(string poID, string TestNo, out string pdfFileName, bool isTest, string AssignedFineName = "")
        {
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

            BaseResult result = new BaseResult();
            pdfFileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtOvenDetail = _FabricOvenTestProvider.GetOvenDetailForExcel(poID, TestNo);
                DataTable dtOven = _FabricOvenTestProvider.GetOven(poID, TestNo);

                if (dtOvenDetail.Rows.Count < 1)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                var distOvenDetailSubmitDate = dtOvenDetail.AsEnumerable().ToList();

                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = poID }).ToList();
                List<Style> listStyle = _StyleProvider.Get(new Style() { Ukey = listOrders[0].StyleUkey }).ToList();

                string styleUkey = string.Empty;
                string styleID = string.Empty;
                string seasonID = string.Empty;
                string CustPONO = string.Empty;
                string brandID = string.Empty;

                if (listOrders.Count > 0)
                {
                    styleUkey = listOrders[0].StyleUkey.ToString();
                    styleID = listOrders[0].StyleID;
                    seasonID = listOrders[0].SeasonID;
                    CustPONO = listOrders[0].CustPONO;
                    brandID = listOrders[0].BrandID;
                }
                var testCode = _QualityBrandTestCodeProvider.Get(brandID, "Fabric Oven Test");

                tmpName = $"Fabric Oven Test_{poID}_" +
                    $"{styleID}_" +
                    $"{dtOven.Rows[0]["Article"]}_" +
                    $"{dtOven.Rows[0]["Result"]}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                string strXltName = baseFilePath + "\\XLT\\FabricOvenTestDetailReportToPDF.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);

                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // 預設頁在第5頁，前4頁是用來複製的格式，最後在刪除
                int defaultSheet = 5;

                #region Set Picture              
                string imgPath_Signature = string.Empty;
                string imgPath_ApvSignature = string.Empty;
                string imgPath_BeforePicture = string.Empty;
                string imgPath_AfterPicture = string.Empty;
                if (dtOven.Rows[0]["Signature"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["Signature"];
                    MemoryStream ms = new MemoryStream(bytes);
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_Signature = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_Signature = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath_Signature);
                }

                if (dtOven.Rows[0]["ApvSignature"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["ApvSignature"];
                    MemoryStream ms = new MemoryStream(bytes);
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_ApvSignature = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_ApvSignature = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath_ApvSignature);
                }

                if (dtOven.Rows[0]["TestBeforePicture"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["TestBeforePicture"];
                    imgPath_BeforePicture = ToolKit.PublicClass.AddImageSignWord(bytes, dtOven.Rows[0]["ReportNo"].ToString(), ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                }

                //Excel.Range cellAfter = worksheet.Cells[3, 10];
                if (dtOven.Rows[0]["TestAfterPicture"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])dtOven.Rows[0]["TestAfterPicture"];
                    imgPath_AfterPicture = ToolKit.PublicClass.AddImageSignWord(bytes, dtOven.Rows[0]["ReportNo"].ToString(), ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                }
                #endregion

                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];
                // 填入表頭資訊
                worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet];
                if (testCode.Any())
                {
                    worksheet.Cells[1, 2] = $@"Color Migration Test (Oven) Report({testCode.FirstOrDefault().TestCode})";
                }
                worksheet.Cells[4, 3] = distOvenDetailSubmitDate[0]["ReportNo"].ToString();
                worksheet.Cells[4, 9] = dtOven.Rows[0]["ReportDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["ReportDate"]).ToString("yyyy/MM/dd");
                worksheet.Cells[4, 14] = dtOven.Rows[0]["InspDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["InspDate"]).ToString("yyyy/MM/dd");
                worksheet.Cells[5, 3] = poID;
                worksheet.Cells[5, 9] = brandID;
                worksheet.Cells[7, 3] = styleID;
                worksheet.Cells[7, 9] = CustPONO;
                worksheet.Cells[7, 14] = dtOven.Rows[0]["Article"].ToString();
                worksheet.Cells[8, 3] = listStyle[0].StyleName;
                worksheet.Cells[8, 9] = seasonID;

                // 細項
                int headerRow = 10; // 表頭那頁前10列為固定
                int signatureRow = 5; // 簽名有5列

                string signature = dtOven.Rows[0]["InspectorName"].ToString();
                string apvsignature = dtOven.Rows[0]["ApproverName"].ToString();
                Excel.Worksheet worksheetDetail = excel.ActiveWorkbook.Worksheets[1];
                Excel.Worksheet worksheetSignature = excel.ActiveWorkbook.Worksheets[2];
                Excel.Worksheet worksheetPicture = excel.ActiveWorkbook.Worksheets[3];
                Excel.Worksheet worksheetFrame = excel.ActiveWorkbook.Worksheets[4];

                DataRow[] dr = dtOvenDetail.Select();
                worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet];

                for (int j = 0; j < dr.Length; j++)
                {
                    Excel.Range paste = worksheet.get_Range($"A{headerRow + 1 + j}", Type.Missing);
                    Excel.Range r = worksheetDetail.get_Range("A1").EntireRow;
                    paste.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r.Copy(Type.Missing));

                    // 細項資料
                    this.SetDetailData(worksheet, j + headerRow + 1, dr[j]);
                }
                worksheet.get_Range($"N11:O{dr.Length + 11}").WrapText = true;
                worksheet.Cells.EntireRow.AutoFit();

                // 簽名格子塞入後的Row Index
                int afterSignatureRow;

                Excel.Range cell;
                Excel.Range paste1 = worksheet.get_Range($"A{dr.Length + headerRow + 1}", Type.Missing);
                Excel.Range r1 = worksheetSignature.get_Range("A1:A5").EntireRow;
                paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r1.Copy(Type.Missing));
                if (!string.IsNullOrEmpty(imgPath_Signature))
                {
                    cell = worksheet.Cells[dr.Length + headerRow + 3, 2];
                    worksheet.Shapes.AddPicture(imgPath_Signature, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 60, cell.Top + 4, 100, 24);
                }
                worksheet.Cells[dr.Length + headerRow + 5, 2] = signature;

                if (!string.IsNullOrEmpty(imgPath_ApvSignature))
                {
                    cell = worksheet.Cells[dr.Length + headerRow + 3, 14];
                    worksheet.Shapes.AddPicture(imgPath_ApvSignature, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top + 4, 100, 24);
                }
                worksheet.Cells[dr.Length + headerRow + 4, 13] = apvsignature;
                afterSignatureRow = dr.Length + headerRow + signatureRow;

                Excel.Range paste2 = worksheet.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));

                if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                {
                    cell = worksheet.Cells[afterSignatureRow + 3, 2];
                    worksheet.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 420, 280);
                }
                if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                {
                    cell = worksheet.Cells[afterSignatureRow + 3, 10];
                    worksheet.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 420, 280);
                }


                Excel.Range paste3 = worksheet.get_Range($"A{afterSignatureRow + 3 + 20 + 1}", Type.Missing);
                Excel.Range r3 = worksheetFrame.get_Range("A4:A42").EntireRow;
                paste3.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r3.Copy(Type.Missing));

                #region Title

                string FactoryNameEN = _FabricOvenTestProvider.GetFactoryNameEN(poID, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                // 1. 插入一列
                worksheet.Rows["1"].Insert();
                // 2. 合併欄位 (B1:K1)
                Excel.Range mergedRange = worksheet.Range["B1", "O1"];
                mergedRange.Merge();

                // 3. 設置文字和樣式
                mergedRange.Value = FactoryNameEN; // 替換為你的 FactoryNameEN 變數
                mergedRange.Font.Name = "Arial";      // 設置字體類型
                mergedRange.Font.Size = 25;          // 設置字體大小
                mergedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // 水平置中
                mergedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   // 垂直置中
                mergedRange.Font.Bold = true;        // 設置字體加粗
                #endregion

                for (int i = 0; i < 4; i++)
                {
                    worksheet = excel.ActiveWorkbook.Worksheets[1];
                    worksheet.Delete();
                }

                

                #region Save & Show Excel
                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                tmpName = FileNameHelper.SanitizeFileName(tmpName);
                pdfFileName = $"{tmpName}.pdf";
                string excelFileName = $"{tmpName}.xlsx";

                pdfFileName = excelFileName; // 暫時只匯出excel
                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion

                // To PDF
                //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //officeService.ConvertExcelToPdf(excelPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //ConvertToPDF.ExcelToPDF(excelPath, pdfPath);

                result.Result = true;

            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }



            return result;
        }

        // 插入圖片的工具方法
        private void InsertImage(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
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

        public BaseResult DeleteOven(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _FabricOvenTestProvider.DeleteOven(poID, TestNo);
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
