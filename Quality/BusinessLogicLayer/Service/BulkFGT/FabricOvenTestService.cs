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

namespace BusinessLogicLayer.Service
{
    public class FabricOvenTestService : IFabricOvenTestService
    {
        IFabricOvenTestProvider _FabricOvenTestProvider;
        IScaleProvider _ScaleProvider;
        IOrdersProvider _OrdersProvider;
        IStyleProvider _StyleProvider;

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

        public FabricOvenTest_Detail_Result GetFabricOvenTest_Detail_Result(string poID, string TestNo)
        {
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = new FabricOvenTest_Detail_Result();
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                fabricOvenTest_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();
                
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
                        fabricOvenTest_Detail_Detail.ResultStain.ToUpper() == "FAIL"  ||
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

        public SendMail_Result SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricOvenTestProvider.GetFailMailContentData(poID, TestNo);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, out System.Net.Mail.AlternateView plainView);
                SendMail_Request sendMail_Request = new SendMail_Request()
                { 
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Oven Test - Test Fail",
                    Body = mailBody,
                    alternateView = plainView,
                };
                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.ToString();
            }

            return result;
        }

        public BaseResult ToExcelFabricOvenTestDetail(string poID, string TestNo, out string excelFileName, bool isTest = false)
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

                Excel.Range cellBefore = worksheet.Cells[11, 1];
                if (dtOven.Rows[0]["TestBeforePicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestBeforePicture"];
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left + 5, cellBefore.Top + 5, 370, 240);
                }

                Excel.Range cellAfter = worksheet.Cells[11, 7];
                if (dtOven.Rows[0]["TestAfterPicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestAfterPicture"];
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left + 5, cellAfter.Top + 5, 358, 240);
                }

                if (dtOvenDetail.Rows.Count > 0)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A4:K4", Type.Missing).EntireRow;
                    for (int i = 1; i < dtOvenDetail.Rows.Count; i++)
                    {
                        rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    }

                    Marshal.ReleaseComObject(rngToInsert);
                }

                int startRow = 4;
                for (int i = 0; i < dtOvenDetail.Rows.Count; i++)
                {
                    worksheet.Cells[startRow + i, 1] = ret[i, 0];
                    worksheet.Cells[startRow + i, 2] = ret[i, 1];
                    worksheet.Cells[startRow + i, 3] = ret[i, 2];
                    worksheet.Cells[startRow + i, 4] = ret[i, 3];
                    worksheet.Cells[startRow + i, 5] = ret[i, 4];
                    worksheet.Cells[startRow + i, 6] = ret[i, 5];
                    worksheet.Cells[startRow + i, 7] = ret[i, 6];
                    worksheet.Cells[startRow + i, 8] = ret[i, 7];
                    worksheet.Cells[startRow + i, 9] = ret[i, 8];
                    worksheet.Cells[startRow + i, 10] = ret[i, 9];
                    worksheet.Cells[startRow + i, 11] = ret[i, 10];
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

        private void SetDetailData(Excel.Worksheet worksheet, int setRow, DataRow dr)
        {
            worksheet.Cells[setRow, 2] = dr["SubmitDate"] == DBNull.Value ? string.Empty : ((DateTime)dr["SubmitDate"]).ToString("yyyy/MM/dd");
            worksheet.Cells[setRow, 3] = dr["Refno"];
            worksheet.Cells[setRow, 4] = dr["Colorid"];
            worksheet.Cells[setRow, 6] = dr["Dyelot"];
            worksheet.Cells[setRow, 7] = dr["Roll"];
            worksheet.Cells[setRow, 8] = dr["Changescale"];
            worksheet.Cells[setRow, 10] = dr["ResultChange"];
            worksheet.Cells[setRow, 11] = dr["StainingScale"];
            worksheet.Cells[setRow, 12] = dr["ResultStain"];
            worksheet.Cells[setRow, 13] = MyUtility.Convert.GetString(dr["Temperature"]) + "˚C";
            worksheet.Cells[setRow, 14] = MyUtility.Convert.GetString(dr["Time"]) + " hrs";
            worksheet.Cells[setRow, 15] = dr["Remark"];
        }

        public BaseResult ToPdfFabricOvenTestDetail_Ori(string poID, string TestNo, out string pdfFileName, bool isTest)
        {
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);

            BaseResult result = new BaseResult();
            pdfFileName = string.Empty;

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

                var distOvenDetailSubmitDate = dtOvenDetail.AsEnumerable()
                    .Select(s => s["SubmitDate"] == DBNull.Value ? string.Empty : ((DateTime)s["SubmitDate"]).ToString("yyyy/MM/dd"))
                    .Distinct().ToList();

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

                string strXltName = baseFilePath + "\\XLT\\FabricOvenTestDetailReportToPDF.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }
                excel.Visible = true;
                excel.DisplayAlerts = false;

                // 預設頁在第5頁，前4頁是用來複製的格式，最後在刪除
                // 依據 submitDate 複製分頁
                int defaultSheet = 5;
                for (int c = 1; c < distOvenDetailSubmitDate.Count(); c++)
                {
                    Excel.Worksheet worksheetFirst = excel.ActiveWorkbook.Worksheets[defaultSheet];
                    Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + c];
                    worksheetFirst.Copy(worksheetn);
                }

                #region Set Picture
                //Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[4];
                //Excel.Range cellBefore = worksheet.Cells[3, 2];
                string imgPath_BeforePicture = string.Empty;
                string imgPath_AfterPicture = string.Empty;
                if (dtOven.Rows[0]["TestBeforePicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_BeforePicture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_BeforePicture = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestBeforePicture"];
                    using (var imageFile = new FileStream(imgPath_BeforePicture, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    //worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left + 5, cellBefore.Top + 5, 430, 295);
                }

                //Excel.Range cellAfter = worksheet.Cells[3, 10];
                if (dtOven.Rows[0]["TestAfterPicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_AfterPicture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_AfterPicture = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestAfterPicture"];
                    using (var imageFile = new FileStream(imgPath_AfterPicture, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    //worksheet.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left + 5, cellAfter.Top + 5, 430, 295);
                }
                #endregion

                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];
                // 依據 submitDate 填入表頭資訊
                for (int i = 0; i < distOvenDetailSubmitDate.Count(); i++)
                {
                    worksheet = excel.ActiveWorkbook.Worksheets[i + defaultSheet];
                    worksheet.Cells[4, 3] = distOvenDetailSubmitDate[i];
                    worksheet.Cells[4, 6] = dtOven.Rows[0]["InspDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["InspDate"]).ToString("yyyy/MM/dd");
                    worksheet.Cells[4, 9] = poID;
                    worksheet.Cells[4, 14] = brandID;
                    worksheet.Cells[6, 3] = styleID;
                    worksheet.Cells[6, 9] = CustPONO;
                    worksheet.Cells[6, 14] = dtOven.Rows[0]["Article"].ToString();
                    worksheet.Cells[7, 3] = listStyle[0].StyleName;
                    worksheet.Cells[7, 9] = seasonID;
                }

                // 細項
                int setRow = 78; // 78 列為一頁
                int headerRow = 9; // 表頭那頁前9列為固定
                int signatureRow = 4; // 簽名有4列
                int frameRow = 34; // 框 33列 + 1 列空白
                int alladdSheet = 0;

                string signature = dtOven.Rows[0]["InspectorName"].ToString();
                Excel.Worksheet worksheetDetail = excel.ActiveWorkbook.Worksheets[1];
                Excel.Worksheet worksheetSignature = excel.ActiveWorkbook.Worksheets[2];
                Excel.Worksheet worksheetFrame = excel.ActiveWorkbook.Worksheets[3];
                Excel.Worksheet worksheetPicture = excel.ActiveWorkbook.Worksheets[4];

                for (int i = 0; i < distOvenDetailSubmitDate.Count; i++)
                {
                    DataRow[] dr = dtOvenDetail.Select(MyUtility.Check.Empty(distOvenDetailSubmitDate[i]) ? $@"submitDate is null" : $"submitDate = '{distOvenDetailSubmitDate[i]}'");

                    int underHeaderRow = setRow - headerRow;
                    if (dr.Length > underHeaderRow)
                    {
                        int overRow = dr.Length - underHeaderRow;
                        int addSheets = (int)Math.Ceiling(overRow * 1.0 / setRow);

                        // 有表頭那頁下方的細項格線
                        worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                        for (int j = 0; j < underHeaderRow; j++)
                        {
                            Excel.Range paste1 = worksheet.get_Range($"A{headerRow + 1 + j}", Type.Missing);
                            Excel.Range r = worksheetDetail.get_Range("A1").EntireRow;
                            paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r.Copy(Type.Missing));

                            // 細項資料
                            this.SetDetailData(worksheet, j + headerRow + 1, dr[j]);
                        }
                        worksheet.Cells.EntireRow.AutoFit();

                        // 額外細項分頁
                        for (int k = 0; k < addSheets; k++)
                        {
                            // 新增細項分頁
                            Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];
                            worksheetDetail.Copy(worksheetn);

                            #region worksheetn 的細項格線
                            worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];
                            worksheetn.get_Range("A1").EntireRow.Delete();

                            int addrow = overRow;
                            if (overRow > setRow)
                            {
                                addrow = setRow;
                                overRow -= setRow;
                            }
                            else
                            {
                                overRow = 0;
                            }

                            for (int j = 0; j < addrow; j++)
                            {
                                Excel.Range paste1 = worksheetn.get_Range($"A{j + 1}", Type.Missing);
                                Excel.Range r = worksheetDetail.get_Range("A1").EntireRow;
                                paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r.Copy(Type.Missing));

                                // 細項資料
                                this.SetDetailData(worksheetn, j + 1, dr[j + underHeaderRow + (k * setRow)]);
                            }
                            #endregion

                            int afterSignatureRow = 0;

                            // 簽名列
                            if (overRow <= 0)
                            {
                                if (addrow < setRow - signatureRow)
                                {
                                    Excel.Range paste2 = worksheetn.get_Range($"A{addrow + 1}", Type.Missing);
                                    Excel.Range r2 = worksheetSignature.get_Range("A1:A4").EntireRow;
                                    paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                                    worksheetn.Cells[addrow + 3, 13] = signature;
                                    afterSignatureRow = addrow + 1 + signatureRow;
                                }
                                else
                                {
                                    // 因簽名列塞不小，增加分頁
                                    alladdSheet++;
                                    worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];
                                    worksheetSignature.Copy(worksheetn);
                                    worksheetn.Cells[3, 13] = signature;
                                    afterSignatureRow = signatureRow;
                                }
                            }

                            #region 加入 4*10 的框
                            worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];

                            // 共要加入幾組 frameNum
                            int frameNum = (int)Math.Ceiling(dr.Length * 1.0 / 2);

                            // 有簽名列那頁下方還有空間放下
                            if (afterSignatureRow <= setRow - frameRow)
                            {
                                Excel.Range paste2 = worksheetn.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                                Excel.Range r2 = worksheetFrame.get_Range("A9:A42").EntireRow;
                                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                                frameNum--;
                                afterSignatureRow += frameRow;

                                if (afterSignatureRow <= setRow - frameRow)
                                {
                                    paste2 = worksheetn.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                                    paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                                    frameNum--;
                                }
                            }

                            bool g1 = true;
                            for (int f = 0; f < frameNum; f++)
                            {
                                // 此頁第一組
                                if (g1)
                                {
                                    alladdSheet++;
                                    worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];
                                    worksheetFrame.Copy(worksheetn);
                                }

                                // 此頁第2組
                                else
                                {
                                    worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i + k + 1];
                                    Excel.Range paste2 = worksheetn.get_Range($"A43", Type.Missing);
                                    Excel.Range r2 = worksheetFrame.get_Range("A9:A42").EntireRow;
                                    paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                                }

                                g1 = !g1;
                            }

                            Excel.Range cell;
                            if (frameNum == 0)
                            {
                                worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                Excel.Range paste2 = worksheetn.get_Range($"A52", Type.Missing);
                                Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, r2.Copy(Type.Missing));

                                if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                                {
                                    cell = worksheetn.Cells[54, 2];
                                    worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 280);
                                }
                                if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                                {
                                    cell = worksheetn.Cells[54, 10];
                                    worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 280);
                                }
                            }
                            else if (!g1)
                            {
                                worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                Excel.Range paste2 = worksheetn.get_Range($"A46", Type.Missing);
                                Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, r2.Copy(Type.Missing));

                                if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                                {
                                    cell = worksheetn.Cells[48, 2];
                                    worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 285);
                                }
                                if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                                {
                                    cell = worksheetn.Cells[48, 10];
                                    worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 285);
                                }
                            }
                            else
                            {
                                alladdSheet++;
                                Excel.Worksheet worksheepic = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                worksheetPicture.Copy(worksheepic);

                                worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                                {
                                    cell = worksheetn.Cells[3, 2];
                                    worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 430, 285);
                                }
                                if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                                {
                                    cell = worksheetn.Cells[3, 10];
                                    worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 430, 285);
                                }
                            }
                            #endregion
                        }

                        alladdSheet += addSheets;
                    }
                    else
                    {
                        // 有表頭那頁下方的細項格線
                        worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                        for (int j = 0; j < dr.Length; j++)
                        {
                            Excel.Range paste1 = worksheet.get_Range($"A{headerRow + 1 + j}", Type.Missing);
                            Excel.Range r = worksheetDetail.get_Range("A1").EntireRow;
                            paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r.Copy(Type.Missing));

                            // 細項資料
                            this.SetDetailData(worksheet, j + headerRow + 1, dr[j]);
                        }
                        worksheet.Cells.EntireRow.AutoFit();

                        int afterSignatureRow;

                        // 簽名列
                        if (dr.Length < underHeaderRow - signatureRow)
                        {
                            Excel.Range paste2 = worksheet.get_Range($"A{dr.Length + headerRow + 1}", Type.Missing);
                            Excel.Range r2 = worksheetSignature.get_Range("A1:A4").EntireRow;
                            paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                            worksheet.Cells[dr.Length + headerRow + 3, 13] = signature;
                            afterSignatureRow = dr.Length + headerRow + signatureRow;
                        }
                        else
                        {
                            // 因簽名列塞不小，增加分頁
                            alladdSheet++;
                            Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                            worksheetSignature.Copy(worksheetn);
                            worksheetn.Cells[3, 13] = signature;
                            afterSignatureRow = signatureRow;
                        }

                        #region 加入 4*10 的框
                        worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];

                        // 共要加入幾組 frameNum
                        int frameNum = (int)Math.Ceiling(dr.Length * 1.0 / 2);

                        // 有簽名列那頁下方還有空間放下
                        if (afterSignatureRow <= setRow - frameRow)
                        {
                            Excel.Range paste2 = worksheet.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                            Excel.Range r2 = worksheetFrame.get_Range("A9:A42").EntireRow;
                            paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                            frameNum--;
                            afterSignatureRow += frameRow;

                            if (afterSignatureRow <= setRow - frameRow)
                            {
                                paste2 = worksheet.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                                frameNum--;
                            }
                        }

                        bool g1 = true;
                        for (int f = 0; f < frameNum; f++)
                        {
                            // 此頁第一組
                            if (g1)
                            {
                                alladdSheet++;
                                Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                worksheetFrame.Copy(worksheetn);
                            }

                            // 此頁第2組
                            else
                            {
                                Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                                Excel.Range paste2 = worksheetn.get_Range($"A43", Type.Missing);
                                Excel.Range r2 = worksheetFrame.get_Range("A9:A42").EntireRow;
                                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));
                            }

                            g1 = !g1;
                        }
                        #endregion

                        Excel.Range cell;
                        if (frameNum == 0)
                        {
                            Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                            Excel.Range paste2 = worksheetn.get_Range($"A52", Type.Missing);
                            Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                            paste2.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, r2.Copy(Type.Missing));

                            if(!string.IsNullOrEmpty(imgPath_BeforePicture))
                            {
                                cell = worksheetn.Cells[54, 2];
                                worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 280);
                            }
                            if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                            {
                                cell = worksheetn.Cells[54, 10];
                                worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 280);
                            }
                        }
                        else if (!g1) {
                            Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                            Excel.Range paste2 = worksheetn.get_Range($"A46", Type.Missing);
                            Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                            paste2.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, r2.Copy(Type.Missing));

                            if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                            {
                                cell = worksheetn.Cells[48, 2];
                                worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 285);
                            }
                            if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                            {
                                cell = worksheetn.Cells[48, 10];
                                worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 435, 285);
                            }
                        }
                        else
                        {
                            alladdSheet++;
                            Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                            worksheetPicture.Copy(worksheetn);

                            worksheetn = excel.ActiveWorkbook.Worksheets[defaultSheet + alladdSheet + i];
                            if (!string.IsNullOrEmpty(imgPath_BeforePicture))
                            {
                                cell = worksheetn.Cells[3, 2];
                                worksheetn.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 430, 285);
                            }
                            if (!string.IsNullOrEmpty(imgPath_AfterPicture))
                            {
                                cell = worksheetn.Cells[3, 10];
                                worksheetn.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 430, 285);
                            }
                        }
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    worksheet = excel.ActiveWorkbook.Worksheets[1];
                    worksheet.Delete();
                }

                #region Save & Show Excel

                pdfFileName = $"FabricOvenTestDetailReportToPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                string excelFileName = $"FabricOvenTestDetailReportToPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                if (!isCreatePdfOK)
                {
                    result.Result = false;
                    result.ErrorMessage = "ConvertToPDF fail";
                    return result;
                }

                #endregion
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            

            return result;
        }

        public BaseResult ToPdfFabricOvenTestDetail(string poID, string TestNo, out string pdfFileName, bool isTest)
        {
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);

            BaseResult result = new BaseResult();
            pdfFileName = string.Empty;

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
                string imgPath_BeforePicture = string.Empty;
                string imgPath_AfterPicture = string.Empty;
                if (dtOven.Rows[0]["TestBeforePicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_BeforePicture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_BeforePicture = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestBeforePicture"];
                    using (var imageFile = new FileStream(imgPath_BeforePicture, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                }

                //Excel.Range cellAfter = worksheet.Cells[3, 10];
                if (dtOven.Rows[0]["TestAfterPicture"] != DBNull.Value)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    if (IsTest.ToLower() == "true")
                    {
                        imgPath_AfterPicture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath_AfterPicture = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    byte[] bytes = (byte[])dtOven.Rows[0]["TestAfterPicture"];
                    using (var imageFile = new FileStream(imgPath_AfterPicture, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }

                    //worksheet.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left + 5, cellAfter.Top + 5, 430, 295);
                }
                #endregion

                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];
                // 填入表頭資訊
                worksheet = excel.ActiveWorkbook.Worksheets[defaultSheet];
                worksheet.Cells[4, 3] = distOvenDetailSubmitDate[0]["ReportNo"].ToString();
                worksheet.Cells[4, 6] = dtOven.Rows[0]["InspDate"] == DBNull.Value ? string.Empty : ((DateTime)dtOven.Rows[0]["InspDate"]).ToString("yyyy/MM/dd");
                worksheet.Cells[4, 9] = poID;
                worksheet.Cells[4, 14] = brandID;
                worksheet.Cells[6, 3] = styleID;
                worksheet.Cells[6, 9] = CustPONO;
                worksheet.Cells[6, 14] = dtOven.Rows[0]["Article"].ToString();
                worksheet.Cells[7, 3] = listStyle[0].StyleName;
                worksheet.Cells[7, 9] = seasonID;

                // 細項
                int headerRow = 9; // 表頭那頁前9列為固定
                int signatureRow = 4; // 簽名有4列

                string signature = dtOven.Rows[0]["InspectorName"].ToString();
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
                worksheet.Cells.EntireRow.AutoFit();

                // 簽名格子塞入後的Row Index
                int afterSignatureRow;

                Excel.Range paste1 = worksheet.get_Range($"A{dr.Length + headerRow + 1}", Type.Missing);
                Excel.Range r1 = worksheetSignature.get_Range("A1:A4").EntireRow;
                paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r1.Copy(Type.Missing));
                worksheet.Cells[dr.Length + headerRow + 3, 13] = signature;
                afterSignatureRow = dr.Length + headerRow + signatureRow;


                Excel.Range paste2 = worksheet.get_Range($"A{afterSignatureRow + 1}", Type.Missing);
                Excel.Range r2 = worksheetPicture.get_Range("A1:A20").EntireRow;
                paste2.Insert(Excel.XlInsertShiftDirection.xlShiftDown, r2.Copy(Type.Missing));

                Excel.Range cell;
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

                for (int i = 0; i < 4; i++)
                {
                    worksheet = excel.ActiveWorkbook.Worksheets[1];
                    worksheet.Delete();
                }

                #region Save & Show Excel

                pdfFileName = $"FabricOvenTestDetailReportToPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                string excelFileName = $"FabricOvenTestDetailReportToPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                if (!isCreatePdfOK)
                {
                    result.Result = false;
                    result.ErrorMessage = "ConvertToPDF fail";
                    return result;
                }

                #endregion
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }



            return result;
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
