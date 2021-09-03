using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class FabricCrkShrkTest_Service : IFabricCrkShrkTest_Service
    {
        IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider;
        IScaleProvider _ScaleProvider;
        IOrdersProvider _OrdersProvider;
        IStyleProvider _StyleProvider;

        public BaseResult AmendFabricCrkShrkTestCrockingDetail(long ID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);

                if (fabricCrkShrkTestCrocking_Main.CrockingEncdoe == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricCrocking(ID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult AmendFabricCrkShrkTestHeatDetail(long ID)
        {
            throw new NotImplementedException();
        }

        public BaseResult AmendFabricCrkShrkTestWashDetail(long ID)
        {
            throw new NotImplementedException();
        }

        public BaseResult EncodeFabricCrkShrkTestCrockingDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = this.GetFabricCrkShrkTestCrocking_Result(ID);

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingEncdoe == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string crockingResult = "Pass";

                if (fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    crockingResult = "Fail";
                }

                testResult = crockingResult;

                DateTime? crockingDate = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricCrocking(ID, testResult, crockingDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
                return baseResult;
            }
        }

        public BaseResult EncodeFabricCrkShrkTestHeatDetail(long ID, out string ovenTestResult)
        {
            throw new NotImplementedException();
        }

        public BaseResult EncodeFabricCrkShrkTestWashDetail(long ID, out string ovenTestResult)
        {
            throw new NotImplementedException();
        }

        public FabricCrkShrkTestCrocking_Result GetFabricCrkShrkTestCrocking_Result(long ID)
        {
            try
            {
                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = new FabricCrkShrkTestCrocking_Result();
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestCrocking_Result.Crocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);

                fabricCrkShrkTestCrocking_Result.Crocking_Detail = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Detail(ID);

                fabricCrkShrkTestCrocking_Result.ID = ID;

                fabricCrkShrkTestCrocking_Result.CrockingTestOption = _FabricCrkShrkTestProvider.GetCrockingTestOption(ID);

                fabricCrkShrkTestCrocking_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();

                return fabricCrkShrkTestCrocking_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public FabricCrkShrkTestHeat_Result GetFabricCrkShrkTestHeat_Result(long ID)
        {
            throw new NotImplementedException();
        }

        public FabricCrkShrkTestWash_Result GetFabricCrkShrkTestWash_Result(long ID)
        {
            throw new NotImplementedException();
        }

        public FabricCrkShrkTest_Result GetFabricCrkShrkTest_Result(string POID)
        {
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                return _FabricCrkShrkTestProvider.GetFabricCrkShrkTest_Main(POID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult SaveFabricCrkShrkTestCrockingDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                bool isRollDyelotEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot));
                if (isRollDyelotEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Roll and Dyelot cannot be empty.";
                    return baseResult;
                }

                bool isScaleResultEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s =>
                                                    MyUtility.Check.Empty(s.DryScale) ||
                                                    MyUtility.Check.Empty(s.ResultDry) ||
                                                    MyUtility.Check.Empty(s.WetScale) ||
                                                    MyUtility.Check.Empty(s.ResultWet) ||
                                                    (fabricCrkShrkTestCrocking_Result.CrockingTestOption == 1 && 
                                                        (MyUtility.Check.Empty(s.DryScale_Weft) ||
                                                         MyUtility.Check.Empty(s.ResultDry_Weft) ||
                                                         MyUtility.Check.Empty(s.WetScale_Weft) ||
                                                         MyUtility.Check.Empty(s.ResultWet_Weft)))
                                                    );

                if (isScaleResultEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Scale and Result cannot be empty.";
                    return baseResult;
                }

                bool isInspectorEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => MyUtility.Check.Empty(s.Inspector));
                if (isInspectorEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Lab Tech cannot be empty.";
                    return baseResult;
                }

                //再檢查一次Result
                foreach (FabricCrkShrkTestCrocking_Detail fabricCrkShrkTestCrocking_Detail in fabricCrkShrkTestCrocking_Result.Crocking_Detail)
                {
                    if (fabricCrkShrkTestCrocking_Detail.ResultDry.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultDry_Weft.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultWet.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultWet_Weft.ToUpper() == "FAIL"
                        )
                    {
                        fabricCrkShrkTestCrocking_Detail.Result = "Fail";
                    }
                    else
                    {
                        fabricCrkShrkTestCrocking_Detail.Result = "Pass";
                    }
                }

                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture = new byte[0];
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricCrockingTestDetail(fabricCrkShrkTestCrocking_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestHeatDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID)
        {
            throw new NotImplementedException();
        }

        public BaseResult SaveFabricCrkShrkTestMain(FabricCrkShrkTest_Result fabricCrkShrkTest_Result)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                _FabricCrkShrkTestProvider.SaveFabricCrkShrkTest_Main(fabricCrkShrkTest_Result);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestWashDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID)
        {
            throw new NotImplementedException();
        }

        public SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetCrockingFailMailContentData(ID);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Crocking Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request, isTest);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.ToString();
            }

            return result;
        }

        public BaseResult ToExcelFabricCrkShrkTestCrockingDetail(long ID, out string excelFileName, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtCrockingDetail = _FabricCrkShrkTestProvider.GetCrockingDetailForReport(ID);
                FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);
                string[] columnNames;
                string excelName = string.Empty;

                int crockingTestOption = _FabricCrkShrkTestProvider.GetCrockingTestOption(ID);

                switch (crockingTestOption)
                {
                    case 0:
                        columnNames = new string[] { "Roll", "Dyelot", "DryScale", "WetScale", "Result", "InspDate", "Inspector", "Remark", "LastUpdate" };
                        excelName = baseFilePath + "\\XLT\\FabricCrockingTest.xltx";
                        break;
                    default:
                        columnNames = new string[] { "Roll", "Dyelot", "DryScale", "DryScale_Weft", "WetScale", "WetScale_Weft", "Result", "InspDate", "Inspector", "Remark", "LastUpdate" };
                        excelName = baseFilePath + "\\XLT\\FabricCrockingTestWeftWarp.xltx";
                        break;
                }

                var ret = Array.CreateInstance(typeof(object), dtCrockingDetail.Rows.Count, columnNames.Length) as object[,];
                for (int i = 0; i < dtCrockingDetail.Rows.Count; i++)
                {
                    DataRow row = dtCrockingDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

                if (dtCrockingDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                // 撈取seasonID
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestCrocking_Main.POID}).ToList();

                string seasonID;

                if (listOrders.Count == 0)
                {
                    seasonID = string.Empty;
                }
                else
                {
                    seasonID = listOrders[0].SeasonID;
                }

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Workbooks.Add();
                MyUtility.Excel.CopyToXls(ret, xltFileName: excelName, fileName: string.Empty, openfile: false, headerline: 5, excelAppObj: excel);
                Microsoft.Office.Interop.Excel.Worksheet excelSheets = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表
                excel.Cells[2, 2] = fabricCrkShrkTestCrocking_Main.POID;
                excel.Cells[2, 4] = fabricCrkShrkTestCrocking_Main.SEQ;
                excel.Cells[2, 6] = fabricCrkShrkTestCrocking_Main.ColorID;
                excel.Cells[2, 8] = fabricCrkShrkTestCrocking_Main.StyleID;
                excel.Cells[2, 10] = seasonID;
                excel.Cells[3, 2] = fabricCrkShrkTestCrocking_Main.SCIRefno;
                excel.Cells[3, 4] = fabricCrkShrkTestCrocking_Main.ExportID;
                excel.Cells[3, 6] = fabricCrkShrkTestCrocking_Main.Crocking;
                excel.Cells[3, 8] = fabricCrkShrkTestCrocking_Main.CrockingDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestCrocking_Main.CrockingDate).ToString("yyyy/MM/dd");
                excel.Cells[3, 10] = fabricCrkShrkTestCrocking_Main.BrandID;
                excel.Cells[4, 2] = fabricCrkShrkTestCrocking_Main.Refno;
                excel.Cells[4, 4] = fabricCrkShrkTestCrocking_Main.ArriveQty;
                excel.Cells[4, 6] = fabricCrkShrkTestCrocking_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestCrocking_Main.WhseArrival).ToString("yyyy/MM/dd");
                excel.Cells[4, 8] = fabricCrkShrkTestCrocking_Main.Supp;
                excel.Cells[4, 10] = fabricCrkShrkTestCrocking_Main.NonCrocking.ToString();

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();       ////自動欄高

                #region Save & Show Excel
                excelFileName = $"FabricCrockingTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }

        public BaseResult ToExcelFabricCrkShrkTestHeatDetail(long ID, out string excelFileName, bool isTest)
        {
            throw new NotImplementedException();
        }

        public BaseResult ToExcelFabricCrkShrkTestWashDetail(long ID, out string excelFileName, bool isTest)
        {
            throw new NotImplementedException();
        }

        public BaseResult ToPdfFabricCrkShrkTestCrockingDetail(long ID, out string pdfFileName, bool isTest)
        {
            throw new NotImplementedException();
        }

        public BaseResult ToPdfFabricCrkShrkTestHeatDetail(long ID, out string pdfFileName, bool isTest)
        {
            throw new NotImplementedException();
        }

        public BaseResult ToPdfFabricCrkShrkTestWashDetail(long ID, out string pdfFileName, bool isTest)
        {
            throw new NotImplementedException();
        }
    }
}
