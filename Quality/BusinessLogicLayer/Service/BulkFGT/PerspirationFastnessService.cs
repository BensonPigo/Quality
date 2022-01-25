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
    public class PerspirationFastnessService
    {
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];
        public PerspirationFastnessProvider _PerspirationFastnessProvider;
        public IScaleProvider _ScaleProvider;
        public IOrdersProvider _OrdersProvider;
        public IStyleProvider _StyleProvider;

        public BaseResult AmendPerspirationFastnessDetail(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = _PerspirationFastnessProvider.GetPerspirationFastness_Detail(poID, TestNo);

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
                baseResult.ErrorMessage = ex.Message.ToString();
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
                baseResult.ErrorMessage = ex.Message.ToString();
                return baseResult;
            }
        }

        public PerspirationFastness_Detail_Result GetPerspirationFastness_Detail_Result(string poID, string TestNo)
        {
            try
            {
                PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = new PerspirationFastness_Detail_Result();
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                PerspirationFastness_Detail_Result = _PerspirationFastnessProvider.GetPerspirationFastness_Detail(poID, TestNo);

                PerspirationFastness_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();

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
                result.ErrorMessage = ex.Message.ToString();
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
                baseResult.ErrorMessage = ex.Message.ToString();
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
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public SendMail_Result SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _PerspirationFastnessProvider.GetFailMailContentData(poID, TestNo);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Oven Test - Test Fail",
                    Body = mailBody
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

        private void SetDetailData(Excel.Worksheet worksheet, int setRow, DataRow dr)
        {
            worksheet.Cells[setRow, 2] = dr["Refno"];
            worksheet.Cells[setRow, 3] = dr["Colorid"];
            worksheet.Cells[setRow, 4] = dr["Dyelot"];
            worksheet.Cells[setRow, 6] = dr["Roll"];

            worksheet.Cells[setRow, 7] = dr["AlkalineChangescale"];
            worksheet.Cells[setRow, 8] = dr["AlkalineResultChange"];
            worksheet.Cells[setRow, 9] = dr["AlkalineAcetatescale"];
            worksheet.Cells[setRow, 10] = dr["AlkalineResultAcetate"];
            worksheet.Cells[setRow, 11] = dr["AlkalineCottonscale"];
            worksheet.Cells[setRow, 12] = dr["AlkalineResultCotton"];
            worksheet.Cells[setRow, 13] = dr["AlkalineNylonscale"];
            worksheet.Cells[setRow, 14] = dr["AlkalineResultNylon"];
            worksheet.Cells[setRow, 15] = dr["AlkalinePolyesterscale"];
            worksheet.Cells[setRow, 16] = dr["AlkalineResultPolyester"];
            worksheet.Cells[setRow, 17] = dr["AlkalineAcrylicscale"];
            worksheet.Cells[setRow, 18] = dr["AlkalineResultAcrylic"];
            worksheet.Cells[setRow, 19] = dr["AlkalineWoolscale"];
            worksheet.Cells[setRow, 20] = dr["AlkalineResultWool"];

            worksheet.Cells[setRow, 7] = dr["AcidChangescale"];
            worksheet.Cells[setRow, 8] = dr["AcidResultChange"];
            worksheet.Cells[setRow, 9] = dr["AcidAcetatescale"];
            worksheet.Cells[setRow, 10] = dr["AcidResultAcetate"];
            worksheet.Cells[setRow, 11] = dr["AcidCottonscale"];
            worksheet.Cells[setRow, 12] = dr["AcidResultCotton"];
            worksheet.Cells[setRow, 13] = dr["AcidNylonscale"];
            worksheet.Cells[setRow, 14] = dr["AcidResultNylon"];
            worksheet.Cells[setRow, 15] = dr["AcidPolyesterscale"];
            worksheet.Cells[setRow, 16] = dr["AcidResultPolyester"];
            worksheet.Cells[setRow, 17] = dr["AcidAcrylicscale"];
            worksheet.Cells[setRow, 18] = dr["AcidResultAcrylic"];
            worksheet.Cells[setRow, 19] = dr["AcidWoolscale"];
            worksheet.Cells[setRow, 20] = dr["AcidResultWool"];


            worksheet.Cells[setRow, 21] = MyUtility.Convert.GetString(dr["Temperature"]) + "˚C";
            worksheet.Cells[setRow, 22] = MyUtility.Convert.GetString(dr["Time"]) + " hrs";
            worksheet.Cells[setRow, 23] = dr["Remark"];
        }

        public BaseResult ToReport(string ID, out string FileName, bool isPDF, bool isTest = false)
        {
            BaseResult result = new BaseResult();
            _PerspirationFastnessProvider = new PerspirationFastnessProvider(Common.ProductionDataAccessLayer);
            List<PerspirationFastness_Excel> dataList = new List<PerspirationFastness_Excel>();

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

                    currenSheet.Cells[2, 3] = currenData.SubmitDate.HasValue ? currenData.SubmitDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                    currenSheet.Cells[2, 8] = DateTime.Now.ToString("yyyy/MM/dd");

                    currenSheet.Cells[3, 3] = currenData.SeasonID;
                    currenSheet.Cells[3, 8] = currenData.BrandID;

                    currenSheet.Cells[4, 3] = currenData.StyleID;
                    currenSheet.Cells[4, 8] = currenData.POID;

                    currenSheet.Cells[5, 3] = currenData.Roll;
                    currenSheet.Cells[5, 8] = currenData.Dyelot;

                    currenSheet.Cells[6, 3] = currenData.SCIRefno_Color;

                    // Test Request
                    currenSheet.Cells[8, 3] = currenData.Temperature;
                    currenSheet.Cells[8, 8] = currenData.Time;

                    currenSheet.Cells[12, 2] = currenData.AlkalineChangeScale;
                    currenSheet.Cells[12, 3] = currenData.AlkalineAcetateScale;
                    currenSheet.Cells[12, 4] = currenData.AlkalineCottonScale;
                    currenSheet.Cells[12, 5] = currenData.AlkalineNylonScale;
                    currenSheet.Cells[12, 6] = currenData.AlkalinePolyesterScale;
                    currenSheet.Cells[12, 7] = currenData.AlkalineAcrylicScale;
                    currenSheet.Cells[12, 8] = currenData.AlkalineWoolScale;

                    currenSheet.Cells[13, 2] = currenData.AlkalineResultChange;
                    currenSheet.Cells[13, 3] = currenData.AlkalineResultAcetate;
                    currenSheet.Cells[13, 4] = currenData.AlkalineResultCotton;
                    currenSheet.Cells[13, 5] = currenData.AlkalineResultNylon;
                    currenSheet.Cells[13, 6] = currenData.AlkalineResultPolyester;
                    currenSheet.Cells[13, 7] = currenData.AlkalineResultAcrylic;
                    currenSheet.Cells[13, 8] = currenData.AlkalineResultWool;

                    currenSheet.Cells[17, 2] = currenData.AcidChangeScale;
                    currenSheet.Cells[17, 3] = currenData.AcidAcetateScale;
                    currenSheet.Cells[17, 4] = currenData.AcidCottonScale;
                    currenSheet.Cells[17, 5] = currenData.AcidNylonScale;
                    currenSheet.Cells[17, 6] = currenData.AcidPolyesterScale;
                    currenSheet.Cells[17, 7] = currenData.AcidAcrylicScale;
                    currenSheet.Cells[17, 8] = currenData.AcidWoolScale;

                    currenSheet.Cells[18, 2] = currenData.AcidResultChange;
                    currenSheet.Cells[18, 3] = currenData.AcidResultAcetate;
                    currenSheet.Cells[18, 4] = currenData.AcidResultCotton;
                    currenSheet.Cells[18, 5] = currenData.AcidResultNylon;
                    currenSheet.Cells[18, 6] = currenData.AcidResultPolyester;
                    currenSheet.Cells[18, 7] = currenData.AcidResultAcrylic;
                    currenSheet.Cells[18, 8] = currenData.AcidResultWool;

                    currenSheet.Cells[19, 2] = currenData.Remark;
                    currenSheet.Cells[75, 3] = currenData.Inspector;
                    currenSheet.Cells[75, 7] = currenData.Inspector;


                    #region 添加圖片
                    Excel.Range cellBeforePicture = currenSheet.Cells[50, 1];
                    if (currenData.TestBeforePicture != null)
                    {
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath;

                        if (isTest)
                        {
                            imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                        }
                        else
                        {
                            imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                        }

                        byte[] bytes = currenData.TestBeforePicture;
                        using (var imageFile = new FileStream(imgPath, FileMode.Create))
                        {
                            imageFile.Write(bytes, 0, bytes.Length);
                            imageFile.Flush();
                        }
                        currenSheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 380, 300);
                    }

                    Excel.Range cellAfterPicture = currenSheet.Cells[50, 5];
                    if (currenData.TestAfterPicture != null)
                    {
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath;

                        if (isTest)
                        {
                            imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                        }
                        else
                        {
                            imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                        }

                        byte[] bytes = currenData.TestAfterPicture;
                        using (var imageFile = new FileStream(imgPath, FileMode.Create))
                        {
                            imageFile.Write(bytes, 0, bytes.Length);
                            imageFile.Flush();
                        }
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
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }
    }
}
