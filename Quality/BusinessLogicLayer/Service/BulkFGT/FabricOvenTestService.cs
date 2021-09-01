using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{
    public class FabricOvenTestService : IFabricOvenTestService
    {
        IFabricOvenTestProvider _FabricOvenTestProvider;
        IScaleProvider _ScaleProvider;

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
                baseResult.ErrorMessage = ex.ToString();
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
                baseResult.ErrorMessage = ex.ToString();
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
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);

                return _FabricOvenTestProvider.GetFabricOvenTest_Main(POID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
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

                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);

                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.TestNo))
                {
                    _FabricOvenTestProvider.AddFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID);
                }
                else
                {
                    _FabricOvenTestProvider.EditFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID);
                }
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
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
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricOvenTestProvider.GetFailMailContentData(poID, TestNo);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                //string resultMsg = MailTools.MailToHtml(toAddress, );

            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public string ToExcelFabricOvenTestDetail(string poID, string TestNo)
        {
            DataTable dt = (DataTable)this.gridbs.DataSource;
            string[] columnNames = new string[] { "OvenGroup", "SEQ", "Roll", "Dyelot", "SCIRefno", "Colorid", "Supplier", "Changescale", "StainingScale", "Result", "Remark" };
            var ret = Array.CreateInstance(typeof(object), dt.Rows.Count, this.grid.Columns.Count) as object[,];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                for (int j = 0; j < columnNames.Length; j++)
                {
                    ret[i, j] = row[columnNames[j]];
                }
            }

            if (dt.Rows.Count == 0)
            {
                MyUtility.Msg.WarningBox("Data not found!");
                return;
            }

            string styleID;
            string seasonID;
            string status;
            string brandID;
            DBProxy.Current.Select(null, string.Format("select * from PO WITH (NOLOCK) where id='{0}'", this.PoID), out DataTable dtPo);
            if (dtPo.Rows.Count == 0)
            {
                styleID = string.Empty;
                seasonID = string.Empty;
                status = string.Empty;
                brandID = string.Empty;
            }
            else
            {
                styleID = dtPo.Rows[0]["StyleID"].ToString();
                seasonID = dtPo.Rows[0]["SeasonID"].ToString();
                status = this.dtOven.Rows[0]["status"].ToString();
                brandID = dtPo.Rows[0]["BrandID"].ToString();
            }

            string strXltName = Env.Cfg.XltPathDir + "\\Quality_P05_Detail_Report.xltx";
            Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
            if (excel == null)
            {
                return;
            }

            Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];

            worksheet.Cells[1, 2] = this.txtSP.Text.ToString();
            worksheet.Cells[1, 4] = styleID;
            worksheet.Cells[1, 6] = dtPo.Rows[0]["SeasonID"].ToString();
            worksheet.Cells[1, 8] = seasonID;
            worksheet.Cells[1, 10] = this.txtNoofTest.Text.ToString();
            worksheet.Cells[2, 2] = status;
            worksheet.Cells[2, 4] = this.comboResult.Text;
            worksheet.Cells[2, 6] = this.dateTestDate.Text;
            worksheet.Cells[2, 8] = this.txtuserInspector.TextBox1.Text.ToString();
            worksheet.Cells[2, 10] = brandID;

            int startRow = 4;
            for (int i = 0; i < dt.Rows.Count; i++)
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
            MyUtility.Msg.WaitClear();

            #region Save & Show Excel
            string strExcelName = Class.MicrosoftFile.GetName("Quality_P05_Detail_Report");
            excel.ActiveWorkbook.SaveAs(strExcelName);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            Marshal.ReleaseComObject(worksheet);

            strExcelName.OpenFile();
        }

        public string ToPdfFabricOvenTestDetail(string poID, string TestNo)
        {
            throw new NotImplementedException();
        }
    }
}
