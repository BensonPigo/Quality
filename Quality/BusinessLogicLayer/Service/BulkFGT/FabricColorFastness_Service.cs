using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using MICS.DataAccessLayer.Interface;
using MICS.DataAccessLayer.Provider.MSSQL;
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
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class FabricColorFastness_Service : IFabricColorFastness_Service
    {
        private IColorFastnessProvider _IColorFastnessProvider;
        private IColorFastnessDetailProvider _IColorFastnessDetailProvider;
        private IOrdersProvider _IOrdersProvider;

        public enum DetailStatus
        {
            Encode,
            Amend,
        }

        public List<string> Get_Scales()
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessProvider.GetScales();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public FabricColorFastness_ViewModel Get_Main(string PoID)
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            FabricColorFastness_ViewModel result = new FabricColorFastness_ViewModel();
            result.Result = true;
            try
            {
                if (string.IsNullOrEmpty(PoID))
                {
                    result.Result = false;
                    result.ErrorMessage = "PoID cannot be empty!";
                    return result;
                }

                var source = _IColorFastnessProvider.GetMain(PoID);
                if (!source.Result)
                {
                    result.Result = source.Result;
                    result.ErrorMessage = source.ErrorMessage;
                    return result;
                }

                if (string.IsNullOrEmpty(source.PoID))
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                result = source;
                result.TargetLeadTime = _IColorFastnessProvider.Get_Target_LeadTime(result.CutInLine, result.MinSciDelivery);
                result.CompletionDate = (source.ArticlePercent >= 100) ? source.CompletionDate : null;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public ColorFastness_Result GetDetailHeader(string ID)
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            ColorFastness_Result result = new ColorFastness_Result();
            try
            {
                if (string.IsNullOrEmpty(ID))
                {
                    result.baseResult.Result = false;
                    result.baseResult.ErrorMessage = "ID cannot be empty!";
                    return result;
                }

                var list = _IColorFastnessProvider.Get(ID);
                if (!list.Any() || list.Count() == 0)
                {
                    result.baseResult.Result = false;
                    result.baseResult.ErrorMessage = "Data not found!";
                    return result;
                }

                result = list.FirstOrDefault();
                result.baseResult = new DatabaseObject.BaseResult();
                result.baseResult.Result = true;
                result.baseResult.ErrorMessage = "";
            }
            catch (Exception ex)
            {
                result.baseResult.Result = false;
                result.baseResult.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel GetDetailBody(string ID)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            try
            {
                result = _IColorFastnessDetailProvider.Get_DetailBody(ID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<PO_Supp_Detail> GetSeq(string POID, string Seq1 = "", string Seq2 = "")
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessDetailProvider.Get_Seq(POID, Seq1, Seq2);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<FtyInventory> GetRoll(string POID, string Seq1, string Seq2)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessDetailProvider.Get_Roll(POID, Seq1, Seq2);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult Save_ColorFastness_1stPage(string PoID, string Remark)
        {
            BaseResult baseResult = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessProvider = new ColorFastnessProvider(_ISQLDataTransaction);
                baseResult.Result = _IColorFastnessProvider.Save_PO(PoID, Remark);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }

        public BaseResult Save_ColorFastness_2ndPage(Fabric_ColorFastness_Detail_ViewModel source , string Mdivision, string UserID)
        {
            BaseResult baseResult = new BaseResult();
            baseResult.Result = true;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(_ISQLDataTransaction);

                // 若表身資料重複 跳訊息
                

                foreach (var item in source.Detail)
                {
                    int repeatCnt = source.Detail.Where(s => s.ID == item.ID && s.ColorFastnessGroup == item.ColorFastnessGroup && s.Seq == item.Seq).Count();
                    if (repeatCnt > 1)
                    {
                        baseResult.ErrorMessage = $@"＜Body: {item.ColorFastnessGroup}＞, ＜Seq1: {item.Seq}＞ is repeat cannot save.";
                        baseResult.Result = false;
                        return baseResult;
                    }

                    if (string.IsNullOrEmpty(item.Seq))
                    {
                        baseResult.ErrorMessage = $"＜SEQ#＞ cannot be empty.";
                        baseResult.Result = false;
                        return baseResult;
                    }
                }


                _IColorFastnessDetailProvider.Save_ColorFastness(source, Mdivision, UserID, out string NewID);
                _ISQLDataTransaction.Commit();
                if (!string.IsNullOrEmpty(NewID))
                {
                    baseResult.ErrorMessage = NewID;
                }

                baseResult.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }

        public Fabric_ColorFastness_Detail_ViewModel Encode_ColorFastness(string ID, DetailStatus status, string UserID)
        {
            Fabric_ColorFastness_Detail_ViewModel result = GetDetailBody(ID);
            result.sentMail = false;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessProvider = new ColorFastnessProvider(_ISQLDataTransaction);
                // Check Detail Result
                bool detailResult = true;
                
                foreach (var item in result.Detail)
                {
                    if (item.Result.ToUpper() == "FAIL")
                    {
                        detailResult = false;
                    }
                }

                string strResult = detailResult ? "Pass" : "Fail";

                switch (status)
                {
                    case DetailStatus.Encode:
                        result.sentMail = !detailResult;                       
                        result.Result = _IColorFastnessProvider.Encode_ColorFastness(ID, "Confirmed", strResult, UserID);
                        break;
                    case DetailStatus.Amend:
                        result.Result = _IColorFastnessProvider.Encode_ColorFastness(ID, "New", strResult, UserID);
                        break;
                    default:
                        break;
                }

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public BaseResult SentMail(string POID, string ID, string ToAddress, string CCAddress)
        {
            BaseResult result = new BaseResult();
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                if (string.IsNullOrEmpty(ToAddress) == true)
                {
                    result.Result = false;
                    result.ErrorMessage = "ToEmail address is empty!";
                    return result;
                }

                DataTable dtContent = _IColorFastnessProvider.Get_Mail_Content(POID, ID);
                string strHtml = MailTools.DataTableChangeHtml(dtContent);

                SendMail_Request request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CCAddress,
                    Subject = "Washing Fastness-Crocking Test – Test Fail",
                    Body = strHtml,
                };

                MailTools.SendMail(request);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel ToExcel()
        {
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel ToPDF(string ID, bool test = false )
        {
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

            result = GetDetailBody(ID);

            if (string.IsNullOrEmpty(result.Main.ID))
            {
                result.Result = false;
                result.ErrorMessage = "ID cannot be empty!";
                return result;
            }

            if (result.Detail.Count == 0)
            {
                result.Result = false;
                result.ErrorMessage = "Detail data not found!";
                return result;
            }

            DataTable dtSubDate = _IColorFastnessDetailProvider.Get_SubmitDate(result.Main.ID);
            if (dtSubDate.Rows.Count < 1)
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }
           
            DataTable dtOrders = _IOrdersProvider.Get_Orders_DataTable("", result.Main.POID);
            string styleID, seasonID, CustPONO, brandID, styleUkey;

            if (dtOrders.Rows.Count == 0)
            {
                styleID = string.Empty;
                seasonID = string.Empty;
                CustPONO = string.Empty;
                brandID = string.Empty;
                styleUkey = string.Empty;
            }
            else
            {
                styleID = dtOrders.Rows[0]["StyleID"].ToString();
                styleUkey = dtOrders.Rows[0]["StyleUkey"].ToString();
                seasonID = dtOrders.Rows[0]["SeasonID"].ToString();
                CustPONO = dtOrders.Rows[0]["CustPONO"].ToString();
                brandID = dtOrders.Rows[0]["BrandID"].ToString();
            }

            IStyleProvider styleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            string StyleName = styleProvider.GetStyleName(styleID, seasonID, brandID);

            string basefileName = "FabricColorFastness_ToPDF";
            string openfilepath;
            if (test)
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

            for (int c = 1; c < dtSubDate.Rows.Count; c++)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheetFirst = excel.ActiveWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet worksheetn = excel.ActiveWorkbook.Worksheets[1 + c];
                worksheetFirst.Copy(worksheetn);
            }

            int nSheet = 1;
            for (int i = 0; i < dtSubDate.Rows.Count; i++)
            {
                worksheet = excel.ActiveWorkbook.Worksheets[nSheet];
                worksheet.Cells[3, 4] = dtSubDate.Rows[i]["submitDate"].ToString();
                worksheet.Cells[3, 7] = result.Main.InspDate;
                worksheet.Cells[3, 9] = result.Main.POID;
                worksheet.Cells[3, 12] = brandID;

                worksheet.Cells[5, 4] = styleID;
                worksheet.Cells[5, 10] = CustPONO;
                worksheet.Cells[5, 12] = result.Main.Article;
                worksheet.Cells[6, 4] = StyleName;
                worksheet.Cells[6, 10] = seasonID;

                worksheet.Cells[9, 4] = MyUtility.Check.Empty(result.Main.Temperature) ? "0" : result.Main.Temperature + "˚C";
                worksheet.Cells[9, 7] = MyUtility.Check.Empty(result.Main.Cycle) ? "0" : result.Main.Cycle.ToString();
                worksheet.Cells[9, 9] = result.Main.Detergent;
                worksheet.Cells[10, 4] = result.Main.Machine;
                worksheet.Cells[10, 7] = result.Main.Drying;
                worksheet.Cells[72, 8] = _IColorFastnessProvider.Get_InspectName(result.Main.Inspector);

                List<Fabric_ColorFastness_Detail_Result> dr = new List<Fabric_ColorFastness_Detail_Result>();
                foreach (var item in result.Detail)
                {
                    if (MyUtility.Check.Empty(dtSubDate.Rows[i]["submitDate"]) && MyUtility.Check.Empty(item.SubmitDate))
                    {
                        dr.Add(item);
                    }
                    else if (!MyUtility.Check.Empty(dtSubDate.Rows[i]["submitDate"]) && !MyUtility.Check.Empty(item.SubmitDate))
                    {
                        if (DateTime.Compare(Convert.ToDateTime(item.SubmitDate), Convert.ToDateTime(dtSubDate.Rows[i]["submitDate"])) == 0)
                        {
                            dr.Add(item);
                        }
                    }
                }
                for (int ii = 1; ii < dr.Count; ii++)
                {
                    Microsoft.Office.Interop.Excel.Range rngToInsert = worksheet.get_Range("A13:A13", Type.Missing).EntireRow;
                    rngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                    Marshal.ReleaseComObject(rngToInsert);
                }

                int k = 0;
                foreach (var row in dr)
                {
                    Microsoft.Office.Interop.Excel.Range rang = worksheet.Range[worksheet.Cells[2][13 + k], worksheet.Cells[12][13 + k]];
                    rang.NumberFormat = "@";
                    worksheet.Cells[13 + k, 2] = row.ColorFastnessGroup;
                    worksheet.Cells[13 + k, 3] = row.Seq;
                    worksheet.Cells[13 + k, 4] = row.Roll;
                    worksheet.Cells[13 + k, 5] = row.Dyelot;
                    worksheet.Cells[13 + k, 6] = row.Refno;
                    worksheet.Cells[13 + k, 7] = row.ColorID;
                    worksheet.Cells[13 + k, 8] = row.changeScale.ToString();
                    worksheet.Cells[13 + k, 9] = row.ResultChange;
                    worksheet.Cells[13 + k, 10] = row.StainingScale.ToString();
                    worksheet.Cells[13 + k, 11] = row.ResultStain;
                    worksheet.Cells[13 + k, 12] = row.Remark;
                    rang.Font.Bold = false;
                    rang.Font.Size = 12;

                    // 水平,垂直置中
                    rang.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rang.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    k++;
                }

                nSheet++;
            }
            #region Save & Show Excel

            string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
            string filexlsx = fileName + ".xlsx";
            string fileNamePDF = fileName + ".pdf";

            string filepath;
            string filepathpdf;
            if (test)
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

            if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
            {
                result.reportPath = fileNamePDF;
                result.Result = true;
            }
            else
            {
                result.ErrorMessage = "Convert To PDF Fail";
                result.Result = false;
            }
            #endregion

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel ToExcel(string ID, bool test)
        {
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            result = GetDetailBody(ID);

            if (string.IsNullOrEmpty(result.Main.ID))
            {
                result.Result = false;
                result.ErrorMessage = "ID cannot be empty!";
                return result;
            }

            if (result.Detail.Count == 0)
            {
                result.Result = false;
                result.ErrorMessage = "Detail data not found!";
                return result;
            }

            DataTable dtSubDate = _IColorFastnessDetailProvider.Get_SubmitDate(result.Main.ID);
            if (dtSubDate.Rows.Count < 1)
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

            DataTable dtPO = _IColorFastnessProvider.Get_PO_DataTable(result.Main.POID);
            string styleID, seasonID, brandID;

            if (dtPO.Rows.Count == 0)
            {
                styleID = string.Empty;
                seasonID = string.Empty;
                brandID = string.Empty;
            }
            else
            {
                styleID =   dtPO.Rows[0]["StyleID"].ToString();
                seasonID =  dtPO.Rows[0]["SeasonID"].ToString();
                brandID = dtPO.Rows[0]["BrandID"].ToString();
            }

            IStyleProvider styleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            string StyleName = styleProvider.GetStyleName(styleID, seasonID, brandID);

            string basefileName = "FabricColorFastness_ToExcel";
            string openfilepath;
            if (test)
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

            worksheet.Cells[1, 2] = result.Main.POID;
            worksheet.Cells[1, 4] = styleID;
            worksheet.Cells[1, 6] = seasonID;
            worksheet.Cells[1, 8] = result.Main.Article;
            worksheet.Cells[1, 10] = result.Main.TestNo;
            worksheet.Cells[2, 2] = result.Main.Status;
            worksheet.Cells[2, 4] = result.Main.Result;
            worksheet.Cells[2, 6] = result.Main.InspDate;
            worksheet.Cells[2, 8] = result.Main.Inspector;
            worksheet.Cells[2, 10] = brandID;

            int startRow = 4;
            int idx = 0;
            foreach (var item in result.Detail)
            {
                worksheet.Cells[startRow + idx, 1] = item.ColorFastnessGroup;
                worksheet.Cells[startRow + idx, 2] = item.Seq;
                worksheet.Cells[startRow + idx, 3] = item.Roll;
                worksheet.Cells[startRow + idx, 4] = item.Dyelot;
                worksheet.Cells[startRow + idx, 5] = item.SCIRefno;
                worksheet.Cells[startRow + idx, 6] = item.ColorID;
                worksheet.Cells[startRow + idx, 7] = _IColorFastnessProvider.Get_Supplier(result.Main.POID, item.SEQ1);
                worksheet.Cells[startRow + idx, 8] = item.changeScale;
                worksheet.Cells[startRow + idx, 9] = item.StainingScale;
                worksheet.Cells[startRow + idx, 10] = item.Result;
                worksheet.Cells[startRow + idx, 11] = item.Remark;
                idx++;
            }

            worksheet.Cells.EntireColumn.AutoFit();
            worksheet.Cells.EntireRow.AutoFit();

            worksheet.Select();

            #region Save & Show Excel

            string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
            string filexlsx = fileName + ".xlsx";

            string filepath;
            if (test)
            {
                filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx);
            }
            else
            {
                filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);
            }

            Excel.Workbook workbook = excel.ActiveWorkbook;
            workbook.SaveAs(filepath);
            workbook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);

            result.reportPath = filexlsx;
            result.Result = true;
            #endregion

            return result;
        }

        public BaseResult DeleteColorFastness(string ID)
        {
            BaseResult baseResult = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessProvider = new ColorFastnessProvider(_ISQLDataTransaction);
                baseResult.Result = _IColorFastnessProvider.Delete_ColorFastness(ID);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }
    }
}
