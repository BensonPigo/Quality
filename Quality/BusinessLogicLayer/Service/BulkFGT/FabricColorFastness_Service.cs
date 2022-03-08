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

        public List<ColorFastness_Excel> GetExcel(string ID)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            List<ColorFastness_Excel> result = new List<ColorFastness_Excel>();
            try
            {
                result = _IColorFastnessDetailProvider.GetExcel(ID).ToList();
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
                string strHtml = MailTools.DataTableChangeHtml(dtContent, out System.Net.Mail.AlternateView plainView);

                SendMail_Request request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CCAddress,
                    Subject = "Washing Fastness-Crocking Test – Test Fail",
                    Body = strHtml,
                    alternateView = plainView,
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
            List<ColorFastness_Excel> dataList = new List<ColorFastness_Excel>();

            dataList = _IColorFastnessDetailProvider.GetExcel(ID).ToList();
            if (!dataList.Any())
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

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
                ColorFastness_Excel currenData = dataList[j - 1];

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
                currenSheet.Cells[8, 8] = currenData.CycleTime;

                currenSheet.Cells[12, 2] = currenData.ChangeScale;
                currenSheet.Cells[12, 3] = currenData.AcetateScale;
                currenSheet.Cells[12, 4] = currenData.CottonScale;
                currenSheet.Cells[12, 5] = currenData.NylonScale;
                currenSheet.Cells[12, 6] = currenData.PolyesterScale;
                currenSheet.Cells[12, 7] = currenData.AcrylicScale;
                currenSheet.Cells[12, 8] = currenData.WoolScale;

                currenSheet.Cells[13, 2] = currenData.ResultChange;
                currenSheet.Cells[13, 3] = currenData.ResultAcetate;
                currenSheet.Cells[13, 4] = currenData.ResultCotton;
                currenSheet.Cells[13, 5] = currenData.ResultNylon;
                currenSheet.Cells[13, 6] = currenData.ResultPolyester;
                currenSheet.Cells[13, 7] = currenData.ResultAcrylic;
                currenSheet.Cells[13, 8] = currenData.ResultWool;

                currenSheet.Cells[14, 2] = currenData.Remark;
                currenSheet.Cells[70, 3] = currenData.Inspector;


                #region 添加圖片
                Excel.Range cellBeforePicture = currenSheet.Cells[45, 1];
                if (currenData.TestBeforePicture != null)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;

                    if (test)
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

                Excel.Range cellAfterPicture = currenSheet.Cells[45, 5];
                if (currenData.TestAfterPicture != null)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;

                    if (test)
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
            List<ColorFastness_Excel> dataList = new List<ColorFastness_Excel>();

            dataList = _IColorFastnessDetailProvider.GetExcel(ID).ToList();
            if (!dataList.Any())
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

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
                ColorFastness_Excel currenData = dataList[j-1];

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
                currenSheet.Cells[8, 8] = currenData.CycleTime;

                currenSheet.Cells[12, 2] = currenData.ChangeScale;
                currenSheet.Cells[12, 3] = currenData.AcetateScale;
                currenSheet.Cells[12, 4] = currenData.CottonScale;
                currenSheet.Cells[12, 5] = currenData.NylonScale;
                currenSheet.Cells[12, 6] = currenData.PolyesterScale;
                currenSheet.Cells[12, 7] = currenData.AcrylicScale;
                currenSheet.Cells[12, 8] = currenData.WoolScale;

                currenSheet.Cells[13, 2] = currenData.ResultChange;
                currenSheet.Cells[13, 3] = currenData.ResultAcetate;
                currenSheet.Cells[13, 4] = currenData.ResultCotton;
                currenSheet.Cells[13, 5] = currenData.ResultNylon;
                currenSheet.Cells[13, 6] = currenData.ResultPolyester;
                currenSheet.Cells[13, 7] = currenData.ResultAcrylic;
                currenSheet.Cells[13, 8] = currenData.ResultWool;

                currenSheet.Cells[14, 2] = currenData.Remark;
                currenSheet.Cells[70, 3] = currenData.Inspector;


                #region 添加圖片
                Excel.Range cellBeforePicture = currenSheet.Cells[45, 1];
                if (currenData.TestBeforePicture != null)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;

                    if (test)
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

                Excel.Range cellAfterPicture = currenSheet.Cells[45, 5];
                if (currenData.TestAfterPicture != null)
                {
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;

                    if (test)
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
