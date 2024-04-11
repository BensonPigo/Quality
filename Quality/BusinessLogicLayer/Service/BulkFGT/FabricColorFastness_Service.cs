using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using Microsoft.IdentityModel.Tokens;
using MICS.DataAccessLayer.Interface;
using MICS.DataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Ocsp;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class FabricColorFastness_Service : IFabricColorFastness_Service
    {
        private IColorFastnessProvider _IColorFastnessProvider;
        private IColorFastnessDetailProvider _IColorFastnessDetailProvider;
        private IOrdersProvider _IOrdersProvider;
        private MailToolsService _MailService;

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
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                result.baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel GetDetailBody(string ID)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            try
            {
                result = _IColorFastnessDetailProvider.Get_DetailBody(ID);

                if (result.Main.TestNo > 0)
                {
                    DataTable dtContent = _IColorFastnessProvider.Get_Mail_Content(result.Main.POID, ID, result.Main.TestNo.ToString());
                    string Subject = $"Washing Fastness Test /{dtContent.Rows[0]["ID"]}/" +
                        $"{dtContent.Rows[0]["StyleID"]}/" +
                        $"{dtContent.Rows[0]["Article"]}/" +
                        $"{dtContent.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                    result.Main.MailSubject = Subject;
                }
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }

        public BaseResult Save_ColorFastness_2ndPage(Fabric_ColorFastness_Detail_ViewModel source, string Mdivision, string UserID)
        {
            BaseResult baseResult = new BaseResult();
            baseResult.Result = true;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(_ISQLDataTransaction);

                // 若表身資料重複 跳訊息


                foreach (var item in source.Details)
                {
                    int repeatCnt = source.Details.Where(s => s.ID == item.ID && s.ColorFastnessGroup == item.ColorFastnessGroup && s.Seq == item.Seq).Count();
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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

                foreach (var item in result.Details)
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

        public BaseResult SentMail(string POID, string ID, string TestNo, string ToAddress, string CCAddress, string Subject, string Body, List<HttpPostedFileBase> Files)
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

                DataTable dtContent = _IColorFastnessProvider.Get_Mail_Content(POID, ID, TestNo);

                string name = $"Washing Fastness Test_{dtContent.Rows[0]["ID"]}_" +
                        $"{dtContent.Rows[0]["StyleID"]}_" +
                        $"{dtContent.Rows[0]["Article"]}_" +
                        $"{dtContent.Rows[0]["Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                Fabric_ColorFastness_Detail_ViewModel ColorFastnessDetailView = ToReport(ID, false, AssignedFineName: name);
                string FileName = ColorFastnessDetailView.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", ColorFastnessDetailView.reportPath) : string.Empty;
                SendMail_Request request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CCAddress,
                    //Subject = "Washing Fastness - Test Fail",
                    Subject = $"Washing Fastness Test /{dtContent.Rows[0]["ID"]}/" +
                        $"{dtContent.Rows[0]["StyleID"]}/" +
                        $"{dtContent.Rows[0]["Article"]}/" +
                        $"{dtContent.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = strHtml,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Washing Fastness",
                    OrderID = POID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(request);
                string buyReadyDate = _MailService.GetBuyReadyDate(request);
                string strHtml = MailTools.DataTableChangeHtml(dtContent, comment, buyReadyDate, Body, out System.Net.Mail.AlternateView plainView);

                request.Body = strHtml;
                request.alternateView = plainView;

                MailTools.SendMail(request);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }


        public Fabric_ColorFastness_Detail_ViewModel ToReport(string ID, bool IsPDF, string AssignedFineName = "")
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

            string tmpName = $"Washing Fastness Test_{dataList[0].POID}_" +
                    $"{dataList[0].StyleID}_" +
                    $"{dataList[0].Article}_" +
                    $"{dataList[0].ColorFastnessResult}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            string basefileName = "FabricColorFastness_ToExcel";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);
            excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出

            Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表
            worksheet.Cells[2, 3] = dataList[0].ReportNo;
            worksheet.Cells[3, 3] = dataList[0].SubmitDate.HasValue ? dataList[0].SubmitDate.Value.ToString("yyyy/MM/dd") : string.Empty;
            worksheet.Cells[3, 8] = DateTime.Now.ToString("yyyy/MM/dd");
            worksheet.Cells[4, 3] = dataList[0].SeasonID;
            worksheet.Cells[4, 8] = dataList[0].BrandID;
            worksheet.Cells[5, 3] = dataList[0].StyleID;
            worksheet.Cells[5, 8] = dataList[0].POID;
            worksheet.Cells[6, 3] = dataList[0].Article;

            worksheet.Cells[9, 2] = dataList[0].Temperature;
            worksheet.Cells[9, 4] = dataList[0].Cycle;
            worksheet.Cells[9, 7] = dataList[0].CycleTime;
            worksheet.Cells[10, 2] = dataList[0].Detergent;
            worksheet.Cells[10, 4] = dataList[0].Machine;
            worksheet.Cells[10, 7] = dataList[0].Drying;

            worksheet.Cells[13, 3] = dataList[0].Checkby;
            Excel.Range cellSignature = worksheet.get_Range("D14:D15");
            if (dataList[0].Signature != null)
            {
                string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList[0].Signature, dataList[0].ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellSignature.Left + 5, cellSignature.Top + 5, cellSignature.Width - 10, cellSignature.Height - 10);
            }

            Excel.Worksheet worksheet2 = excel.ActiveWorkbook.Worksheets[2];
            Excel.Worksheet worksheet3 = excel.ActiveWorkbook.Worksheets[3]; //FOR UA
            int nowRow = 12; //init
            // 表格
            if (dataList[0].BrandID == "U.ARMOUR")
            {
                Excel.Range rngToCopy = worksheet3.get_Range("A1:H1").EntireRow;
                Excel.Range rngToInsert = worksheet.get_Range($"A{nowRow}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上

                nowRow = nowRow + 1; //表格內容起始
                foreach (ColorFastness_Excel item in dataList)
                {
                    rngToCopy = worksheet3.get_Range("A2:H2").EntireRow;
                    rngToInsert = worksheet.get_Range($"A{nowRow}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上

                    worksheet.Cells[nowRow, 1] = item.SEQ;
                    worksheet.Cells[nowRow, 2] = item.Roll;
                    worksheet.Cells[nowRow, 3] = item.Dyelot;
                    worksheet.Cells[nowRow, 4] = item.SCIRefno_Color;
                    worksheet.Cells[nowRow, 5] = item.ChangeScale;
                    worksheet.Cells[nowRow, 6] = item.StainingScale;
                    worksheet.Cells[nowRow, 7] = item.Result;
                    worksheet.Cells[nowRow, 8] = item.Remark;
                    nowRow = nowRow + 1;
                }
            }
            else
            {
                foreach (ColorFastness_Excel item in dataList)
                {
                    Excel.Range rngToCopy = worksheet2.get_Range("A1:H6").EntireRow;
                    Excel.Range rngToInsert = worksheet.get_Range($"A{nowRow}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上

                    worksheet.Cells[nowRow, 2] = item.SEQ;
                    worksheet.Cells[nowRow, 4] = item.Roll;
                    worksheet.Cells[nowRow, 6] = item.Dyelot;
                    worksheet.Cells[nowRow, 8] = item.SCIRefno_Color;

                    worksheet.Cells[nowRow + 3, 2] = item.ChangeScale;
                    worksheet.Cells[nowRow + 3, 3] = item.AcetateScale;
                    worksheet.Cells[nowRow + 3, 4] = item.CottonScale;
                    worksheet.Cells[nowRow + 3, 5] = item.NylonScale;
                    worksheet.Cells[nowRow + 3, 6] = item.PolyesterScale;
                    worksheet.Cells[nowRow + 3, 7] = item.AcrylicScale;
                    worksheet.Cells[nowRow + 3, 8] = item.WoolScale;

                    worksheet.Cells[nowRow + 4, 2] = item.ResultChange;
                    worksheet.Cells[nowRow + 4, 3] = item.ResultAcetate;
                    worksheet.Cells[nowRow + 4, 4] = item.ResultCotton;
                    worksheet.Cells[nowRow + 4, 5] = item.ResultNylon;
                    worksheet.Cells[nowRow + 4, 6] = item.ResultPolyester;
                    worksheet.Cells[nowRow + 4, 7] = item.ResultAcrylic;
                    worksheet.Cells[nowRow + 4, 8] = item.ResultWool;

                    worksheet.Cells[nowRow + 5, 2] = item.Remark;
                    nowRow = nowRow + 6;
                }
            }

            // 圖片
            Excel.Worksheet worksheet4 = excel.ActiveWorkbook.Worksheets[4];
            nowRow = 17 + (dataList.Count * (dataList[0].BrandID == "U.ARMOUR" ? 2 : 6));
            foreach (ColorFastness_Excel item in dataList)
            {
                Excel.Range rngToCopy = worksheet4.get_Range("A1:H42");
                Excel.Range rngToInsert = worksheet.get_Range($"A{nowRow}", Type.Missing); // 選擇要被貼上的位置
                rngToInsert.Insert(rngToCopy.Copy(Type.Missing)); // 貼上

                Excel.Range cellBefore = worksheet.Cells[nowRow + 23, 1];
                if (dataList[0].TestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList[0].TestBeforePicture, dataList[0].ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left, cellBefore.Top, 200, 300);
                }

                Excel.Range cellAfter = worksheet.Cells[nowRow + 23, 5];
                if (dataList[0].TestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList[0].TestAfterPicture, dataList[0].ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left, cellAfter.Top, 200, 300);
                }

                nowRow = nowRow + 42;
            }

            #region Save & Show Excel

            //string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
            if (!string.IsNullOrWhiteSpace(AssignedFineName))
            {
                tmpName = AssignedFineName;
            }
            char[] invalidChars = Path.GetInvalidFileNameChars();

            foreach (char invalidChar in invalidChars)
            {
                tmpName = tmpName.Replace(invalidChar.ToString(), "");
            }
            string filexlsx = tmpName + ".xlsx";
            string fileNamePDF = tmpName + ".pdf";

            string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);
            string filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);

            Excel.Workbook workbook = excel.ActiveWorkbook;
            worksheet2.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            worksheet3.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            worksheet4.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            workbook.SaveAs(filepath);
            workbook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(worksheet2);
            Marshal.ReleaseComObject(worksheet3);
            Marshal.ReleaseComObject(worksheet4);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);

            if (IsPDF && ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
            {
                result.reportPath = fileNamePDF;
                result.Result = true;
            }
            else if (!IsPDF)
            {
                result.reportPath = filexlsx;
                result.Result = true;
            }
            else
            {
                result.ErrorMessage = "Fail";
                result.Result = false;
            }
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }
    }
}
