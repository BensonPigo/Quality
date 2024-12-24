using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
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
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

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
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<ColorFastness_Excel> dataList = _IColorFastnessDetailProvider.GetExcel(ID).ToList();

            if (!dataList.Any())
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

            var testCode = _QualityBrandTestCodeProvider.Get(dataList[0].BrandID, "Washing Fastness");

            string tmpName = string.IsNullOrWhiteSpace(AssignedFineName) ?
                $"Washing Fastness Test_{dataList[0].POID}_{dataList[0].StyleID}_{dataList[0].Article}_{dataList[0].ColorFastnessResult}_{DateTime.Now:yyyyMMddHHmmss}" :
                AssignedFineName;

            tmpName = Path.GetInvalidFileNameChars()
                .Concat(new[] { '-', '+' })
                .Aggregate(tmpName, (current, c) => current.Replace(c.ToString(), ""));

            string baseFileName = "FabricColorFastness_ToExcel";
            string templatePath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "XLT", $"{baseFileName}.xltx");
            string outputDirectory = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP");

            if (!File.Exists(templatePath))
            {
                result.Result = false;
                result.ErrorMessage = "Excel template not found!";
                return result;
            }

            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string excelFileName = $"{tmpName}.xlsx";
            string excelPath = Path.Combine(outputDirectory, excelFileName);
            string pdfFileName = $"{tmpName}.pdf";
            string pdfPath = Path.Combine(outputDirectory, pdfFileName);

            try
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var worksheet2 = workbook.Worksheet(2);
                    var worksheet3 = workbook.Worksheet(3);
                    var worksheet4 = workbook.Worksheet(4);

                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 1).Value = $"Washing Fastness Test Report({testCode.FirstOrDefault().TestCode})";
                    }

                    worksheet.Cell(2, 3).Value = dataList[0].ReportNo;
                    worksheet.Cell(3, 3).Value = dataList[0].SubmitDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell(3, 8).Value = dataList[0].InspDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell(4, 3).Value = dataList[0].SeasonID;
                    worksheet.Cell(4, 8).Value = dataList[0].BrandID;
                    worksheet.Cell(5, 3).Value = dataList[0].StyleID;
                    worksheet.Cell(5, 8).Value = dataList[0].POID;
                    worksheet.Cell(6, 3).Value = dataList[0].Article;

                    worksheet.Cell(9, 2).Value = dataList[0].Temperature;
                    worksheet.Cell(9, 4).Value = dataList[0].Cycle;
                    worksheet.Cell(9, 7).Value = dataList[0].CycleTime;
                    worksheet.Cell(10, 2).Value = dataList[0].Detergent;
                    worksheet.Cell(10, 4).Value = dataList[0].Machine;
                    worksheet.Cell(10, 7).Value = dataList[0].Drying;

                    worksheet.Cell(13, 3).Value = dataList[0].Checkby;

                    // 插入簽名圖片於 D14:D15
                    if (dataList[0].Signature != null)
                    {
                        AddImageToWorksheet(worksheet, dataList[0].Signature, 15, 3, 40, 20);
                    }


                    int nowRow = 12;

                    // 複製與插入邏輯
                    if (dataList[0].BrandID == "U.ARMOUR")
                    {

                        // 插入每筆數據
                        foreach (var item in dataList)
                        {
                            // 來源工作表 = Worksheet3
                            var sourceRange = worksheet3.Range("A1:H2"); // 複製的範圍（標題）

                            worksheet.Row(nowRow).InsertRowsAbove(2);

                            // 插入表頭
                            var destinationRange = worksheet.Range($"A{nowRow}:H{nowRow+1}");
                            sourceRange.CopyTo(destinationRange);

                            nowRow++;

                            // 填充數據
                            worksheet.Cell(nowRow, 1).Value = item.SEQ;
                            worksheet.Cell(nowRow, 2).Value = item.Roll;
                            worksheet.Cell(nowRow, 3).Value = item.Dyelot;
                            worksheet.Cell(nowRow, 4).Value = item.SCIRefno_Color;
                            worksheet.Cell(nowRow, 5).Value = item.ChangeScale;
                            worksheet.Cell(nowRow, 6).Value = item.StainingScale;
                            worksheet.Cell(nowRow, 7).Value = item.Result;
                            worksheet.Cell(nowRow, 8).Value = item.Remark;

                            nowRow++;
                        }
                    }
                    else
                    {
                        foreach (var item in dataList)
                        {
                            // 複製模板行， 非 UA 的模板在第二個 Sheet
                            var sourceRange = worksheet2.Range("A1:H6"); // 複製的範圍（6 行模板）

                            worksheet.Row(nowRow).InsertRowsAbove(6);

                            var destinationRange = worksheet.Range($"A{nowRow}:H{nowRow + 5}");
                            sourceRange.CopyTo(destinationRange);

                            // 填充數據
                            worksheet.Cell(nowRow, 2).Value = item.SEQ;
                            worksheet.Cell(nowRow, 4).Value = item.Roll;
                            worksheet.Cell(nowRow, 6).Value = item.Dyelot;
                            worksheet.Cell(nowRow, 8).Value = item.SCIRefno_Color;

                            worksheet.Cell(nowRow + 3, 2).Value = item.ChangeScale;
                            worksheet.Cell(nowRow + 3, 3).Value = item.AcetateScale;
                            worksheet.Cell(nowRow + 3, 4).Value = item.CottonScale;
                            worksheet.Cell(nowRow + 3, 5).Value = item.NylonScale;
                            worksheet.Cell(nowRow + 3, 6).Value = item.PolyesterScale;
                            worksheet.Cell(nowRow + 3, 7).Value = item.AcrylicScale;
                            worksheet.Cell(nowRow + 3, 8).Value = item.WoolScale;

                            worksheet.Cell(nowRow + 4, 2).Value = item.ResultChange;
                            worksheet.Cell(nowRow + 4, 3).Value = item.ResultAcetate;
                            worksheet.Cell(nowRow + 4, 4).Value = item.ResultCotton;
                            worksheet.Cell(nowRow + 4, 5).Value = item.ResultNylon;
                            worksheet.Cell(nowRow + 4, 6).Value = item.ResultPolyester;
                            worksheet.Cell(nowRow + 4, 7).Value = item.ResultAcrylic;
                            worksheet.Cell(nowRow + 4, 8).Value = item.ResultWool;

                            worksheet.Cell(nowRow + 5, 2).Value = item.Remark;

                            nowRow += 6;
                        }
                    }

                    // 圖片
                    nowRow = 17 + (dataList.Count * (dataList[0].BrandID == "U.ARMOUR" ? 2 : 6));

                    foreach (ColorFastness_Excel item in dataList)
                    {
                        var sourceRange = worksheet4.Range("A1:H42"); // 複製的範圍

                        worksheet.Row(nowRow).InsertRowsAbove(42);

                        var destinationRange = worksheet.Range($"A{nowRow}:H{nowRow + 41}");
                        sourceRange.CopyTo(destinationRange);

                        if (dataList[0].TestBeforePicture != null)
                        {
                            AddImageToWorksheet(worksheet, dataList[0].TestBeforePicture, nowRow + 23, 1, 500, 400);
                        }
                        if (dataList[0].TestAfterPicture != null)
                        {
                            AddImageToWorksheet(worksheet, dataList[0].TestAfterPicture, nowRow + 23, 5, 500, 400);
                        }

                        nowRow = nowRow + 42;
                    }

                    workbook.SaveAs(excelPath);

                    if (IsPDF)
                    {
                        ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                        result.reportPath = pdfFileName;
                    }
                    else
                    {
                        result.reportPath = excelFileName;
                    }

                    result.Result = true;
                }
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $"Error: {ex.Message}, StackTrace: {ex.StackTrace}";
            }

            return result;
        }


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
