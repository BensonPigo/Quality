using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Provider.MSSQL.BukkFGT;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AccessoryOvenWashService
    {
        private AccessoryOvenWashProvider _AccessoryOvenWashProvider;
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;
        private string IsTest = ConfigurationManager.AppSettings["IsTest"].ToString();

        public Accessory_ViewModel GetMainData(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetHead(Req.ReqOrderID);
                result.DataList = _AccessoryOvenWashProvider.GetDetail(Req.ReqOrderID).ToList();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError('{ex.Message.Replace("'", string.Empty)}');";
            }

            return result;
        }

        public Accessory_ViewModel Update(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                int r = _AccessoryOvenWashProvider.Update_AIR_Laboratory(Req);

                result.Result = r > 0;

                _AccessoryOvenWashProvider.Update_AIR_Laboratory_AllResult(Req);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError('{ex.Message.Replace("'", string.Empty)}');";
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public void UpdateInspPercent(string POID)
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
            _AccessoryOvenWashProvider.UpdateInspPercent(POID);
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
        #region Oven
        public Accessory_Oven GetOvenTest(Accessory_Oven Req)
        {

            Accessory_Oven result = new Accessory_Oven()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetOvenTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                DataTable dt = _AccessoryOvenWashProvider.GetData_OvenDataTable(Req);
                string Subject = $"Accessory Oven Test/{Req.POID}/" +
                    $"{dt.Rows[0]["Style"]}/" +
                    $"{dt.Rows[0]["Refno"]}/" +
                    $"{dt.Rows[0]["Color"]}/" +
                    $"{dt.Rows[0]["Oven Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                result.MailSubject = Subject;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message.Replace("'", string.Empty)}"");";
            }

            return result;
        }

        public Accessory_Oven UpdateOven(Accessory_Oven Req)
        {
            Accessory_Oven result = new Accessory_Oven();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);
                AccessoryOvenWashProvider _MESProvider = new AccessoryOvenWashProvider(Common.ManufacturingExecutionDataAccessLayer);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateOvenTest(Req);
                int Mes_r = _MESProvider.UpdateOvenTestPic(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Oven_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError('{ex.Message.Replace("'", string.Empty)}');";
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public Accessory_Oven EncodeAmendOven(Accessory_Oven Req)
        {
            Accessory_Oven result = new Accessory_Oven();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                var check = _AccessoryOvenWashProvider.Oven_EncodeCheck(Req);
                int r = _AccessoryOvenWashProvider.EncodeAmendOven(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Oven_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public SendMail_Result SendOvenMail(Accessory_Oven Req, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_OvenDataTable(Req);
                string name = $"Accessory Oven Test_{Req.POID}_" +
                    $"{dt.Rows[0]["Style"]}_" +
                    $"{dt.Rows[0]["Refno"]}_" +
                    $"{dt.Rows[0]["Color"]}_" +
                    $"{dt.Rows[0]["Oven Result"]}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = OvenTestExcel(Req.AIR_LaboratoryID.ToString(), Req.POID, Req.Seq1, Req.Seq2, true, out string excelFileName, AssignedFineName: name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName)
                    : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    // 測試名稱/SP號/款號/料號/顏色/測試結果/寄信年月日時分秒
                    Subject = $"Accessory Oven Test/{Req.POID}/" +
                    $"{dt.Rows[0]["Style"]}/" +
                    $"{dt.Rows[0]["Refno"]}/" +
                    $"{dt.Rows[0]["Color"]}/" +
                    $"{dt.Rows[0]["Oven Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Accessory Oven & Wash Test",
                    OrderID = Req.POID,

                };

                if (!string.IsNullOrEmpty(Req.Subject))
                {
                    sendMail_Request.Subject = Req.Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, Req.Body, out System.Net.Mail.AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }


            return result;
        }

        public BaseResult OvenTestExcel(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool isPDF, out string FileName, string AssignedFineName = "")
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            BaseResult result = new BaseResult();

            Accessory_OvenExcel Model = new Accessory_OvenExcel();

            FileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");

                Model = _AccessoryOvenWashProvider.GetOvenTestExcel(new Accessory_Oven()
                {
                    AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
                    POID = POID,
                    Seq1 = Seq1,
                    Seq2 = Seq2,
                });

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-Oven");

                tmpName = $"Accessory Oven Test_{POID}_{Model.StyleID}_{Model.Refno}_{Model.Color}_{Model.OvenResult}_{DateTime.Now:yyyyMMddHHmmss}";

                string templatePath = Path.Combine(baseFilePath, "XLT", "AccessoryOvenTest.xltx");

                if (!File.Exists(templatePath))
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入主表數據
                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 3).Value = $@"ACCESSORY COLOR MIGRATION TEST REPORT (Oven)({testCode.First().TestCode})";
                    }
                    worksheet.Cell("B2").Value = Model.ReportNo;
                    worksheet.Cell("F2").Value = Model.POID;
                    worksheet.Cell("J2").Value = Model.Supplier;

                    worksheet.Cell("B3").Value = Model.BrandID;
                    worksheet.Cell("F3").Value = Model.Refno;
                    worksheet.Cell("J3").Value = Model.WKNo;

                    worksheet.Cell("B4").Value = "Bulk";
                    worksheet.Cell("F4").Value = Model.Color;
                    worksheet.Cell("J4").Value = string.Empty;

                    worksheet.Cell("B5").Value = Model.StyleID;
                    worksheet.Cell("F5").Value = Model.Size;
                    worksheet.Cell("J5").Value = Model.SeasonID;
                    worksheet.Cell("L5").Value = Model.Seq;

                    worksheet.Cell("B9").Value = Model.OvenDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("J9").Value = DateTime.Now.ToString("yyyy/MM/dd");

                    worksheet.Cell("B11").Value = Model.OvenResult;
                    worksheet.Cell("E11").Value = string.Empty;
                    worksheet.Cell("J11").Value = string.Empty;

                    worksheet.Cell("B13").Value = Model.Remark;

                    worksheet.Cell("C22").Value = Model.OvenInspector;
                    worksheet.Cell("I22").Value = Model.OvenInspector;

                    // 插入圖片（使用共用方法）
                    AddImageToWorksheet(worksheet, Model.OvenTestBeforePicture, 20, 1, 400, 300);
                    AddImageToWorksheet(worksheet, Model.OvenTestAfterPicture, 20, 7, 400, 300);

                    workbook.SaveAs(outputFilePath);
                    FileName = $"{tmpName}.xlsx";
                }

                if (isPDF)
                {
                    string pdfPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.pdf");
                    if (ConvertToPDF.ExcelToPDF(outputFilePath, pdfPath))
                    {
                        FileName = $"{tmpName}.pdf";
                    }
                    else
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
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
        #endregion

        #region Wash
        public Accessory_Wash GetWashTest(Accessory_Wash Req)
        {

            Accessory_Wash result = new Accessory_Wash()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetWashTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                DataTable dt = _AccessoryOvenWashProvider.GetData_WashDataTable(Req);
                string Subject = $"Accessory Wash Test/{Req.POID}/" +
                    $"{dt.Rows[0]["Style"]}/" +
                    $"{dt.Rows[0]["Refno"]}/" +
                    $"{dt.Rows[0]["Color"]}/" +
                    $"{dt.Rows[0]["Wash Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                result.MailSubject = Subject;

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError('{ex.Message.Replace("'", string.Empty)}');";
            }

            return result;
        }

        public Accessory_Wash UpdateWash(Accessory_Wash Req)
        {
            Accessory_Wash result = new Accessory_Wash();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);
                AccessoryOvenWashProvider _MESProvider = new AccessoryOvenWashProvider(Common.ManufacturingExecutionDataAccessLayer);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateWashTest(Req);
                int Mes_r = _MESProvider.UpdateWashTestPic(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Wash_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError('{ex.Message.Replace("'", string.Empty)}');";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public Accessory_Wash EncodeAmendWash(Accessory_Wash Req)
        {
            Accessory_Wash result = new Accessory_Wash();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                var check = _AccessoryOvenWashProvider.Wash_EncodeCheck(Req);
                int r = _AccessoryOvenWashProvider.EncodeAmendWash(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Wash_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $"msg.WithError( \"{ex.Message}\" ); ";// "msg.WithError(\""+ ex.Message+"\"); ";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }
        public SendMail_Result SendWashMail(Accessory_Wash Req, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_WashDataTable(Req);
                string name = $"Accessory Wash Test_{Req.POID}_" +
                    $"{dt.Rows[0]["Style"]}_" +
                    $"{dt.Rows[0]["Refno"]}_" +
                    $"{dt.Rows[0]["Color"]}_" +
                    $"{dt.Rows[0]["Wash Result"]}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = WashTestExcel(Req.AIR_LaboratoryID.ToString(), Req.POID, Req.Seq1, Req.Seq2, true, out string excelFileName, AssignedFineName: name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = $"Accessory Wash Test/{Req.POID}/" +
                    $"{dt.Rows[0]["Style"]}/" +
                    $"{dt.Rows[0]["Refno"]}/" +
                    $"{dt.Rows[0]["Color"]}/" +
                    $"{dt.Rows[0]["Wash Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Accessory Oven & Wash Test",
                    OrderID = Req.POID,
                };

                if (!string.IsNullOrEmpty(Req.Subject))
                {
                    sendMail_Request.Subject = Req.Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, Req.Body, out System.Net.Mail.AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }


            return result;
        }
        public BaseResult WashTestExcel(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool isPDF, out string FileName, string AssignedFineName = "")
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

            BaseResult result = new BaseResult();
            Accessory_WashExcel Model = new Accessory_WashExcel();

            FileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");

                Model = _AccessoryOvenWashProvider.GetWashTestExcel(new Accessory_Wash()
                {
                    AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
                    POID = POID,
                    Seq1 = Seq1,
                    Seq2 = Seq2,
                });

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-701Wash");

                tmpName = $"Accessory Wash Test_{POID}_{Model.StyleID}_{Model.Refno}_{Model.Color}_{Model.WashResult}_{DateTime.Now:yyyyMMddHHmmss}";

                string templatePath = Path.Combine(baseFilePath, "XLT", "AccessoryWashTest.xltx");

                if (!File.Exists(templatePath))
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入主表數據
                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 3).Value = $@"ACCESSORY WASH TEST REPORT({testCode.First().TestCode})";
                    }
                    worksheet.Cell("B2").Value = Model.ReportNo;
                    worksheet.Cell("F2").Value = Model.POID;
                    worksheet.Cell("J2").Value = Model.Supplier;

                    worksheet.Cell("B3").Value = Model.BrandID;
                    worksheet.Cell("F3").Value = Model.Refno;
                    worksheet.Cell("J3").Value = Model.WKNo;

                    worksheet.Cell("B4").Value = "Bulk";
                    worksheet.Cell("F4").Value = Model.Color;
                    worksheet.Cell("J4").Value = string.Empty;

                    worksheet.Cell("B5").Value = Model.StyleID;
                    worksheet.Cell("F5").Value = Model.Size;
                    worksheet.Cell("J5").Value = Model.SeasonID;
                    worksheet.Cell("L5").Value = Model.Seq;

                    worksheet.Cell("B9").Value = Model.WashDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("J9").Value = DateTime.Now.ToString("yyyy/MM/dd");

                    worksheet.Cell("B12").Value = Model.MachineWash;
                    worksheet.Cell("F12").Value = Model.WashingTemperature;
                    worksheet.Cell("J12").Value = Model.DryProcess;

                    worksheet.Cell("B13").Value = Model.MachineModel;
                    worksheet.Cell("F13").Value = Model.WashingCycle;

                    worksheet.Cell("B15").Value = Model.WashResult;
                    worksheet.Cell("B17").Value = Model.Remark;

                    worksheet.Cell("C26").Value = Model.WashInspector;
                    worksheet.Cell("I26").Value = Model.WashInspector;

                    // 插入圖片（使用共用方法）
                    AddImageToWorksheet(worksheet, Model.WashTestBeforePicture, 24, 1, 400, 300);
                    AddImageToWorksheet(worksheet, Model.WashTestAfterPicture, 24, 7, 400, 300);

                    workbook.SaveAs(outputFilePath);
                    FileName = $"{tmpName}.xlsx";
                }

                if (isPDF)
                {
                    string pdfPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.pdf");
                    if (ConvertToPDF.ExcelToPDF(outputFilePath, pdfPath))
                    {
                        FileName = $"{tmpName}.pdf";
                    }
                    else
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
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

        #endregion

        #region WashingFastness
        public Accessory_WashingFastness GetWashingFastness(Accessory_WashingFastness Req)
        {

            Accessory_WashingFastness result = new Accessory_WashingFastness()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetWashingFastness(Req);
                result.ChangeScale = string.IsNullOrEmpty(result.ChangeScale) ? "4-5" : result.ChangeScale;
                result.CrossStainingScale = string.IsNullOrEmpty(result.CrossStainingScale) ? "4-5" : result.CrossStainingScale;
                result.AcetateScale = string.IsNullOrEmpty(result.AcetateScale) ? "4-5" : result.AcetateScale;
                result.CottonScale = string.IsNullOrEmpty(result.CottonScale) ? "4-5" : result.CottonScale;
                result.NylonScale = string.IsNullOrEmpty(result.NylonScale) ? "4-5" : result.NylonScale;
                result.PolyesterScale = string.IsNullOrEmpty(result.PolyesterScale) ? "4-5" : result.PolyesterScale;
                result.AcrylicScale = string.IsNullOrEmpty(result.AcrylicScale) ? "4-5" : result.AcrylicScale;
                result.WoolScale = string.IsNullOrEmpty(result.WoolScale) ? "4-5" : result.WoolScale;
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                DataTable dt = _AccessoryOvenWashProvider.GetData_WashingFastnessDataTable(Req);

                string Subject = $"Accessory Washing Fastness Test_" +
                        $"{Req.POID}_" +
                        $"{dt.Rows[0]["Style"]}_" +
                        $"{dt.Rows[0]["Refno"]}_" +
                        $"{dt.Rows[0]["Color"]}_" +
                        $"{dt.Rows[0]["Washing Fastness Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                result.MailSubject = Subject;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return result;
        }

        public Accessory_WashingFastness UpdateWashingFastness(Accessory_WashingFastness Req)
        {
            Accessory_WashingFastness result = new Accessory_WashingFastness();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);
                AccessoryOvenWashProvider _MESProvider = new AccessoryOvenWashProvider(Common.ManufacturingExecutionDataAccessLayer);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateWashingFastness(Req);
                int Mes_r = _MESProvider.UpdateWashingFastnessPic(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_WashingFastness_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public Accessory_WashingFastness EncodeAmendWashingFastness(Accessory_WashingFastness Req)
        {
            Accessory_WashingFastness result = new Accessory_WashingFastness();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                var check = _AccessoryOvenWashProvider.WashingFastness_EncodeCheck(Req);
                int r = _AccessoryOvenWashProvider.UpdateWashingFastness_WashEncodeAmend(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_WashingFastness_AllResult(Req);
                _AccessoryOvenWashProvider.UpdateInspPercent(Req.POID);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }
        public SendMail_Result SendWashingFastnessMail(Accessory_WashingFastness Req, List<HttpPostedFileBase> Files)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_WashingFastnessDataTable(Req);

                string name = $"Accessory Washing Fastness Test_" +
                        $"{Req.POID}_" +
                        $"{dt.Rows[0]["Style"]}_" +
                        $"{dt.Rows[0]["Refno"]}_" +
                        $"{dt.Rows[0]["Color"]}_" +
                        $"{dt.Rows[0]["Washing Fastness Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = WashingFastnessExcel(Req.AIR_LaboratoryID.ToString(), Req.POID, Req.Seq1, Req.Seq2, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;


                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = $"Accessory Washing Fastness Test/" +
                        $"{Req.POID}/" +
                        $"{dt.Rows[0]["Style"]}/" +
                        $"{dt.Rows[0]["Refno"]}/" +
                        $"{dt.Rows[0]["Color"]}/" +
                        $"{dt.Rows[0]["Washing Fastness Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Accessory Oven & Wash Test",
                    OrderID = Req.POID,
                };

                if (!string.IsNullOrEmpty(Req.Subject))
                {
                    sendMail_Request.Subject = Req.Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, Req.Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }


            return result;
        }        
        public BaseResult WashingFastnessExcel(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool isPDF, out string FileName, string AssignedFineName = "")
{
    _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
    _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

    BaseResult result = new BaseResult();
    Accessory_WashingFastnessExcel Model = new Accessory_WashingFastnessExcel();

    FileName = string.Empty;
    string tmpName = string.Empty;

    try
    {
        string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");

        Model = _AccessoryOvenWashProvider.GetWashingFastnessExcel(new Accessory_WashingFastness()
        {
            AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
            POID = POID,
            Seq1 = Seq1,
            Seq2 = Seq2,
        });

        if (Model == null)
        {
            result.ErrorMessage = "Data not found!";
            result.Result = false;
            return result;
        }

        var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-501Wash");

        tmpName = $"Accessory Washing Fastness Test_{POID}_{Model.StyleID}_{Model.Refno}_{Model.Color}_{Model.WashingFastnessResult}_{DateTime.Now:yyyyMMddHHmmss}";

        string templatePath = Path.Combine(baseFilePath, "XLT", "AccessoryWashingFastness.xltx");

        if (!File.Exists(templatePath))
        {
            result.ErrorMessage = "Excel template not found!";
            result.Result = false;
            return result;
        }

        string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

        using (var workbook = new XLWorkbook(templatePath))
        {
            var worksheet = workbook.Worksheet(1);

            // 填入主表數據
            if (testCode.Any())
            {
                worksheet.Cell(1, 1).Value = $@"Washing Fastness (Accessories) ({testCode.First().TestCode})";
            }
            worksheet.Cell("B2").Value = Model.ReportNo;
            worksheet.Cell("F2").Value = Model.WashingFastnessReceivedDate?.ToString("yyyy/MM/dd") ?? string.Empty;

            worksheet.Cell("B3").Value = Model.FactoryID;
            worksheet.Cell("F3").Value = Model.WashingFastnessReportDate?.ToString("yyyy/MM/dd") ?? string.Empty;

            worksheet.Cell("B4").Value = Model.StyleID;
            worksheet.Cell("B5").Value = Model.Article;
            worksheet.Cell("F5").Value = Model.Refno;

            worksheet.Cell("B6").Value = Model.SeasonID;
            worksheet.Cell("F6").Value = Model.Color;

            worksheet.Cell("D8").Value = Model.ChangeScale;
            worksheet.Cell("E8").Value = Model.ResultChange;

            worksheet.Cell("D9").Value = Model.AcetateScale;
            worksheet.Cell("E9").Value = Model.ResultAcetate;

            worksheet.Cell("D10").Value = Model.CottonScale;
            worksheet.Cell("E10").Value = Model.ResultCotton;

            worksheet.Cell("D11").Value = Model.NylonScale;
            worksheet.Cell("E11").Value = Model.ResultNylon;

            worksheet.Cell("D12").Value = Model.PolyesterScale;
            worksheet.Cell("E12").Value = Model.ResultPolyester;

            worksheet.Cell("D13").Value = Model.AcrylicScale;
            worksheet.Cell("E13").Value = Model.ResultAcrylic;

            worksheet.Cell("D14").Value = Model.WoolScale;
            worksheet.Cell("E14").Value = Model.ResultWool;

            worksheet.Cell("D15").Value = Model.CrossStainingScale;
            worksheet.Cell("E15").Value = Model.ResultCrossStaining;

            worksheet.Cell("E51").Value = Model.PreparedText;

            if (Model.Conclusions == "APPROVED")
            {
                worksheet.Cell("B48").Value = "V";
            }
            if (Model.Conclusions == "REJECTED")
            {
                worksheet.Cell("E48").Value = "V";
            }

            // 簽名檔圖片
            AddImageToWorksheet(worksheet, Model.Prepared, 51, 5, 400, 300);
            AddImageToWorksheet(worksheet, Model.Executive, 56, 5, 400, 300);

            // 測試前後圖片
            AddImageToWorksheet(worksheet, Model.WashingFastnessTestBeforePicture, 33, 1, 400, 300);
            AddImageToWorksheet(worksheet, Model.WashingFastnessTestAfterPicture, 33, 5, 400, 300);

            workbook.SaveAs(outputFilePath);
            FileName = $"{tmpName}.xlsx";
        }

        if (isPDF)
        {
            string pdfPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.pdf");
            if (ConvertToPDF.ExcelToPDF(outputFilePath, pdfPath))
            {
                FileName = $"{tmpName}.pdf";
            }
            else
            {
                result.Result = false;
                result.ErrorMessage = "ConvertToPDF fail";
                return result;
            }
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
        #endregion
    }
}
