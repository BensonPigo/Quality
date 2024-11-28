using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Ocsp;
using ProductionDataAccessLayer.Provider.MSSQL.BukkFGT;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

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
                //DataTable dtOvenDetail = _AccessoryOvenWashProvider.GetOvenTestDataTable(new Accessory_Oven() 
                //{ 
                //    AIR_LaboratoryID = Convert.ToInt64( AIR_LaboratoryID),
                //    POID = POID,
                //    Seq1 = Seq1,
                //    Seq2 = Seq2,
                //});

                Model = _AccessoryOvenWashProvider.GetOvenTestExcel(new Accessory_Oven()
                {
                    AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
                    POID = POID,
                    Seq1 = Seq1,
                    Seq2 = Seq2,
                });

                var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-Oven");

                tmpName = $"Accessory Oven Test"
                    + $"_{POID}"
                    + $"_{Model.StyleID}"
                    + $"_{Model.Refno}"
                    + $"_{Model.Color}"
                    + $"_{Model.OvenResult}"
                    + $"_{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string strXltName = baseFilePath + "\\XLT\\AccessoryOvenTest.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                excel.DisplayAlerts = false;
                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];
                if (testCode.Any())
                {
                    worksheet.Cells[1, 3] = $@"ACCESSORY COLOR MIGRATION TEST REPORT (Oven)({testCode.FirstOrDefault().TestCode})";
                }
                worksheet.Cells[2, 2] = Model.ReportNo;
                worksheet.Cells[2, 6] = Model.POID;
                worksheet.Cells[2, 10] = Model.Supplier;

                worksheet.Cells[3, 2] = Model.BrandID;
                worksheet.Cells[3, 6] = Model.Refno;
                worksheet.Cells[3, 10] = Model.WKNo;

                worksheet.Cells[4, 2] = "Bulk";
                worksheet.Cells[4, 6] = Model.Color;
                worksheet.Cells[4, 10] = string.Empty;

                worksheet.Cells[5, 2] = Model.StyleID;
                worksheet.Cells[5, 6] = Model.Size;
                worksheet.Cells[5, 10] = Model.SeasonID;
                worksheet.Cells[5, 12] = Model.Seq;

                worksheet.Cells[9, 2] = Model.OvenDate.HasValue ? Model.OvenDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                worksheet.Cells[9, 10] = DateTime.Now.ToString("yyyy/MM/dd");

                worksheet.Cells[11, 2] = Model.OvenResult;
                worksheet.Cells[11, 5] = string.Empty;
                worksheet.Cells[11, 10] = string.Empty;

                worksheet.Cells[13, 2] = Model.Remark;

                worksheet.Cells[22, 3] = Model.OvenInspector;
                worksheet.Cells[22, 9] = Model.OvenInspector;

                #region 添加圖片
                Excel.Range cellBeforePicture = worksheet.Cells[20, 1];
                if (Model.OvenTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.OvenTestBeforePicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 400, 300);
                }

                Excel.Range cellAfterPicture = worksheet.Cells[20, 7];
                if (Model.OvenTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.OvenTestAfterPicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 400, 300);
                }

                #endregion


                #region Save & Show Excel

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                char[] invalidChars = Path.GetInvalidFileNameChars();
                char[] additionalChars = { '-', '+' }; // 您想要新增的字元
                char[] updatedInvalidChars = invalidChars.Concat(additionalChars).ToArray();

                foreach (char invalidChar in updatedInvalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }

                string pdfFileName = $"{tmpName}.pdf";
                FileName = $"{tmpName}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", FileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                if (isPDF)
                {
                    bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                    FileName = pdfFileName;
                    if (!isCreatePdfOK)
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
                }

                MyUtility.Excel.KillExcelProcess(excel);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
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

                var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-701Wash");

                tmpName = $"Accessory Wash Test"
                    + $"_{POID}"
                    + $"_{Model.StyleID}"
                    + $"_{Model.Refno}"
                    + $"_{Model.Color}"
                    + $"_{Model.WashResult}"
                    + $"_{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string strXltName = baseFilePath + "\\XLT\\AccessoryWashTest.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                excel.DisplayAlerts = false;
                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];
                if (testCode.Any())
                {
                    worksheet.Cells[1, 3] = $@"ACCESSORY WASH TEST REPORT({testCode.FirstOrDefault().TestCode})";
                }
                worksheet.Cells[2, 2] = Model.ReportNo;
                worksheet.Cells[2, 6] = Model.POID;
                worksheet.Cells[2, 10] = Model.Supplier;

                worksheet.Cells[3, 2] = Model.BrandID;
                worksheet.Cells[3, 6] = Model.Refno;
                worksheet.Cells[3, 10] = Model.WKNo;

                worksheet.Cells[4, 2] = "Bulk";
                worksheet.Cells[4, 6] = Model.Color;
                worksheet.Cells[4, 10] = string.Empty;

                worksheet.Cells[5, 2] = Model.StyleID;
                worksheet.Cells[5, 6] = Model.Size;
                worksheet.Cells[5, 10] = Model.SeasonID;
                worksheet.Cells[5, 12] = Model.Seq;

                worksheet.Cells[9, 2] = Model.WashDate.HasValue ? Model.WashDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                worksheet.Cells[9, 10] = DateTime.Now.ToString("yyyy/MM/dd");

                worksheet.Cells[12, 2] = Model.MachineWash;
                worksheet.Cells[12, 6] = Model.WashingTemperature;
                worksheet.Cells[12, 10] = Model.DryProcess;

                worksheet.Cells[13, 2] = Model.MachineModel;
                worksheet.Cells[13, 6] = Model.WashingCycle;

                worksheet.Cells[15, 2] = Model.WashResult;

                worksheet.Cells[17, 2] = Model.Remark;

                worksheet.Cells[26, 3] = Model.WashInspector;
                worksheet.Cells[26, 9] = Model.WashInspector;

                #region 添加圖片
                Excel.Range cellBeforePicture = worksheet.Cells[24, 1];
                if (Model.WashTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.WashTestBeforePicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 400, 300);
                }

                Excel.Range cellAfterPicture = worksheet.Cells[24, 7];
                if (Model.WashTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.WashTestAfterPicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 400, 300);
                }

                #endregion


                #region Save & Show Excel

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();
                char[] additionalChars = { '-', '+' }; // 您想要新增的字元
                char[] updatedInvalidChars = invalidChars.Concat(additionalChars).ToArray();

                foreach (char invalidChar in updatedInvalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }
                string pdfFileName = $"{tmpName}.pdf";
                FileName = $"{tmpName}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", FileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                if (isPDF)
                {
                    bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                    FileName = pdfFileName;
                    if (!isCreatePdfOK)
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
                }

                MyUtility.Excel.KillExcelProcess(excel);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
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

                var testCode = _QualityBrandTestCodeProvider.Get(Model.BrandID, "Accessory Oven & Wash Test-501Wash");

                tmpName = $"Accessory Washing Fastness Test"
                    + $"_{POID}"
                    + $"_{Model.StyleID}"
                    + $"_{Model.Refno}"
                    + $"_{Model.Color}"
                    + $"_{Model.WashingFastnessResult}"
                    + $"_{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string strXltName = baseFilePath + "\\XLT\\AccessoryWashingFastness.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                excel.Visible = false;
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                excel.DisplayAlerts = false;
                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];

                if (testCode.Any())
                {
                    worksheet.Cells[1, 1] = $@"Washing Fastness (Accessories) ({testCode.FirstOrDefault().TestCode})";
                }
                worksheet.Cells[2, 2] = Model.ReportNo;
                worksheet.Cells[2, 6] = Model.WashingFastnessReceivedDate.HasValue ? Model.WashingFastnessReceivedDate.Value.ToString("yyyy/MM/dd") : string.Empty;

                worksheet.Cells[3, 2] = Model.FactoryID;
                worksheet.Cells[3, 6] = Model.WashingFastnessReportDate.HasValue ? Model.WashingFastnessReportDate.Value.ToString("yyyy/MM/dd") : string.Empty;

                worksheet.Cells[4, 2] = Model.StyleID;

                worksheet.Cells[5, 2] = Model.Article;
                worksheet.Cells[5, 6] = Model.Refno;

                worksheet.Cells[6, 2] = Model.SeasonID;
                worksheet.Cells[6, 6] = Model.Color;

                worksheet.Cells[8, 4] = Model.ChangeScale;
                worksheet.Cells[8, 5] = Model.ResultChange;

                worksheet.Cells[9, 4] = Model.AcetateScale;
                worksheet.Cells[9, 5] = Model.ResultAcetate;

                worksheet.Cells[10, 4] = Model.CottonScale;
                worksheet.Cells[10, 5] = Model.ResultCotton;

                worksheet.Cells[11, 4] = Model.NylonScale;
                worksheet.Cells[11, 5] = Model.ResultNylon;

                worksheet.Cells[12, 4] = Model.PolyesterScale;
                worksheet.Cells[12, 5] = Model.ResultPolyester;

                worksheet.Cells[13, 4] = Model.AcrylicScale;
                worksheet.Cells[13, 5] = Model.ResultAcrylic;

                worksheet.Cells[14, 4] = Model.WoolScale;
                worksheet.Cells[14, 5] = Model.ResultWool;

                worksheet.Cells[15, 4] = Model.CrossStainingScale;
                worksheet.Cells[15, 5] = Model.ResultCrossStaining;


                worksheet.Cells[51, 5] = Model.PreparedText;
                //worksheet.Cells[56, 5] = Model.ExecutiveText;

                if (Model.Conclusions == "APPROVED")
                {
                    worksheet.Cells[48, 2] = "V";
                }
                if (Model.Conclusions == "REJECTED")
                {
                    worksheet.Cells[48, 5] = "V";
                }

                #region 簽名檔圖片
                Excel.Range cellPrepared = worksheet.Cells[51, 5];
                if (Model.Prepared != null)
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

                    byte[] bytes = Model.Prepared;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellPrepared.Left + 2, cellPrepared.Top + 2, 400, 300);
                }


                Excel.Range cellExecutive = worksheet.Cells[56, 5];
                if (Model.Executive != null)
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

                    byte[] bytes = Model.Executive;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellExecutive.Left + 2, cellExecutive.Top + 2, 400, 300);
                }
                #endregion

                #region 添加圖片
                Excel.Range cellBeforePicture = worksheet.Cells[33, 1];
                if (Model.WashingFastnessTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.WashingFastnessTestBeforePicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 400, 300);
                }

                Excel.Range cellAfterPicture = worksheet.Cells[33, 5];
                if (Model.WashingFastnessTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(Model.WashingFastnessTestAfterPicture, Model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 400, 300);
                }

                #endregion


                #region Save & Show Excel

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                char[] invalidChars = Path.GetInvalidFileNameChars();
                char[] additionalChars = { '-', '+' }; // 您想要新增的字元
                char[] updatedInvalidChars = invalidChars.Concat(additionalChars).ToArray();

                foreach (char invalidChar in updatedInvalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }
                string pdfFileName = $"{tmpName}.pdf";
                FileName = $"{tmpName}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", FileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                if (isPDF)
                {
                    bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                    FileName = pdfFileName;
                    if (!isCreatePdfOK)
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
                }

                MyUtility.Excel.KillExcelProcess(excel);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }



            return result;
        }
        #endregion
    }
}
