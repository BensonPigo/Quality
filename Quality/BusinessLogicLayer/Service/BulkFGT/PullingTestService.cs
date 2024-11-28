using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Interop.Excel;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Ocsp;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Web;
using System.Web.UI.WebControls;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PullingTestService
    {
        private PullingTestProvider _PullingTestProvider;
        private bool IsTest = bool.Parse(ConfigurationManager.AppSettings["IsTest"]);
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public PullingTest_ViewModel GetReportNoList(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel result = Req;

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.ReportNo_Source = _PullingTestProvider.GetReportNoList(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }


        public PullingTest_ViewModel GetData(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                result.Detail = _PullingTestProvider.GetData(ReportNo);
                System.Data.DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);
                string Subject = $"Pulling Test/{dt.Rows[0]["POID"]}/" +
                        $"{dt.Rows[0]["StyleID"]}/" +
                        $"{dt.Rows[0]["Article"]}/" +
                        $"{dt.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                result.Detail.MailSubject = Subject;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_Result CheckSP(string POID)
        {
            PullingTest_Result result = new PullingTest_Result();
            PullingTest_Result result2 = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.CheckSP(POID);

                // 透過Brand得到PullForceUnit
                result2 = _PullingTestProvider.GetPullForceUnit(result.BrandID);

                result.PullForceUnit = result2.PullForceUnit;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }


        public PullingTest_Result GetStandard(string BrandID, string TestItem, string PullForceUnit, string StyleType)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.GetStandard(BrandID, TestItem, PullForceUnit, StyleType);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_Result GetPullUnit(string BrandID)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.GetPullForceUnit(BrandID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public PullingTest_ViewModel Insert(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Insert(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Update(PullingTest_Result Req)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Update(Req);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        public PullingTest_ViewModel Delete(string ReportNo)
        {
            PullingTest_ViewModel result = new PullingTest_ViewModel();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                _PullingTestProvider.Delete(ReportNo);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }


        public SendMail_Result SendMail(string ReportNo, string ToAddress, string CcAddress, string Subject, string Body, List<HttpPostedFileBase> Files)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                System.Data.DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);
                string name = $"Pulling Test_{dt.Rows[0]["POID"]}_" +
                        $"{dt.Rows[0]["StyleID"]}_" +
                        $"{dt.Rows[0]["Article"]}_" +
                        $"{dt.Rows[0]["Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                Report_Result baseResult = GetPDF(ReportNo, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;

                string unit = dt.Rows[0]["PullForceUnit"].ToString();
                if (!string.IsNullOrEmpty(unit))
                {
                    dt.Columns["PullForceUnit"].ColumnName = unit;
                    dt.Rows[0][unit] = dt.Rows[0]["PullForce"].ToString();
                    dt.Columns.Remove("PullForce");
                }

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CcAddress,
                    //Subject = "Pulling Test - Test Fail",
                    Subject = $"Pulling Test/{dt.Rows[0]["POID"]}/" +
                        $"{dt.Rows[0]["StyleID"]}/" +
                        $"{dt.Rows[0]["Article"]}/" +
                        $"{dt.Rows[0]["Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Pulling test for Snap/Button/Rivet",
                    StyleID = dt.Rows[0]["StyleID"].ToString(),
                    SeasonID = dt.Rows[0]["SeasonID"].ToString(),
                    BrandID = dt.Rows[0]["BrandID"].ToString(),
                };

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, Body, out AlternateView plainView);

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

        public Report_Result GetPDF(string ReportNo, string AssignedFineName = "")
        {
            Report_Result result = new Report_Result();
            if (string.IsNullOrEmpty(ReportNo))
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            string basefileName = "PullingTest";
            string tmpName = string.Empty;

            try
            {
                if (!this.IsTest)
                {
                    if (!Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                    {
                        Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                    }

                    if (!Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                    {
                        Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                    }
                }

                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
                PullingTest_Result model = _PullingTestProvider.GetData(ReportNo);
                var testCode = _QualityBrandTestCodeProvider.Get(model.BrandID, "Pulling test for Snap/Button/Rivet");
                System.Data.DataTable ReportTechnician = _PullingTestProvider.GetReportTechnician(ReportNo);

                string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx"; ;
                if (this.IsTest)
                {
                    openfilepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLT", $"{basefileName}.xltx");
                }

                tmpName = $"Pulling Test_{model.POID}_" +
                        $"{model.StyleID}_" +
                        $"{model.Article}_" +
                        $"{model.Result}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);
                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];

                if (testCode.Any())
                {
                    worksheet.Cells[1, 1] = $@"Pulling test for Snap/Button/Rivet Report({testCode.FirstOrDefault().TestCode})";
                }

                worksheet.Cells[2, 2] = model.ReportNo;
                worksheet.Cells[2, 4] = DateTime.Now.ToString("yyyy/MM/dd");
                worksheet.Cells[3, 2] = model.POID;
                worksheet.Cells[3, 4] = model.TestDateText;
                worksheet.Cells[4, 2] = model.SeasonID;
                worksheet.Cells[4, 4] = model.StyleID;
                worksheet.Cells[5, 2] = model.BrandID;
                worksheet.Cells[5, 4] = model.Article;
                worksheet.Cells[6, 2] = model.SizeCode;
                worksheet.Cells[6, 4] = model.InspectorName;
                worksheet.Cells[7, 2] = model.TestItem;
                worksheet.Cells[7, 3] = model.PullForceUnit;
                worksheet.Cells[7, 4] = model.PullForce;
                worksheet.Cells[8, 2] = model.Time;
                worksheet.Cells[8, 4] = model.Gender;
                worksheet.Cells[9, 2] = model.FabricRefno;
                worksheet.Cells[9, 4] = model.AccRefno;
                worksheet.Cells[10, 2] = model.SnapOperator;
                worksheet.Cells[10, 4] = model.Result;
                worksheet.Cells[12, 1] = model.Remark;
                worksheet.Cells[22, 4] = model.AddName;

                Range cell = worksheet.Cells[23, 4];
                if (model.Signature != null)
                {
                    MemoryStream ms = new MemoryStream(model.Signature);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;
                    if (this.IsTest)
                    {
                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, cell.Width - 10, cell.Height - 10);
                }


                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[22, 4] = TechnicianName;

                    // Signture 圖片
                    cell = worksheet.Cells[23, 4];
                    if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    }
                }

                // TestBeforePicture 圖片
                if (model.TestBeforePicture != null && model.TestBeforePicture.Length > 1)
                {
                    cell = worksheet.Cells[15, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.TestBeforePicture, model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // TestAfterPicture 圖片
                if (model.TestAfterPicture != null && model.TestAfterPicture.Length > 1)
                {
                    cell = worksheet.Cells[15, 3];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.TestAfterPicture, model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

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

                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepath;
                string filepathpdf;
                if (this.IsTest)
                {
                    filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx);
                    filepathpdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF);
                }
                else
                {
                    filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);
                    filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                }

                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                MyUtility.Excel.KillExcelProcess(excelApp);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                result.TempFileName = filexlsx;
                result.Result = true;

                if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                {
                    result.TempFileName = fileNamePDF;
                    result.Result = true;
                }
                else
                {
                    result.ErrorMessage = "Convert To PDF Fail";
                    result.Result = false;
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
    }
}
