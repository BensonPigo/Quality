using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Interop.Excel;
using Org.BouncyCastle.Asn1.Ocsp;
using Sci;
using System;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.UI.WebControls;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PullingTestService
    {
        private PullingTestProvider _PullingTestProvider;
        private bool IsTest = bool.Parse(ConfigurationManager.AppSettings["IsTest"]);
        private MailToolsService _MailService;

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


        public PullingTest_Result GetStandard(string BrandID, string TestItem, string PullForceUnit)
        {
            PullingTest_Result result = new PullingTest_Result();

            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ProductionDataAccessLayer);
                result = _PullingTestProvider.GetStandard(BrandID, TestItem, PullForceUnit);
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


        public SendMail_Result FailSendMail(string ReportNo, string ToAddress, string CcAddress)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _PullingTestProvider = new PullingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                System.Data.DataTable dt = _PullingTestProvider.GetData_DataTable(ReportNo);

                string unit = dt.Rows[0]["PullForceUnit"].ToString();
                dt.Columns["PullForceUnit"].ColumnName = unit;
                dt.Rows[0][unit] = dt.Rows[0]["PullForce"].ToString();
                dt.Columns.Remove("PullForce");

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CcAddress,
                    Subject = "Pulling Test - Test Fail",
                    //Body = mailBody,
                    //alternateView = plainView,
                    IsShowAIComment = true,
                    AICommentType = "Pulling test for Snap/Button/Rivet",
                    StyleID = dt.Rows[0]["StyleID"].ToString(),
                    SeasonID = dt.Rows[0]["SeasonID"].ToString(),
                    BrandID = dt.Rows[0]["BrandID"].ToString(),
                };

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, out AlternateView plainView);

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

        public Report_Result GetPDF(string ReportNo)
        {
            Report_Result result = new Report_Result();
            if (string.IsNullOrEmpty(ReportNo))
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            string basefileName = "PullingTest";

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
                PullingTest_Result model = _PullingTestProvider.GetData(ReportNo);

                string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx"; ;
                if (this.IsTest)
                {
                    openfilepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLT", $"{basefileName}.xltx");
                }

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);
                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];

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
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 60, 24);
                }

                cell = worksheet.Cells[15, 1];
                if (model.TestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.TestBeforePicture, model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: this.IsTest);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 240, 100);
                }

                cell = worksheet.Cells[15, 3];
                if (model.TestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.TestAfterPicture, model.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: this.IsTest);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5 , cell.Top + 5, 240, 100);
                }

                #region Save & Show Excel

                string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                string filexlsx = fileName + ".xlsx";
                string fileNamePDF = fileName + ".pdf";

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
