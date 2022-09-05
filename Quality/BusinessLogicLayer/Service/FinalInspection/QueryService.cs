using ADOHelper.Template.MSSQL;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Newtonsoft.Json;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using BusinessLogicLayer.Helper;
using System.Configuration;
using DatabaseObject.ResultModel;
using Microsoft.Office.Interop.Excel;
using Sci;
using System.Runtime.InteropServices;

namespace BusinessLogicLayer.Service.FinalInspection
{
    public class QueryService : IQueryService
    {
        private IMailToProvider _IMailToProvider;
        private IFinalInspectionProvider _FinalInspectionProvider;
        private IFinalInspFromPMSProvider _FinalInspFromPMSProvider;
        private IStyleProvider _StyleProvider;
        private static readonly string CryptoKey = ConfigurationManager.AppSettings["CryptoKey"].ToString();
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];

        public List<QueryFinalInspection> GetFinalinspectionQueryList_Default(QueryFinalInspection_ViewModel request)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetFinalinspectionQueryList_Default(request).ToList();
        }
        public List<QueryFinalInspection> GetFinalinspectionQueryList(QueryFinalInspection_ViewModel request)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetFinalinspectionQueryList(request).ToList();
        }

        public QueryReport GetFinalInspectionReport(string finalInspectionID)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            QueryReport queryReport = new QueryReport();

            try
            {
                System.Data.DataTable dtQueryReportInfo = _FinalInspectionProvider.GetQueryReportInfo(finalInspectionID);

                queryReport.FinalInspection = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                queryReport.SP = dtQueryReportInfo.Rows[0]["SP"].ToString();
                queryReport.StyleID = dtQueryReportInfo.Rows[0]["StyleID"].ToString();
                queryReport.BrandID = dtQueryReportInfo.Rows[0]["BrandID"].ToString();
                queryReport.FinalInspection.CFA = dtQueryReportInfo.Rows[0]["CFA"].ToString();
                queryReport.Dest = dtQueryReportInfo.Rows[0]["Dest"].ToString();
                queryReport.TotalSPQty = (int)dtQueryReportInfo.Rows[0]["TotalSPQty"];

                queryReport.AQLPlan = new FinalInspectionService().GetAQLPlanDesc(queryReport.FinalInspection.AcceptableQualityLevelsUkey);

                _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
                string OrderID = queryReport.SP.IndexOf(",") > 0 ? queryReport.SP.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)[0] : string.Empty;
                queryReport.MeasurementUnit = _StyleProvider.GetSizeUnitByCustPONO(queryReport.FinalInspection.CustPONO, OrderID);

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<FinalInspectionDefectItem> finalInspectionDefectItems = _FinalInspFromPMSProvider.GetFinalInspectionDefectItems(finalInspectionID).ToList();
                if (finalInspectionDefectItems.Any(s => s.Qty > 0))
                {
                    finalInspectionDefectItems = finalInspectionDefectItems.Where(s => s.Qty > 0).ToList();
                }
                else
                {
                    finalInspectionDefectItems = new List<FinalInspectionDefectItem>();
                }

                queryReport.ListDefectItem = finalInspectionDefectItems;

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BACriteriaItem> bACriteriaItems = _FinalInspectionProvider.GetBeautifulProductAuditForInspection(finalInspectionID).ToList();
                if (bACriteriaItems.Any(s => s.Qty > 0))
                {
                    bACriteriaItems = bACriteriaItems.Where(s => s.Qty > 0).ToList();
                }
                else
                {
                    bACriteriaItems = new List<BACriteriaItem>();
                }
                queryReport.ListBACriteriaItem = bACriteriaItems;

                queryReport.ListCartonInfo = _FinalInspectionProvider.GetListCartonInfo(finalInspectionID).ToList();

                queryReport.ListShipModeSeq = _FinalInspectionProvider.GetListShipModeSeq(finalInspectionID).ToList();

                queryReport.ListViewMoistureResult = _FinalInspectionProvider.GetViewMoistureResult(finalInspectionID).ToList();

                queryReport.ListMeasurementViewItem = _FinalInspectionProvider.GetMeasurementViewItem(finalInspectionID).ToList();

                queryReport.GoOnInspectURL = this.GetCurrentAction(queryReport.FinalInspection.InspectionStep);

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                foreach (MeasurementViewItem measurementViewItem in queryReport.ListMeasurementViewItem)
                {
                    System.Data.DataTable dtMeasurementData = _FinalInspectionProvider.GetMeasurement(finalInspectionID, measurementViewItem.Article, measurementViewItem.Size, measurementViewItem.ProductType);
                    measurementViewItem.MeasurementDataByJson = JsonConvert.SerializeObject(dtMeasurementData);
                }
            }
            catch (Exception ex)
            {
                queryReport.Result = false;
                queryReport.ErrorMessage = ex.ToString();
            }

            return queryReport;
        }

        private string GetCurrentAction(string InspectionStep)
        {
            string ActionName = string.Empty;

            switch (InspectionStep)
            {
                case "Setting":
                    ActionName = "Setting";
                    break;
                case "Insp-General":
                    ActionName = "General";
                    break;
                case "Insp-CheckList":
                    ActionName = "CheckList";
                    break;
                case "Insp-AddDefect":
                    ActionName = "AddDefect";
                    break;
                case "Insp-BA":
                    ActionName = "BeautifulProductAudit";
                    break;
                case "Insp-Moisture":
                    ActionName = "Moisture";
                    break;
                case "Insp-Measurement":
                    ActionName = "Measurement";
                    break;
                case "Insp-Others":
                    ActionName = "Others";
                    break;
                case "Submit":
                    ActionName = "Others";
                    break;
                default:
                    break;
            }

            return ActionName;
        }

        //寄信
        public BaseResult SendMail(string finalInspectionID, string WebHost, bool isTest)
        {
            BaseResult baseResult = new BaseResult();
            // 取得資料
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            // 寄信設定
            _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                DataRow drReportMailInfo = _FinalInspectionProvider.GetReportMailInfo(finalInspectionID).Rows[0];

                List<MailTo> mailTos = _IMailToProvider.Get(new MailTo() { ID = "401" }).ToList();
                string toAddress = mailTos.Select(s => s.ToAddress).FirstOrDefault();
                string ccAddress = mailTos.Select(s => s.CcAddress).FirstOrDefault();
                string subject = $"Final Inspection Report(PO#: {drReportMailInfo["CustPONO"]})-{drReportMailInfo["InspectionResult"]}";

                //對照HomeController的RedirectToPage Action裡面的順序設定
                string action = this.GetCurrentAction(finalInspection.InspectionStep);

                string OriInfo = $"FinalInspection+Query+Detail+FinalInspectionID+{finalInspectionID}";

                string code = StringEncryptHelper.AesEncryptBase64(OriInfo, CryptoKey);

                StringBuilder content = new StringBuilder();
                content.Append($@"
Hi all,<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This is final inspection report for [PO#]:<font style='color: blue'> {drReportMailInfo["CustPONO"]}</font>. Please refer to below information.<br/>
<b>[Factory]:</b><font style='color: blue'> {drReportMailInfo["FactoryID"]}</font><br/>
<b>[SP#]:</b><font style='color: blue'>  {drReportMailInfo["SP"]}</font><br/>
<b>[Style]:</b><font style='color: blue'>  {drReportMailInfo["StyleID"]}</font><br/>
<b>[Season]:</b><font style='color: blue'>  {drReportMailInfo["SeasonID"]}</font><br/>
<b>[Brand]:</b><font style='color: blue'>  {drReportMailInfo["BrandID"]}</font><br/>
<b>[CFA]:</b><font style='color: blue'>  {drReportMailInfo["CFA"]}</font><br/>
<b>[Inspection Stage]:</b><font style='color: blue'>  {drReportMailInfo["InspectionStage"]}</font><br/>
<b>[Submit Date]:</b><font style='color: blue'>  {drReportMailInfo["SubmitDate"]}</font><br/>
<b>[Audit Date]:</b><font style='color: blue'>  {drReportMailInfo["AuditDate"]}</font><br/>
<a href='{WebHost}/Home/RedirectToPage?Code={code}'>
More detail please click here
</a>

<br/>
<br/>
NOTE: This is an automated reply from a system mailbox. Please do not reply to this email.<br/>
");

                string mailFrom = string.Empty;
                string mailServer = string.Empty;
                string eMailID = string.Empty;
                string eMailPwd = string.Empty;
                int MailServerPort = 25;

                string result = string.Empty;
                //寄件者 & 收件者

                SQLParameterCollection objParameter = new SQLParameterCollection();
                System.Data.DataTable dt = SQLDAL.ExecuteDataTable(CommandType.Text, "select * from Production.dbo.System", objParameter, ADOHelper.DBToolKit.Common.ProductionDataAccessLayer);
                if (dt != null || dt.Rows.Count > 0)
                {
                    mailFrom = dt.Rows[0]["Sendfrom"].ToString();
                    mailServer = dt.Rows[0]["mailServer"].ToString();
                    eMailID = dt.Rows[0]["eMailID"].ToString();
                    eMailPwd = dt.Rows[0]["eMailPwd"].ToString();
                    MailServerPort = Convert.ToInt32(dt.Rows[0]["MailServerPort"]);
                }

                if (isTest)
                {
                    mailFrom = "foxpro@sportscity.com.tw";
                    mailServer = "Mail.sportscity.com.tw";
                    eMailID = "foxpro";
                    eMailPwd = "orpxof";
                }

                MailMessage message = new MailMessage();
                message.Subject = subject;

                foreach (var to in toAddress.Replace("\r", "").Replace("\n", "").Split(';'))
                {
                    if (!string.IsNullOrEmpty(to))
                    {
                        message.To.Add(to);
                    }
                }

                foreach (var cc in ccAddress.Replace("\r", "").Replace("\n", "").Split(';'))
                {
                    if (!string.IsNullOrEmpty(cc))
                    {
                        message.To.Add(cc);
                    }
                }

                message.From = new MailAddress(mailFrom);
                message.Body = content.ToString();
                message.IsBodyHtml = true;

                // mail Smtp
                SmtpClient client = new SmtpClient(mailServer, MailServerPort);

                // 寄件者 帳密
                client.Credentials = new NetworkCredential(eMailID, eMailPwd);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                client.Send(message);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
                return baseResult;
            }

            return baseResult;
        }

        public Report_Result QueryReport(QueryFinalInspection_ViewModel model)
        {
            Report_Result result = new Report_Result();
            if (model == null)
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            try
            {
                if (!(IsTest.ToLower() == "true"))
                {
                    if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                    {
                        System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                    }

                    if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                    {
                        System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                    }
                }

                List<QueryFinalInspection> finalInspections = new List<QueryFinalInspection>();
                if (string.IsNullOrEmpty(model.SP) &&
                    string.IsNullOrEmpty(model.CustPONO) &&
                    string.IsNullOrEmpty(model.StyleID) &&
                    (!model.AuditDateStart.HasValue || !model.AuditDateEnd.HasValue) &&
                    string.IsNullOrEmpty(model.InspectionResult))
                {
                    finalInspections = GetFinalinspectionQueryList_Default(model);
                }
                else
                {
                    finalInspections = GetFinalinspectionQueryList(model);
                }

                if (finalInspections.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Get Data Fail!";
                    return result;
                }

                string basefileName = "FinalInspectionQuery";
                string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx"; ;
                if (IsTest.ToLower() == "true")
                {
                    openfilepath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLT", $"{basefileName}.xltx");
                }

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);
                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];

                #region 表身資料
                // 塞進資料
                int start_row = 2;
                foreach (var item in finalInspections)
                {
                    worksheet.Cells[start_row, 1] = item.InspectionResult;
                    worksheet.Cells[start_row, 2] = item.SP;
                    worksheet.Cells[start_row, 3] = item.CustPONO;
                    worksheet.Cells[start_row, 4] = item.AuditDate;
                    worksheet.Cells[start_row, 5] = item.SPQty;
                    worksheet.Cells[start_row, 6] = item.StyleID;
                    worksheet.Cells[start_row, 7] = item.Season;
                    worksheet.Cells[start_row, 8] = item.BrandID;
                    worksheet.Cells[start_row, 9] = item.Article;
                    worksheet.Cells[start_row, 10] = item.InspectionTimes;
                    worksheet.Cells[start_row, 11] = item.InspectionStage;
                    worksheet.Cells[start_row, 12] = item.IsTransferToPMS;
                    worksheet.Cells[start_row, 13] = item.IsTransferToPivot88;
                    worksheet.Rows[start_row].Font.Bold = false;
                    worksheet.Rows[start_row].WrapText = true;
                    start_row++;
                }
                worksheet.Columns.AutoFit();
                #endregion

                string fileName = $"FinalInspectionQueryReport{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                string filexlsx = fileName + ".xlsx";
                string filepath = System.IO.Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);

                Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                result.TempFileName = filexlsx;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
                result.Result = false;
            }
            return result;
        }

        public BaseResult ClickJunk(string ID)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            return _FinalInspectionProvider.UpdateJunk(ID);
        }
    }
}
