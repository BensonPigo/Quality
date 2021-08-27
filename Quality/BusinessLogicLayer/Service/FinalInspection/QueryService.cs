using ADOHelper.Template.MSSQL;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.FinalInspection
{
    public class QueryService : IQueryService
    {
        private IMailToProvider _IMailToProvider;
        private IFinalInspectionProvider _FinalInspectionProvider;

        //寄信
        public BaseResult SendMail(string finalInspectionID, bool isTest)
        {
            BaseResult baseResult = new BaseResult();
            // 取得資料
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            // 寄信設定
            _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                DataRow drReportMailInfo = _FinalInspectionProvider.GetReportMailInfo(finalInspectionID).Rows[0];

                List<MailTo> mailTos = _IMailToProvider.Get(new MailTo() { ID = "401" }).ToList();
                string toAddress = mailTos.Select(s => s.ToAddress).FirstOrDefault();
                string ccAddress = mailTos.Select(s => s.CcAddress).FirstOrDefault();
                string subject = $"Final Inspection Report(PO#: {drReportMailInfo["POID"]})-{drReportMailInfo["InspectionResult"]}";
                StringBuilder content = new StringBuilder();
                content.Append($@"
Hi all,<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This is final inspection report for [PO#]:<font style='color: blue'> {drReportMailInfo["POID"]}</font>. Please refer to below information.<br/>
<b>[Factory]:</b><font style='color: blue'> {drReportMailInfo["FactoryID"]}</font><br/>
<b>[SP#]:</b><font style='color: blue'>  {drReportMailInfo["SP"]}</font><br/>
<b>[Style]:</b><font style='color: blue'>  {drReportMailInfo["StyleID"]}</font><br/>
<b>[Season]:</b><font style='color: blue'>  {drReportMailInfo["SeasonID"]}</font><br/>
<b>[Brand]:</b><font style='color: blue'>  {drReportMailInfo["BrandID"]}</font><br/>
<b>[CFA]:</b><font style='color: blue'>  {drReportMailInfo["CFA"]}</font><br/>
<b>[Submit Date]:</b><font style='color: blue'>  {drReportMailInfo["SubmitDate"]}</font><br/>
<b>[Audit Date]:</b><font style='color: blue'>  {drReportMailInfo["AuditDate"]}</font><br/>
More detail please click here<br/>
<br/>
NOTE: This is an automated reply from a system mailbox. Please do not reply to this email.<br/>
");

                string mailFrom = string.Empty;
                string mailServer = string.Empty;
                string eMailID = string.Empty;
                string eMailPwd = string.Empty;

                string result = string.Empty;
                //寄件者 & 收件者

                SQLParameterCollection objParameter = new SQLParameterCollection();

                DataTable dt = SQLDAL.ExecuteDataTable(CommandType.Text, "select * from Production.dbo.System", objParameter);
                if (dt != null || dt.Rows.Count > 0)
                {
                    mailFrom = dt.Rows[0]["Sendfrom"].ToString();
                    mailServer = dt.Rows[0]["mailServer"].ToString();
                    eMailID = dt.Rows[0]["eMailID"].ToString();
                    eMailPwd = dt.Rows[0]["eMailPwd"].ToString();
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

                foreach (var to in toAddress.Split(';'))
                {
                    if (!string.IsNullOrEmpty(to))
                    {
                        message.To.Add(to);
                    }
                }

                foreach (var cc in ccAddress.Split(';'))
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
                SmtpClient client = new SmtpClient(mailServer);

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

    }
}
