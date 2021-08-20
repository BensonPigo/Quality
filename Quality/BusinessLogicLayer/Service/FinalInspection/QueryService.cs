using ADOHelper.Template.MSSQL;
using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
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
    public class QueryService: IQueryService
    {
        private IOrdersProvider _IOrdersProvider;
        private IMailToProvider _IMailToProvider;
        //寄信
        public bool SendMail (DatabaseObject.ManufacturingExecutionDB.FinalInspection Req)
        {
            bool boolTest = true;
            //透過Req從後端取得資料
            QueryReport data = new QueryReport();
            if (boolTest)
            {
                data.SP = "21060448HH003";
                data.StyleID = "22256";
                data.BrandID = "REI";
                data.FinalInspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection() { POID = "21060448HH", FactoryID = "ES2", CFA = "AAA", SubmitDate = DateTime.Now, AuditDate = DateTime.Now, InspectionResult = "Pass" };
            }

            // GetForFinalInspection 取得 SeasonID
            _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            FinalInspection_Request requestItem = new FinalInspection_Request() { SP = data.SP, POID = data.FinalInspection.POID, FactoryID = data.FinalInspection.FactoryID, StyleID = data.StyleID };
            IList<Orders> orders = _IOrdersProvider.GetForFinalInspection(requestItem);

            string submitDate = data.FinalInspection.SubmitDate.HasValue ? ((DateTime)data.FinalInspection.SubmitDate).ToString("yyyy/MM/dd") : string.Empty;
            string auditDate = data.FinalInspection.AuditDate.HasValue ? ((DateTime)data.FinalInspection.AuditDate).ToString("yyyy/MM/dd") : string.Empty;

            // 寄信設定
            _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<MailTo> mailTos = _IMailToProvider.Get(new MailTo() { ID = "401" }).ToList();
            string toAddress = mailTos.Select(s => s.ToAddress).FirstOrDefault();
            string ccAddress = mailTos.Select(s => s.CcAddress).FirstOrDefault();
            string subject = $"Final Inspection Report(PO#: {data.FinalInspection.POID})-{data.FinalInspection.InspectionResult}";
            StringBuilder content = new StringBuilder();
            content.Append($@"
Hi all,<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This is final inspection report for [PO#]:<font style='color: blue'> {data.FinalInspection.POID}</font>. Please refer to below information.<br/>
<b>[Factory]:</b><font style='color: blue'> {data.FinalInspection.FactoryID}</font><br/>
<b>[SP#]:</b><font style='color: blue'>  {data.SP}</font><br/>
<b>[Style]:</b><font style='color: blue'>  {data.StyleID}</font><br/>
<b>[Season]:</b><font style='color: blue'>  {orders.Select(s => s.SeasonID).FirstOrDefault()}</font><br/>
<b>[Brand]:</b><font style='color: blue'>  {data.BrandID}</font><br/>
<b>[CFA]:</b><font style='color: blue'>  {data.FinalInspection.CFA}</font><br/>
<b>[Submit Date]:</b><font style='color: blue'>  {submitDate}</font><br/>
<b>[Audit Date]:</b><font style='color: blue'>  {auditDate}</font><br/>
More detail please click here<br/>
<br/>
NOTE: This is an automated reply from a system mailbox. Please do not reply to this email.<br/>
");

            string mailFrom = "foxpro@sportscity.com.tw";
            string mailServer = "Mail.sportscity.com.tw";
            string eMailID = "foxpro";
            string eMailPwd = "orpxof";

            string result = string.Empty;
            //寄件者 & 收件者

            if (!boolTest)
            {
                //ExecuteDataTable()
                SQLParameterCollection objParameter = new SQLParameterCollection();

                DataTable dt = SQLDAL.ExecuteDataTable(CommandType.Text, "select * from Production.dbo.System", objParameter);
                if (dt != null || dt.Rows.Count > 0)
                {
                    mailFrom = dt.Rows[0]["Sendfrom"].ToString();
                    mailServer = dt.Rows[0]["mailServer"].ToString();
                    eMailID = dt.Rows[0]["eMailID"].ToString();
                    eMailPwd = dt.Rows[0]["eMailPwd"].ToString();
                }
            }
            try
            {
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
                return false;
            }

            return true;
        }
        
    }
}
