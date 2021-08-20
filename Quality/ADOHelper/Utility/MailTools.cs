using ADOHelper.Template.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ADOHelper.Utility
{
    public class MailTools
    {
        public static string MailToHtml(string mailTO, string subject, string attach, string desc)
        {
            bool boolTest = true;
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
            // MailAddress from = new MailAddress(mailFrom);
            // MailAddress to = new MailAddress(mailTo);
            // MailAddress cc = new MailAddress(mailCC);
            MailMessage message = new MailMessage();
            message.Subject = subject;
            message.From = new MailAddress(mailFrom);

            string[] mails = mailTO.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string mail in mails)
            {
                message.To.Add(mail);
            }

            if (message.To.Count == 0)
            {
                return "Mail Address is empty";
            }

            message.Body = desc;
            message.IsBodyHtml = true;
            //Gmail Smtp
            SmtpClient client = new SmtpClient(mailServer);
            //寄件者 帳密
            client.Credentials = new NetworkCredential(eMailID, eMailPwd);

            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            //夾檔
            string filePath = attach;
            if (!filePath.Equals(string.Empty))
            {
                Attachment attachFile = new Attachment(filePath);
                message.Attachments.Add(attachFile);
            }
            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString();
            }

            return result;
        }
    }
}
