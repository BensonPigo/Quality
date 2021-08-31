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

        public string DataTableChangeHtml(DataTable dt)
        {
            string html = "<html> ";

            html += @"
<style>
    .DataTable {
        width: 92vw;
        font-size: 1rem;
        font-weight: bold;
        border: solid 1px black;
        background-color: white;
    }
        .DataTable > tbody > tr:nth-of-type(odd) {
            background-color: #ffffff;
        }

        .DataTable > tbody > tr:nth-of-type(even) {
            background-color: #F0F2F2;
        }

        .DataTable > tbody > tr > td {
            border: solid 1px gray;
            padding: 1em;
            text-align: left;
            vertical-align: middle;
        }
</style>
";

            html += "<body> ";
            html += "<table class='DataTable'> ";
            html += "<thead><tr> ";
            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                html += "<th>" + dt.Columns[i].ColumnName + "</th> ";
            }
            html += "</tr></thead> ";
            html += "<tbody> ";
            for (int i = 0; i <= dt.Rows.Count - 1 ; i++)
            {
                html += "<tr> ";
                for (int j = 0; j <= dt.Columns.Count -1; j++)
                {
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td> ";
                }
                html += "</tr> ";
            }
            html += "</tbody> ";
            html += "</table> ";
            html += "</body> ";
            html += "</html> ";
            return html;
        }
    }
}
