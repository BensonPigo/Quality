using ADOHelper.Template.MSSQL;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace ADOHelper.Utility
{
    public class MailTools
    {
        private static string IsTest = ConfigurationManager.AppSettings["IsTest"];
        public static SendMail_Result SendMail(SendMail_Request SendMail_Request)
        {
            SendMail_Result sendMail_Result = new SendMail_Result();
            try
            {
                bool isTest = IsTest.ToLower() == "true";
                SQLParameterCollection objParameter = new SQLParameterCollection();

                string mailFrom = "foxpro@sportscity.com.tw";
                string mailServer = "Mail.sportscity.com.tw";
                string EmailID = "foxpro";
                string EmailPwd = "orpxof";

                string result = string.Empty;
                //寄件者 & 收件者

                if (!isTest)
                {
                    //ExecuteDataTable()
                    DataTable dt = SQLDAL.ExecuteDataTable(CommandType.Text, "select * from Production.dbo.System", new SQLParameterCollection());
                    if (dt != null || dt.Rows.Count > 0)
                    {
                        mailFrom = dt.Rows[0]["Sendfrom"].ToString();
                        mailServer = dt.Rows[0]["mailServer"].ToString();
                        EmailID = dt.Rows[0]["eMailID"].ToString();
                        EmailPwd = dt.Rows[0]["eMailPwd"].ToString();
                    }
                    else
                    {
                        sendMail_Result.result = false;
                        sendMail_Result.resultMsg = "Get system datas fail!";
                        return sendMail_Result; ;
                    }
                }

                MailMessage message = new MailMessage();
                message.From = new MailAddress(mailFrom);

                if (SendMail_Request.To != null && SendMail_Request.To != string.Empty)
                {
                    foreach (var to in SendMail_Request.To.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(to))
                        {
                            message.To.Add(to);
                        }
                    }
                }

                if (SendMail_Request.CC != null && SendMail_Request.CC != string.Empty)
                {
                    foreach (var cc in SendMail_Request.CC.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(cc))
                        {
                            message.CC.Add(cc);
                        }
                    }
                }

                if (SendMail_Request.Subject != null)
                {
                    message.Subject = SendMail_Request.Subject;
                }

                message.IsBodyHtml = true;
                if (SendMail_Request.Body != null)
                {
                    message.Body = SendMail_Request.Body;
                }

                // mail Smtp
                SmtpClient client = new SmtpClient(mailServer);

                // 寄件者 帳密
                client.Credentials = new NetworkCredential(EmailID, EmailPwd);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                // 夾檔
                if (SendMail_Request.FileUploader != null && SendMail_Request.FileUploader.Count > 0)
                {
                    foreach (var fileUploader in SendMail_Request.FileUploader)
                    {
                        if (fileUploader != null)
                        {
                            string fileName = Path.GetFileName(fileUploader.FileName);
                            message.Attachments.Add(new Attachment(fileUploader.InputStream, fileName));
                        }
                    }
                }


                // 夾檔(在server上的檔案)
                if (SendMail_Request.FileonServer != null && SendMail_Request.FileonServer.Count > 0)
                {
                    foreach (string file in SendMail_Request.FileonServer)
                    {
                        if (!string.IsNullOrEmpty(file))
                        {
                            // 找到server 上的檔案
                            string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", file);
                            message.Attachments.Add(new Attachment(filepath));
                        }
                    }
                }

                client.Send(message);

                sendMail_Result.result = true;
                sendMail_Result.resultMsg = string.Empty;
            }
            catch (System.Exception ex)
            {
                sendMail_Result.result = false;
                sendMail_Result.resultMsg = ex.ToString();
            }

            return sendMail_Result;
        }

        public static string DataTableChangeHtml(DataTable dt)
        {
            string html = "<html> ";

            html += @"
<style>
    .tg {border-collapse:collapse;border-spacing:0;}
.tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
.tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
        }
</style>
";

            html += "<body> ";
            html += "<table class='tg'> ";
            html += "<thead><tr bgcolor=\"#FFDEA1\" > ";
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
