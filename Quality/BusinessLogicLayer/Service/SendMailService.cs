using BusinessLogicLayer.Interface;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace BusinessLogicLayer.Service
{
    public class SendMailService : ISendMailService
    {
        private ISystemProvider _SystemProvider;
        public SendMail_Result SendMail(SendMail_Request SendMail_Request)
        {
            SendMail_Result sendMail_Result = new SendMail_Result();
            try
            {
                _SystemProvider = new SystemProvider(Common.ProductionDataAccessLayer);
                var system = _SystemProvider.Get();
                if (system == null || system.Count == 0)
                {
                    sendMail_Result.result = false;
                    sendMail_Result.resultMsg = "Get system datas fail!";
                    return sendMail_Result; ;
                }

                MailMessage message = new MailMessage();
                message.From = new MailAddress(system[0].Sendfrom);

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
                SmtpClient client = new SmtpClient(system[0].Mailserver);

                // 寄件者 帳密
                client.Credentials = new NetworkCredential(system[0].EmailID, system[0].EmailPwd);
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

    }
}
