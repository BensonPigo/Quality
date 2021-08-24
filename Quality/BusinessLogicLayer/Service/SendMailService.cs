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

                MailMessage message = new MailMessage();
                message.From = new MailAddress(SendMail_Request.From);

                foreach (var to in SendMail_Request.To.Split(';'))
                {
                    if (!string.IsNullOrEmpty(to))
                    {
                        message.To.Add(to);
                    }
                }

                foreach (var cc in SendMail_Request.CC.Split(';'))
                {
                    if (!string.IsNullOrEmpty(cc))
                    {
                        message.To.Add(cc);
                    }
                }

                message.Subject = SendMail_Request.Subject;
                message.Body = SendMail_Request.Description;

                // mail Smtp
                SmtpClient client = new SmtpClient(system[0].Mailserver);

                // 寄件者 帳密
                client.Credentials = new NetworkCredential(system[0].EmailID, system[0].EmailPwd);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                // 夾檔
                foreach (var fileName in SendMail_Request.FileList)
                {
                    if (!fileName.Equals(string.Empty))
                    {
                        string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
                        Attachment attachFile = new Attachment(filepath);
                        message.Attachments.Add(attachFile);
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

            DeleteFileonServer(SendMail_Request.FileList);
            return sendMail_Result;
        }

        private void DeleteFileonServer(List<string> Files)
        {
            foreach (string fileName in Files)
            {
                string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
                if (File.Exists(filepath))
                {
                    File.Delete(filepath);
                }
            }
        }
    }
}
