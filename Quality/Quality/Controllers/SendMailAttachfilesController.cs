using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;

namespace Quality.Controllers
{
    public class SendMailAttachfilesController : BaseController
    {
        private MailTools _SendMail;
        private MailToolsService _Service;

        public SendMailAttachfilesController()
        {
            _SendMail = new MailTools();
            _Service = new MailToolsService();
        }

        public ActionResult SendMailer(string TO, string CC, string subject, string body, string file)
        {
            SendMail_Request request = new SendMail_Request() 
            {
                To = TO,
                CC = CC,
                Subject = subject,
                Body = body,
                FileonServer = new List<string>() { file }
            };

            ViewBag.JS = "";

            return View(request);
        }

        [HttpPost]
        public ActionResult SendMailer(SendMail_Request _Request)
        {
            if (_Request.IsShowAIComment)
            {
                string comment = _Service.GetAICommet(_Request);
                string buyReadyDate = _Service.GetBuyReadyDate(_Request);
                _Request.Body = _Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;
            }

            SendMail_Result result = MailTools.SendMail(_Request);

            string js = "";
           
            js += "<script src='/ThirdParty/SciCustom/js/jquery-3.4.1.min.js'></script> ";
            js += "<script>  $(function () { ";
            if (result.result)
            {
                js += "alert('Success'); ";
            }
            else
            {
                js += "alert('" + result.resultMsg + "'); ";
            }

            js += "window.close(); ";
            js += " }); </script>";

            return Content(js, "text/html");
        }
    }
}
