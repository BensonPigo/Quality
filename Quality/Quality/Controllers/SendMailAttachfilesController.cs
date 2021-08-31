using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Controllers
{
    public class SendMailAttachfilesController : BaseController
    {
        private ISendMailService _SendMailService;

        public SendMailAttachfilesController()
        {
            _SendMailService = new SendMailService();

        }

        public ActionResult SendMailer(string TO, string CC)
        {
            SendMail_Request request = new SendMail_Request() 
            {
                To = TO,
                CC = CC,
            };

            ViewBag.JS = "";

            return View(request);
        }

        [HttpPost]
        public ActionResult SendMailer(SendMail_Request _Request)
        {
            SendMail_Result result = _SendMailService.SendMail(_Request);

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
