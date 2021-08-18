using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ResultModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace Quality.Areas.SampleRFT.Controllers
{
    public class BACriteriaController : BaseController
    {
        // GET: SampleRFT/BACriteria
        public ActionResult Index()
        {
            return View();
        }
    }
}