using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class PullingTestController : BaseController
    {
        // GET: BulkFGT/PullingTest
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(PullingTest_ViewModel Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(PullingTest_ViewModel Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(PullingTest_ViewModel Req)
        {
            return View();
        }
    }
}