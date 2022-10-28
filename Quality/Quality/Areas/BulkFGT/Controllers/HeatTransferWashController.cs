using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class HeatTransferWashController : BaseController
    {
        private HeatTransferWashService _Service;
        public HeatTransferWashController()
        {
            _Service = new HeatTransferWashService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.HeatTransferWash,,";
        }
        // GET: BulkFGT/HeatTransferWash
        public ActionResult Index()
        {
            HeatTransferWash_ViewModel mode = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            return View(mode);
        }
    }
}