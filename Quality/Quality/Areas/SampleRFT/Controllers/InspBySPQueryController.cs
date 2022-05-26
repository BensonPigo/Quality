using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspBySPQueryController : BaseController
    {
        private InspectionBySPService _Service;
        public InspBySPQueryController()
        {
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.InspBySPQuery,,";
            _Service = new InspectionBySPService();
        }
        // GET: SampleRFT/InspBySPQuery
        public ActionResult Index()
        {
            List<string> inspectionlist = new List<string>() {
                "", "Pass","Fail","On-going"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;

            QueryInspectionBySP_ViewModel model = new QueryInspectionBySP_ViewModel() { DataList = new List<QueryInspectionBySP>()};
            //model.DataList = _Service.GetFinalinspectionQueryList_Default(model);

            return View(model);
        }
        public ActionResult Detail(long ID)
        {
            return View();
        }
    }
}