using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.ViewModel.SampleRFT;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspectionDefectListController : BaseController
    {
        private InspectionDefectListService _Service;

        public InspectionDefectListController()
        {
            _Service = new InspectionDefectListService();
        }

        // GET: SampleRFT/InspectionDefectList
        public ActionResult Index()
        {
            InspectionDefectList_ViewModel model = new InspectionDefectList_ViewModel()
            {
                DataList = new List<InspectionDefectList_Result>()
            };
            return View(model);
        }

        public ActionResult IndexGet(string OrderID)
        {
            InspectionDefectList_ViewModel model = _Service.GetData(new InspectionDefectList_ViewModel() { OrderID = OrderID });

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult Query(InspectionDefectList_ViewModel Req)
        {
            InspectionDefectList_ViewModel model = _Service.GetData(Req);

            return View("Index", model);
        }
    }
}