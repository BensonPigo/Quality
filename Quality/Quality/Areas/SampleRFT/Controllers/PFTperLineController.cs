using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ViewModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class PFTperLineController : BaseController
    {
        private IRFTPerLineService _RFTPerLineService;

        public PFTperLineController()
        {
            _RFTPerLineService = new RFTPerLineService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.PFTperLine,,";
        }

        // GET: SampleRFT/PFTperLine
        public ActionResult Index()
        {
            RFTPerLine_ViewModel rftPerLine = _RFTPerLineService.GetQueryPara();

            RFTPerLine_ViewModel rftPerLineQuery = _RFTPerLineService.RFTPerLineQuery(this.FactoryID, rftPerLine.Years.FirstOrDefault(), rftPerLine.Months.FirstOrDefault().Key.ToString());

            return View(rftPerLine);
        }
    }
}