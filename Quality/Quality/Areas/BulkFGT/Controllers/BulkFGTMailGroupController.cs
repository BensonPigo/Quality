using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class BulkFGTMailGroupController : BaseController
    {
        private IBulkFGTMailGroup_Service _BulkFGTMailGroup_Service;

        public BulkFGTMailGroupController()
        {
            _BulkFGTMailGroup_Service = new BulkFGTMailGroup_Service();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.BulkFGTMailGroup,,";
        }

        // GET: BulkFGT/BulkFGTMailGroup
        public ActionResult Index()
        {
            List<Quality_MailGroup> quality_MailGroups = _BulkFGTMailGroup_Service.MailGroupGet(new Quality_MailGroup() { Type = "BulkFGT", FactoryID = this.FactoryID }).ToList();
            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            ViewBag.FactoryList = FactoryList; 
            return View(quality_MailGroups);
        }

        [HttpPost]
        public ActionResult GetDetail(string Factory, string GroupName)
        {
            Quality_MailGroup quality_MailGroups = _BulkFGTMailGroup_Service.MailGroupGet(new Quality_MailGroup() { Type = "BulkFGT", FactoryID = Factory, GroupName = GroupName }).ToList().FirstOrDefault();
            return Json(quality_MailGroups);
        }

        [HttpPost]
        public ActionResult SaveDetail(Quality_MailGroup quality_Mail, string Action)
        {
            quality_Mail.Type = "BulkFGT";
            Quality_MailGroup_ResultModel result = new Quality_MailGroup_ResultModel();
            switch (Action)
            {
                case "Create":
                    result = _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Insert);
                    break;
                case "Update":
                    result = _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Update);
                    break;
                case "Delete":
                    result = _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Delete);
                    break;
            }

            return Json(new { result.Result , result.ErrMsg });
        }
    }
}