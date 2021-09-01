using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
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
            List<Quality_MailGroup> quality_MailGroups = _BulkFGTMailGroup_Service.MailGroupGet(new Quality_MailGroup() { Type = "BulkFGT" }).ToList();

            Quality_MailGroup quality_Mail = new Quality_MailGroup()
            {
                FactoryID = "ESP",
                Type = "BulkFGT",
                GroupName = "ttttTest",
                ToAddress = "jack.hsu@sportscity.com.tw",
                CcAddress = "",
            };

            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            ViewBag.FactoryList = FactoryList;

            _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Insert);


            quality_Mail.ToAddress = "jack.hsujack.hsu@sportscity.com.tw";
            _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Update);


            _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, BulkFGTMailGroup_Service.SaveType.Delete);
            quality_MailGroups.Add(quality_Mail);
            return View(quality_MailGroups);
        }


        [HttpPost]
        public ActionResult GetDetail(string Factory, string GroupName)
        {
            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            ViewBag.FactoryList = FactoryList;
            Quality_MailGroup quality_Mail = new Quality_MailGroup();

            if (!string.IsNullOrEmpty(Factory))
            {
                quality_Mail.FactoryID = "ESP";
                quality_Mail.Type = "BulkFGT";
                quality_Mail.GroupName = "ttttTest";
                quality_Mail.ToAddress = "jack.hsu@sportscity.com.tw";
                quality_Mail.CcAddress = "";
            }

            return Json(quality_Mail);
        }
    }
}