using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
        [SessionAuthorizeAttribute]
        public ActionResult Index()
        {
            List<Quality_MailGroup> quality_MailGroups = _BulkFGTMailGroup_Service.MailGroupGet(new Quality_MailGroup() { Type = "BulkFGT" }).ToList();
            List<SelectListItem> FactoryList = new SetListItem().ItemListBinding(this.Factorys);
            ViewBag.FactoryList = FactoryList; 
            return View(quality_MailGroups);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetDetail(string Factory, string GroupName)
        {
            if (string.IsNullOrEmpty(Factory) && string.IsNullOrEmpty(GroupName))
            {
                return Json(new Quality_MailGroup());
            }

            Quality_MailGroup quality_MailGroups = _BulkFGTMailGroup_Service.MailGroupGet(new Quality_MailGroup() { Type = "BulkFGT", FactoryID = Factory, GroupName = GroupName }).ToList().FirstOrDefault();
            return Json(quality_MailGroups);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult SaveDetail(Quality_MailGroup quality_Mail, string Action)
        {
            quality_Mail.Type = "BulkFGT";
            Quality_MailGroup_ResultModel result = new Quality_MailGroup_ResultModel();

            List<string> to = quality_Mail.ToAddress != null && quality_Mail.ToAddress.Any() ? quality_Mail.ToAddress.Split(';').Where(o=>!string.IsNullOrEmpty(o)).ToList() : new List<string>();
            List<string> cc = quality_Mail.CcAddress != null && quality_Mail.CcAddress.Any() ? quality_Mail.CcAddress.Split(';').Where(o => !string.IsNullOrEmpty(o)).ToList() : new List<string>(); 
            bool allResult = true;
            foreach (var item in to)
            {
                bool OK = Regex.IsMatch(item, @"^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$");
                if (!OK)
                {
                    allResult = false;
                    break;
                }
            }

            if (allResult)
            {
                foreach (var item in cc)
                {
                    bool OK = Regex.IsMatch(item, @"^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$");
                    if (!OK)
                    {
                        allResult = false;
                        break;
                    }
                }
            }

            if (!allResult)
            {
                result.Result = false;
                result.ErrMsg = "Format is not correct";
            }

            if (allResult)
            {
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
            }

            return Json(new { result.Result , result.ErrMsg });
        }
    }
}