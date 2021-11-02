using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.Authority.Controllers
{
    public class AuthoritybyPositionController : BaseController
    {
        private IAuthorityService _AuthorityService;

        public AuthoritybyPositionController()
        {
            _AuthorityService = new AuthorityService();
            this.SelectedMenu = "Authority";
            ViewBag.OnlineHelp = this.OnlineHelp + "Authority.AuthoritybyPosition,,";
        }
        // GET: Authority/AuthoritybyPosition
        public ActionResult Index()
        {
            this.CheckSession();

            AuthoritybyPosition model = _AuthorityService.Get_Position_Browse(this.FactoryID);

            if (this.UserID == "SCIMIS")
            {
                model = _AuthorityService.Get_Position_Browse(string.Empty);
            }
            return View(model);
        }

        [HttpPost]
        public ActionResult GetDetail(string Position)
        {
            this.CheckSession();

            Quality_Position model = _AuthorityService.Get_Position_Detail(Position);

            return Json(model);
        }

        [HttpPost]
        public ActionResult UpdateDetail(Quality_Position_Request Req)
        {
            this.CheckSession();

            if (string.IsNullOrEmpty(Req.Position))
            {
                return Json(new Quality_Position()
                {
                    Result = false,
                    ErrorMessage = "Position can not be empty."
                });
            }

            Quality_Position model = _AuthorityService.Update_Position_Detail(Req);

            return Json(model);
        }


        [HttpPost]
        public ActionResult CreateDetail(Quality_Position_Request Req)
        {
            this.CheckSession();

            if (string.IsNullOrEmpty(Req.Position))
            {
                return Json(new Quality_Position()
                {
                    Result = false,
                    ErrorMessage = "Position can not be empty."
                });
            }

            if (!ModelState.IsValid)
            {
                return Json(new Quality_Position()
                {
                    Result = false,
                    ErrorMessage = "Data valid false."
                });
            }

            Req.Factory = this.FactoryID;
            Quality_Position model = _AuthorityService.Create_Position_Detail(Req);

            return Json(model);
        }

        [HttpPost]
        public ActionResult UpdatePass2()
        {
            this.CheckSession();

            AuthoritybyPosition model = _AuthorityService.Get_Position_Browse(this.FactoryID);
            AuthoritybyPosition res = _AuthorityService.UpdatePass2();

            if (!res.Result)
            {
                model.Result = false;
                model.ErrorMessage = "Update fail.";
            }

            return Json(model);
        }
    }
}