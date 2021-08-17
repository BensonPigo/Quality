using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ResultModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.Authority.Controllers
{
    public class UserListController : BaseController
    {
        private IAuthorityService _AuthorityService;

        public UserListController()
        {
            _AuthorityService = new AuthorityService();
        }

        public ActionResult Index()
        {
            //this.CheckSession();

            UserList model = _AuthorityService.Get_User_List_Browse();
            return View(model);
        }

        [HttpPost]
        public ActionResult GetDetail(string UserID)
        {
            //this.CheckSession();

            //UserList_Authority model = _AuthorityService.Get_User_List_Detail(UserID);

            //return Json(model);

            return Json("");
        }



        [HttpPost]
        public ActionResult Get_Default_Detail(string Position)
        {
            //this.CheckSession();

            //Quality_Position model = _AuthorityService.Get_Position_Detail(Position);

            //return Json(model);

            return Json("");
        }

        [HttpPost]
        public ActionResult GetUserList()
        {
            //this.CheckSession();

            UserList model = new UserList();

            if (this.UserID == "SCIMIS")
            {
                model = _AuthorityService.GetAlUser(string.Empty);
            }
            else
            {
                model = _AuthorityService.GetAlUser(this.FactoryID);
            }
            return Json(model);
        }


        //[HttpPost]
        //public ActionResult UpdateDetail(UserList_Authority_Request Req)
        //{
        //    //this.CheckSession();

        //    if (!ModelState.IsValid)
        //    {
        //        return Json(new UserList_Authority()
        //        {
        //            Result = false,
        //            ErrorMessage = "Data valid false."
        //        });
        //    }

        //    UserList_Authority model = _AuthorityService.Update_User_List_Detail(Req);

        //    return Json(model);
        //}

        //[HttpPost]
        //public ActionResult ImportUsers(List<UserList_Browse> DataList)
        //{
        //    //this.CheckSession();

        //    UserList_Authority model = _AuthorityService.ImportUsers(DataList);

        //    return Json(model);
        //}
    }
}