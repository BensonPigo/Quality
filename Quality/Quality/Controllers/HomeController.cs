using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Controllers
{
    public class HomeController : BaseController
    {
        private readonly ILoginService _LoginService;
        public HomeController()
        {
            _LoginService = new LoginService();
        }

        public ActionResult Index()
        {            
            //Seession失效導向Login頁
            if (Session.Keys.Count == 0)
            {
                return RedirectToAction("Login");
            }

            return View();
        }


        public ActionResult Login()
        {
            if (TempData["Msg"] != null)
            {
                ViewData["Msg"] = TempData["Msg"].ToString();
                TempData.Clear();
                ClearUsersInfo();
                return View("Login");
            }
            return View();
        }

        [HttpPost]
        public ActionResult Login(LogIn_Request LoginUser)
        {
            string agent = Request.Headers["User-Agent"];

            if (Request.Browser.Browser == "InternetExplorer")
            {
                ViewData["Msg"] = "Internet Explorer is not supported, please use Chrome or FireFox.";
                return View("Login");
            }

            LogIn_Result result = _LoginService.LoginValidate(LoginUser);

            if (result.Result)
            {
                this.UserID = result.pass1.ID;
                this.BulkFGT_Brand = result.pass1.BulkFGT_Brand;
                this.MenuList = result.Menus;
                this.Factorys = result.Factorys;
                this.FactoryID = this.Factorys.FirstOrDefault();
                this.Lines = result.Lines;
                this.Line = this.Lines.FirstOrDefault();
                this.Brands = result.Brands;
                this.Brand = this.Brands.Where(x => x.Equals("ADIDAS")).Select(x => x).FirstOrDefault();
                return RedirectToAction("Index");
            }
            else
            {
                ViewData["Msg"] = result.ErrorMessage;
                return View("Login");
            }
        }

        /// <summary>
        /// 登出
        /// </summary>
        /// <returns></returns>
        public ActionResult Logout()
        {
            ClearUsersInfo();
            return View("Login");

        }

        /// <summary>
        /// 清除Seesion
        /// </summary>
        private void ClearUsersInfo()
        {
            Session.Abandon();

            string[] cookieList = { "ASP.NET_SessionId", "Nav", "Fun" };

            for (int i = 0; i < cookieList.Length; i++)
            {
                Response.Cookies.Add(new HttpCookie(cookieList[i], ""));

                HttpCookie httpCookie = new HttpCookie(cookieList[i])
                {
                    Expires = DateTime.Now.AddDays(-1)
                };
                httpCookie.Values.Clear();

                Response.Cookies.Set(httpCookie);
            }
        }
    }
}