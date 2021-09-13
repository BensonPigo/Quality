using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Controllers
{
    public class HomeController : BaseController
    {
        private readonly ILoginService _LoginService;
        private string WebPortalURL = ConfigurationManager.AppSettings["WebPortalURL"];
        private static readonly string CryptoKey = ConfigurationManager.AppSettings["CryptoKey"].ToString();

        public HomeController()
        {
            _LoginService = new LoginService();
            if (this.Factorys == null || this.Factorys.Count() == 0)
                this.Factorys = _LoginService.GetFactory();
        }

        public ActionResult Index()
        {
            //Seession失效導向Login頁
            if (Session.Keys.Count <= 1)
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

            List<SelectListItem> FactoryList = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(this.Factorys);
            ViewData["Factorys"] = FactoryList;

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
                this.MDivisionID = result.MDivisionID;
                this.FactoryID = result.FactoryID;
                this.Lines = result.Lines;
                this.Line = this.Lines.FirstOrDefault();
                this.Brands = result.Brands;
                this.Brand = this.Brands.Where(x => x.Equals("ADIDAS")).Select(x => x).FirstOrDefault();

                if (!string.IsNullOrEmpty(this.TargetArea) && !string.IsNullOrEmpty(this.TargetController) && !string.IsNullOrEmpty(this.TargetAction) 
                    && !string.IsNullOrEmpty(this.TargetPKey_Parameter) && !string.IsNullOrEmpty(this.TargetPKey_Value))
                {
                    string area = this.TargetArea;
                    string controller = this.TargetController;
                    string action = this.TargetAction;
                    string Parameter = this.TargetPKey_Parameter;
                    string value = this.TargetPKey_Value;

                    this.TargetArea = string.Empty;
                    this.TargetController = string.Empty;
                    this.TargetAction = string.Empty;
                    this.TargetPKey_Parameter = string.Empty;
                    this.TargetPKey_Value = string.Empty;

                    return RedirectToAction(action, controller, new { Area = area, FinalInspectionID = value });
                }

                return RedirectToAction("Index");
            }
            else
            {
                List<SelectListItem> FactoryList = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(this.Factorys);
                ViewData["Factorys"] = FactoryList;
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
            ViewData["Factorys"] = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(this.Factorys);
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


        public ActionResult Setting()
        {
            List<SelectListItem> BrandList = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(this.Brands);

            ViewData["Brand"] = this.BulkFGT_Brand;
            ViewData["Brands"] = BrandList;
            ViewData["Msg"] = null;

            return View();
        }

        [HttpPost]
        public ActionResult Setting(string BrandID)
        {
            Quality_Pass1_Request Req = new Quality_Pass1_Request()
            {
                ID = this.UserID,
                SampleTesting_Brand = BrandID
            };

            LogIn_Result info = _LoginService.Update_Pass1(Req);
            if (info.Result)
            {
                List<SelectListItem> BrandList = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(this.Brands);
                this.BulkFGT_Brand = BrandID;
                ViewData["Brand"] = this.BulkFGT_Brand;
                ViewData["Brands"] = BrandList;
                ViewData["Msg"] = "Success!";
                return View();
            }
            else
            {
                ViewData["Msg"] = info.ErrorMessage;
                return View("Login");
            }
        }


        public ActionResult RedirectToPage(string Code)
        {
            string OriInfo = StringEncryptHelper.AesDecryptBase64(Code, CryptoKey);

            this.TargetArea = OriInfo.Split('+')[0];
            this.TargetController = OriInfo.Split('+')[1];
            this.TargetAction = OriInfo.Split('+')[2];
            this.TargetPKey_Parameter = OriInfo.Split('+')[3];
            this.TargetPKey_Value = OriInfo.Split('+')[4];

            return RedirectToAction("Login");
        }
    }
}