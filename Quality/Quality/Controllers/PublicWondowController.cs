using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Controllers
{
    public class PublicWondowController : BaseController
    {
        private PublicWindowService _PublicWindowService;

        public PublicWondowController()
        {
            _PublicWindowService = new PublicWindowService();
        }


        // GET: PublicWondow
        public ActionResult BrandList()
        {
            var model = _PublicWindowService.Get_Brand(null);
            return View(model);
        }

        [HttpPost]
        public ActionResult BrandList(string BrandID)
        {
            var model = _PublicWindowService.Get_Brand(BrandID);
            return View(model);
        }

        public ActionResult SeasonList(string BrandID)
        {
            var model = _PublicWindowService.Get_Season(BrandID, string.Empty);
            ViewData["BrandID"] = BrandID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SeasonList(string BrandID, string SeasonID)
        {
            var model = _PublicWindowService.Get_Season(BrandID, SeasonID);
            ViewData["BrandID"] = BrandID;
            return View(model);
        }

        public ActionResult StyleList(string BrandID, string SeasonID)
        {
            var model = _PublicWindowService.Get_Style(BrandID, SeasonID, string.Empty);
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            return View(model);
        }

        [HttpPost]
        public ActionResult StyleList(string BrandID, string SeasonID, string StyleID)
        {
            var model = _PublicWindowService.Get_Style(BrandID, SeasonID, StyleID);
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            return View(model);
        }

        public ActionResult ArticleList(string OrderID, Int64 StyleUkey)
        {
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, string.Empty);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            return View(model);
        }

        [HttpPost]
        public ActionResult ArticleList(string OrderID, Int64 StyleUkey, string Article)
        {
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, Article);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            return View(model);
        }

        public ActionResult SizeList(string OrderID, Int64 StyleUkey, string Article)
        {
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, Article, string.Empty);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            return View(model);
        }

        [HttpPost]
        public ActionResult SizeList(string OrderID, Int64 StyleUkey, string Article, string Size)
        {
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, Article, Size);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            return View(model);
        }

        public ActionResult TechnicianList(string CallFunction, string Region)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, Region, string.Empty);
            ViewData["Region"] = Region;
            ViewData["CallFunction"] = CallFunction;
            return View(model);
        }

        [HttpPost]
        public ActionResult TechnicianList(string CallFunction, string Region, string ID)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, Region, ID);
            ViewData["Region"] = Region;
            ViewData["CallFunction"] = CallFunction;
            return View(model);
        }

        public ActionResult SewingLineList(string Region)
        {
            var model = _PublicWindowService.Get_SewingLine(Region, string.Empty);
            ViewData["Region"] = Region;
            return View(model);
        }

        [HttpPost]
        public ActionResult SewingLineList(string Region, string ID)
        {
            var model = _PublicWindowService.Get_SewingLine(Region, ID);
            ViewData["Region"] = Region;
            return View(model);
        }

    }
}