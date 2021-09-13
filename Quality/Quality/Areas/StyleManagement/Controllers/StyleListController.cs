using BusinessLogicLayer.Interface.StyleManagement;
using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.StyleManagement.Controllers
{
    public class StyleListController : BaseController
    {
        private IStyleListService _StyleListService;

        public StyleListController()
        {
            _StyleListService = new StyleListService();
            ViewBag.OnlineHelp = this.OnlineHelp + "StyleManagement.StyleList,,";
        }

        // GET: StyleManagement/StyleList
        public ActionResult Index()
        {
            this.CheckSession();
            StyleList model = new StyleList() { DataList = new List<StyleList>() };
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(StyleList_Request Req)
        {
            this.CheckSession();

            StyleList model = new StyleList() { DataList = new List<StyleList>() };
            if (Req == null || (string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID)))
            {
                model.MsgScript = $@"
msg.WithInfo('Style, Brand and Season cannot be empty.');
";
                return View("Index", model);
            }

            model = _StyleListService.Get_StyleInfo(Req);

            if (!model.Result)
            {
                model.MsgScript = $@"
msg.WithInfo('{model.ErrorMessage.Replace("\r\n", "<br />")}');
";
            }

            return View("Index", model);
        }
    }
}