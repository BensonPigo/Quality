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

        public ActionResult TechnicianList(string CallFunction)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, string.Empty);
            ViewData["CallFunction"] = CallFunction;
            return View(model);
        }

        [HttpPost]
        public ActionResult TechnicianList(string CallFunction, string ID)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, ID);
            ViewData["CallFunction"] = CallFunction;
            return View(model);
        }


        public ActionResult LocalSuppList(string Title)
        {
            var model = _PublicWindowService.Get_LocalSupp( string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult LocalSuppList(string Title, string Name)
        {
            var model = _PublicWindowService.Get_LocalSupp(Name);
            ViewData["Title"] = Title;
            return View(model);
        }

        public ActionResult TPESuppList(string Title)
        {
            var model = _PublicWindowService.Get_TPESupp(string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult TPESuppList(string Title, string Name)
        {
            var model = _PublicWindowService.Get_TPESupp(Name);
            ViewData["Title"] = Title;
            return View(model);
        }


        public ActionResult Po_Supp_DetailList(string POID, string FabricType)
        {
            var model = _PublicWindowService.Get_Po_Supp_Detail(POID, FabricType, string.Empty);
            ViewData["POID"] = POID;
            ViewData["FabricType"] = FabricType;
            return View(model);
        }

        [HttpPost]
        public ActionResult Po_Supp_DetailList(string POID, string FabricType, string Seq)
        {
            var model = _PublicWindowService.Get_Po_Supp_Detail(POID, FabricType, Seq);
            ViewData["POID"] = POID;
            ViewData["FabricType"] = FabricType;
            return View(model);
        }

        public ActionResult FtyInventoryList(string Title, string POID, string Seq1, string Seq2)
        {
            var model = _PublicWindowService.Get_FtyInventory(POID, Seq1, Seq2, string.Empty);
            ViewData["Title"] = Title;
            ViewData["POID"] = POID;
            ViewData["Seq1"] = Seq1;
            ViewData["Seq2"] = Seq2;
            return View(model);
        }

        [HttpPost]
        public ActionResult FtyInventoryList(string Title, string POID, string Seq1, string Seq2, string Roll)
        {
            var model = _PublicWindowService.Get_FtyInventory(POID, Seq1, Seq2, Roll);
            ViewData["Title"] = Title;
            ViewData["POID"] = POID;
            ViewData["Seq1"] = Seq1;
            ViewData["Seq2"] = Seq2;
            return View(model);
        }

        public ActionResult AppearanceList(string Lab)
        {
            var model = _PublicWindowService.Get_Appearance(Lab);

            return View(model);
        }


        public ActionResult Pass1List(string Title)
        {
            var model = _PublicWindowService.Get_Pass1(string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult Pass1List(string Title, string ID)
        {
            var model = _PublicWindowService.Get_Pass1(ID);
            ViewData["Title"] = Title;
            return View(model);
        }

        public ActionResult SewingLineList(string FactoryID)
        {
            var model = _PublicWindowService.Get_SewingLine(FactoryID , string.Empty);
            ViewData["FactoryID"] = FactoryID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SewingLineList(string FactoryID, string ID)
        {
            var model = _PublicWindowService.Get_SewingLine(FactoryID, ID);
            ViewData["FactoryID"] = FactoryID;
            return View(model);
        }
    }
}