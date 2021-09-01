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
        public ActionResult BrandList(string BrandID, string ReturnType)
        {
            var model = _PublicWindowService.Get_Brand(BrandID);
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        [HttpPost]
        public ActionResult BrandListOther(string BrandID, string ReturnType ,string OtherTable, string OtherColumn)
        {
            var model = _PublicWindowService.Get_Brand(BrandID, OtherTable, OtherColumn);
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult SeasonList(string BrandID)
        {
            var model = _PublicWindowService.Get_Season(BrandID, string.Empty);
            ViewData["BrandID"] = BrandID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SeasonList(string BrandID, string SeasonID, string ReturnType)
        {
            var model = _PublicWindowService.Get_Season(BrandID, SeasonID);
            ViewData["BrandID"] = BrandID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
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
        public ActionResult StyleList(string BrandID, string SeasonID, string StyleID, string ReturnType)
        {
            var model = _PublicWindowService.Get_Style(BrandID, SeasonID, StyleID);
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult ArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID)
        {
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, string.Empty);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            return View(model);
        }

        [HttpPost]
        public ActionResult ArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, string ReturnType)
        {
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, Article);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult SizeList(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article)
        {
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, BrandID, SeasonID, StyleID, Article, string.Empty);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["StyleID"] = StyleID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SizeList(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article, string Size, string ReturnType)
        {
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, BrandID, SeasonID, StyleID, Article, Size);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["StyleID"] = StyleID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult TechnicianList(string CallFunction)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, string.Empty);
            ViewData["CallFunction"] = CallFunction;
            return View(model);
        }

        [HttpPost]
        public ActionResult TechnicianList(string CallFunction, string ID, string ReturnType)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, ID);
            ViewData["CallFunction"] = CallFunction;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult LocalSuppList(string Title)
        {
            var model = _PublicWindowService.Get_LocalSupp( string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult LocalSuppList(string Title, string Name, string ReturnType)
        {
            var model = _PublicWindowService.Get_LocalSupp(Name);
            ViewData["Title"] = Title;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult TPESuppList(string Title)
        {
            var model = _PublicWindowService.Get_TPESupp(string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult TPESuppList(string Title, string Name, string ReturnType)
        {
            var model = _PublicWindowService.Get_TPESupp(Name);
            ViewData["Title"] = Title;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
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
        public ActionResult Po_Supp_DetailList(string POID, string FabricType, string Seq, string ReturnType)
        {
            var model = _PublicWindowService.Get_Po_Supp_Detail(POID, FabricType, Seq);
            ViewData["POID"] = POID;
            ViewData["FabricType"] = FabricType;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
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
        public ActionResult FtyInventoryList(string Title, string POID, string Seq1, string Seq2, string Roll, string ReturnType)
        {
            var model = _PublicWindowService.Get_FtyInventory(POID, Seq1, Seq2, Roll);
            ViewData["Title"] = Title;
            ViewData["POID"] = POID;
            ViewData["Seq1"] = Seq1;
            ViewData["Seq2"] = Seq2;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult Pass1List(string Title)
        {
            var model = _PublicWindowService.Get_Pass1(string.Empty);
            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        public ActionResult Pass1List(string Title, string ID, string ReturnType)
        {
            var model = _PublicWindowService.Get_Pass1(ID);
            ViewData["Title"] = Title;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult SewingLineList(string FactoryID)
        {
            var model = _PublicWindowService.Get_SewingLine(FactoryID , string.Empty);
            ViewData["FactoryID"] = FactoryID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SewingLineList(string FactoryID, string ID, string ReturnType)
        {
            var model = _PublicWindowService.Get_SewingLine(FactoryID, ID);
            ViewData["FactoryID"] = FactoryID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult AppearanceList(string Lab)
        {
            var model = _PublicWindowService.Get_Appearance(Lab);

            return View(model);
        }
        [HttpPost]
        public ActionResult AppearanceList(string Lab, string ReturnType)
        {
            var model = _PublicWindowService.Get_Appearance(Lab);

            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult ColorList(string Title, string BrandID)
        {
            //var model = _PublicWindowService.Get_Color(BrandID, string.Empty);

            // Color資料太多，畫面開啟不預設帶入資料
            List<DatabaseObject.Public.Window_Color> model = new List<DatabaseObject.Public.Window_Color>();

            ViewData["Title"] = Title;
            return View(model);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult ColorList(string Title, string BrandID, string ID, string PickedIDs, string PickedNames)
        {
            var result = _PublicWindowService.Get_Color(BrandID, ID);
            ViewData["Title"] = Title;

            //整理已選擇的ID和Name，需要保留
            List<string> PickedID_List = PickedIDs.Replace("&#39;", "").Split(',').Where(o => !string.IsNullOrEmpty(o)).ToList();
            List<string> PickedName_List = PickedNames.Replace("&#39;", "").Split(',').Where(o => !string.IsNullOrEmpty(o)).ToList();

            //排除已選擇的ID
            var model = result.Where(o => !PickedID_List.Contains(o.ID)).ToList();

            ViewData["PickedIDs"] = string.Join(",", PickedIDs.Split(',').Where(o => !string.IsNullOrEmpty(o)).Distinct());
            ViewData["PickedNames"] = string.Join(",", PickedNames.Split(',').Where(o => !string.IsNullOrEmpty(o)).Distinct());

            return View(model);
        }

        public ActionResult FGPTList(string VersionID)
        {
            var model = _PublicWindowService.Get_FGPT(VersionID, string.Empty);
            ViewData["VersionID"] = VersionID;
            return View(model);
        }

        [HttpPost]
        public ActionResult FGPTList(string VersionID, string Code)
        {
            var model = _PublicWindowService.Get_FGPT(VersionID, Code);
            ViewData["VersionID"] = VersionID;
            return View(model);
        }

        public ActionResult PictureList(string Title, bool EditMode, string Table, string BrforeColumn, string AfterColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val)
        {
            var model = _PublicWindowService.Get_Picture(Table, BrforeColumn, AfterColumn, PKey_1, PKey_2, PKey_3, PKey_1_Val, PKey_2_Val, PKey_3_Val);
            ViewData["Title"] = Title;
            ViewData["EditMode"] = EditMode;
            return View(model);
        }


        public ActionResult TestFailMailList(string Title, string FactoryID, string Type)
        {
            var model = _PublicWindowService.Get_TestFailMail(FactoryID, Type, string.Empty);
            ViewData["Title"] = Title;
            ViewData["Type"] = Type;
            ViewData["FactoryID"] = FactoryID;
            return View(model);
        }

        [HttpPost]
        public ActionResult TestFailMailList(string Title, string FactoryID, string Type, string GroupNameList)
        {
            var model = _PublicWindowService.Get_TestFailMail(FactoryID, Type, GroupNameList);


            ViewData["Type"] = Type;
            ViewData["Title"] = Title;
            ViewData["FactoryID"] = FactoryID;
            return View(model);
        }
    }
}