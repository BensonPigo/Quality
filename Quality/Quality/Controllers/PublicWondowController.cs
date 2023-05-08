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
        public ActionResult BrandList(string TargetID)
        {
            var model = _PublicWindowService.Get_Brand(null, false);
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult BrandList(string BrandID, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Brand(BrandID, IsExact);
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        [HttpPost]
        public ActionResult BrandListOther(string BrandID, string OtherTable, string OtherColumn, string TargetID, string ReturnType = "")
        {
            var model = _PublicWindowService.Get_Brand(BrandID, OtherTable, OtherColumn);
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult SeasonList(string BrandID, string TargetID)
        {
            var model = _PublicWindowService.Get_Season(BrandID, string.Empty, false);
            ViewData["BrandID"] = BrandID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SeasonList(string BrandID, string SeasonID, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Season(BrandID, SeasonID, IsExact);
            ViewData["BrandID"] = BrandID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult StyleList(string BrandID, string SeasonID, string TargetID)
        {
            var model = _PublicWindowService.Get_Style(BrandID, SeasonID, string.Empty, false);
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult StyleList(string BrandID, string SeasonID, string StyleID, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Style(BrandID, SeasonID, StyleID, IsExact);
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult ArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string TargetID)
        {
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, string.Empty, false);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult ArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Article(OrderID, StyleUkey, StyleID, BrandID, SeasonID, Article, IsExact);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult PoidArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string TargetID)
        {
            var model = _PublicWindowService.Get_PoidArticle(OrderID, StyleUkey, StyleID, BrandID, SeasonID, string.Empty, false);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult PoidArticleList(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_PoidArticle(OrderID, StyleUkey, StyleID, BrandID, SeasonID, Article, IsExact);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["StyleID"] = StyleID;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult SizeList(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article, string TargetID)
        {
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, BrandID, SeasonID, StyleID, Article, string.Empty, false);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["StyleID"] = StyleID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SizeList(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article, string Size, string TargetID, string ReturnType = "")
        {
            // 若是驗證則需要精準判斷
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Size(OrderID, StyleUkey, BrandID, SeasonID, StyleID, Article, Size, IsExact);
            ViewData["OrderID"] = OrderID;
            ViewData["StyleUkey"] = StyleUkey;
            ViewData["Article"] = Article;
            ViewData["BrandID"] = BrandID;
            ViewData["SeasonID"] = SeasonID;
            ViewData["StyleID"] = StyleID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult TechnicianList(string CallFunction, string TargetID)
        {
            var model = _PublicWindowService.Get_Technician(CallFunction, string.Empty, false);
            ViewData["CallFunction"] = CallFunction;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult TechnicianList(string CallFunction, string TargetID, string ID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Technician(CallFunction, ID, IsExact);
            ViewData["CallFunction"] = CallFunction;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult LocalSuppList(string Title, string TargetID)
        {
            var model = _PublicWindowService.Get_LocalSupp(string.Empty, string.Empty, false);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult LocalSuppList(string Title, string TargetID, string SuppID, string Name, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_LocalSupp(SuppID, Name, IsExact);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult TPESuppList(string Title, string TargetID)
        {
            var model = _PublicWindowService.Get_TPESupp(string.Empty, string.Empty, false);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult TPESuppList(string Title, string TargetID, string SuppID, string Name, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_TPESupp(SuppID, Name, IsExact);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult Po_Supp_DetailList(string POID, string FabricType, string TargetID)
        {
            var model = _PublicWindowService.Get_Po_Supp_Detail(POID, FabricType, string.Empty, false);
            ViewData["POID"] = POID;
            ViewData["FabricType"] = FabricType;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult Po_Supp_DetailList(string POID, string FabricType, string Seq, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Po_Supp_Detail(POID, FabricType, Seq, IsExact);
            ViewData["POID"] = POID;
            ViewData["FabricType"] = FabricType;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult Po_Supp_Detail_RefnoList(string OrderID, string MtlTypeID, string TargetID)
        {
            var model = _PublicWindowService.Get_Po_Supp_Detail_Refno(OrderID, MtlTypeID, string.Empty);
            ViewData["OrderID"] = OrderID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult Po_Supp_Detail_RefnoList(string OrderID, string MtlTypeID, string Refno, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Po_Supp_Detail_Refno(OrderID, MtlTypeID, Refno);
            ViewData["OrderID"] = OrderID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult FtyInventoryList(string Title, string POID, string Seq1, string Seq2, string TargetID)
        {
            var model = _PublicWindowService.Get_FtyInventory(POID, Seq1, Seq2, string.Empty, false);
            ViewData["Title"] = Title;
            ViewData["POID"] = POID;
            ViewData["Seq1"] = Seq1;
            ViewData["Seq2"] = Seq2;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult FtyInventoryList(string Title, string POID, string Seq1, string Seq2, string Roll, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_FtyInventory(POID, Seq1, Seq2, Roll, IsExact);
            ViewData["Title"] = Title;
            ViewData["POID"] = POID;
            ViewData["Seq1"] = Seq1;
            ViewData["Seq2"] = Seq2;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }


        public ActionResult Pass1List(string Title, string TargetID)
        {
            var model = _PublicWindowService.Get_Pass1(string.Empty, false);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult Pass1List(string Title, string ID, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_Pass1(ID, IsExact);
            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult SewingLineList(string FactoryID, string TargetID)
        {
            var model = _PublicWindowService.Get_SewingLine(FactoryID, string.Empty, false);
            ViewData["FactoryID"] = FactoryID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult SewingLineList(string FactoryID, string ID, string TargetID, string ReturnType = "")
        {
            bool IsExact = ReturnType.ToUpper() == "JSON";
            var model = _PublicWindowService.Get_SewingLine(FactoryID, ID, IsExact);
            ViewData["FactoryID"] = FactoryID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult AppearanceList(string Lab, string TargetID)
        {
            var model = _PublicWindowService.Get_Appearance(Lab);
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;

            return View(model);
        }
        [HttpPost]
        public ActionResult AppearanceList(string Lab, string TargetID, string ReturnType = "")
        {
            var model = _PublicWindowService.Get_Appearance(Lab);
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;

            if (ReturnType.ToUpper() == "JSON")
            {
                return Json(model);
            }
            return View(model);
        }

        public ActionResult ColorList(string Title, string BrandID, string TargetID)
        {
            //var model = _PublicWindowService.Get_Color(BrandID, string.Empty);

            // Color資料太多，畫面開啟不預設帶入資料
            List<DatabaseObject.Public.Window_Color> model = new List<DatabaseObject.Public.Window_Color>();

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult ColorList(string Title, string BrandID, string ID, string PickedIDs, string TargetID, string PickedNames)
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
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;

            return View(model);
        }

        public ActionResult FGPTList(string VersionID, string TargetID)
        {
            var model = _PublicWindowService.Get_FGPT(VersionID, string.Empty);
            ViewData["VersionID"] = VersionID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult FGPTList(string VersionID, string Code, string TargetID)
        {
            var model = _PublicWindowService.Get_FGPT(VersionID, Code);
            ViewData["VersionID"] = VersionID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        public ActionResult PictureList(string Title, bool EditMode, string TargetBeforeColumn, string TargetAfterColumn, string TargetID, string Table, string BeforeColumn, string AfterColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            var model = _PublicWindowService.Get_Picture(Table, BeforeColumn, AfterColumn, PKey_1, PKey_2, PKey_3, PKey_4, PKey_1_Val, PKey_2_Val, PKey_3_Val, PKey_4_Val);
            ViewData["Title"] = Title;
            ViewData["EditMode"] = EditMode;
            ViewData["BeforeColumn"] = string.IsNullOrEmpty(TargetBeforeColumn) ? BeforeColumn : TargetBeforeColumn;
            ViewData["AfterColumn"] = string.IsNullOrEmpty(TargetAfterColumn) ? AfterColumn : TargetAfterColumn;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }
        public ActionResult InspMeasurement_PictureList(string Title, bool EditMode, string TargetColumnName, string TargetID, string Table, string ColumnName, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val, bool ShowDelete = true)
        {
            var model = _PublicWindowService.Get_SinglePicture(Table, ColumnName, PKey_1, PKey_2, PKey_3, PKey_4, PKey_1_Val, PKey_2_Val, PKey_3_Val, PKey_4_Val);
            ViewData["Title"] = Title;
            ViewData["EditMode"] = EditMode;
            ViewData["ColumnName"] = string.IsNullOrEmpty(TargetColumnName) ? ColumnName : TargetColumnName;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            ViewData["ShowDelete"] = ShowDelete;
            return View(model);
        }

        public ActionResult TestFailMailList(string Title, string FactoryID, string Type, string TargetID, bool ExitReload = true)
        {
            var model = _PublicWindowService.Get_TestFailMail(FactoryID, Type, string.Empty);
            ViewData["Title"] = Title;
            ViewData["Type"] = Type;
            ViewData["FactoryID"] = FactoryID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            ViewBag.ExitReload = ExitReload ? "window.opener.document.location.reload();" : string.Empty;
            return View(model);
        }

        [HttpPost]
        public ActionResult TestFailMailList(string Title, string FactoryID, string Type, string GroupNameList, string TargetID, bool ExitReload = true)
        {
            var model = _PublicWindowService.Get_TestFailMail(FactoryID, Type, GroupNameList);
            ViewData["Type"] = Type;
            ViewData["Title"] = Title;
            ViewData["FactoryID"] = FactoryID;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            ViewBag.ExitReload = ExitReload ? "window.opener.document.location.reload();" : string.Empty;
            return View(model);
        }
        public ActionResult OperationList(string Title, string FinalInspectionID, string TargetID)
        {
            var model = _PublicWindowService.Get_Operation(FinalInspectionID, string.Empty);

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult OperationList(string Title, string FinalInspectionID, string Operation, string TargetID)
        {
            var model = _PublicWindowService.Get_Operation(FinalInspectionID, Operation);

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }
        public ActionResult FabricRefNoList(string Title, string OrderID, string TargetID)
        {
            var model = _PublicWindowService.Get_FabricRefNo(OrderID, string.Empty);

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        [HttpPost]
        public ActionResult FabricRefNoList(string Title, string OrderID, string Refno, string TargetID)
        {
            var model = _PublicWindowService.Get_FabricRefNo(OrderID, Refno);

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

        public ActionResult RollDyelotList(string Title, string OrderID, string Seq1, string Seq2 ,string TargetID)
        {
            var model = _PublicWindowService.Get_RollDyelot(OrderID, Seq1 , Seq2);

            ViewData["Title"] = Title;
            ViewData["TargetID"] = TargetID == null ? string.Empty : TargetID;
            return View(model);
        }

    }
}