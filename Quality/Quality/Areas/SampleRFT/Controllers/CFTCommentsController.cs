
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.ViewModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class CFTCommentsController : BaseController
    {
        private ICFTCommentsService _ICFTCommentsService;

        public CFTCommentsController()
        {
            _ICFTCommentsService = new CFTCommentsService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.CFTComments,,";
            TempData["Model"] = null;
            TempData["tempFilePath"] = null;
        }


        // GET: SampleRFT/CFTComments
        public ActionResult Index()
        {
            this.CheckSession();
            CFTComments_ViewModel model = new CFTComments_ViewModel() { QueryType = "Style",DataList = new List<CFTComments_Result>()};

            if (TempData["Model"] != null)
            {
                model =(CFTComments_ViewModel)TempData["Model"];
            }
            if (TempData["tempFilePath"] != null)
            {
                ViewData["tempFilePath"] = TempData["tempFilePath"].ToString();
            }
            return View(model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(CFTComments_ViewModel Req)
        {
            this.CheckSession();

            CFTComments_ViewModel model = new CFTComments_ViewModel();
            Req.DataList = new List<CFTComments_Result>();

            if (Req.QueryType == "OrderID")
            {
                if (Req.OrderID == null || string.IsNullOrEmpty(Req.OrderID))
                {

                    Req.ErrorMessage = $@"
msg.WithInfo('SP# cannot be emptry');
";
                    return View("Index", Req);
                }

                model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID = Req.OrderID ,QueryType= "OrderID" });

                if (model.OrderID == null)
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('Cannot found SP# {Req.OrderID}');
";
                    return View("Index", Req);
                }

            }
            else if (Req.QueryType == "Style")
            {
                if (string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID)
                    || Req.StyleID == null || Req.BrandID == null || Req.SeasonID == null)
                {

                    Req.ErrorMessage = $@"
msg.WithInfo('Style#, Brand and Season cannot be emptry');
";
                    return View("Index", Req);
                }

                model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel()
                {
                    StyleID = Req.StyleID,
                    BrandID = Req.BrandID,
                    SeasonID = Req.SeasonID, 
                    QueryType = "Style"
                });

                if (model.StyleID == null || model.StyleID == string.Empty)
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('Cannot found combination Style# {Req.StyleID}, Brand {Req.BrandID}, Season {Req.SeasonID}');
";
                    return View("Index", Req);
                }
            }

            // Query
            model = _ICFTCommentsService.Get_CFT_OrderComments(model);

            if (!model.Result)
            {
                model.ErrorMessage = $@"
msg.WithInfo('{model.ErrorMessage.Replace("'",string.Empty)}');
";
            }
            model.QueryType = Req.QueryType;
            TempData["Model"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetOrderinfo(string OrderID)
        {
            this.CheckSession();
            CFTComments_ViewModel model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID=OrderID});

            return View();
        }

        /// <summary>
        /// 使用 NPOI 產生Excel 的寫法，可參考：I:\MIS\Personal\Benson\密技\NPOI.png
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "ToExcel_NPOI")]
        public ActionResult ToExcel_NPOI(CFTComments_ViewModel Req)
        {
            this.CheckSession();

            try
            {

                CFTComments_ViewModel model = new CFTComments_ViewModel();
                Req.DataList = new List<CFTComments_Result>();

                if (Req.QueryType == "OrderID")
                {
                    if (Req.OrderID == null || string.IsNullOrEmpty(Req.OrderID))
                    {

                        Req.ErrorMessage = $@"
msg.WithInfo('SP# cannot be emptry');
";
                        return View("Index", Req);
                    }

                    model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID = Req.OrderID });

                    if (model.OrderID == null)
                    {
                        Req.ErrorMessage = $@"
msg.WithInfo('Cannot found SP# {Req.OrderID}');
";
                        return View("Index", Req);
                    }

                }
                else if (Req.QueryType == "Style")
                {
                    if (string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID)
                        || Req.StyleID == null || Req.BrandID == null || Req.SeasonID == null)
                    {

                        Req.ErrorMessage = $@"
msg.WithInfo('Style#, Brand and Season cannot be emptry');
";
                        return View("Index", Req);
                    }

                    model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel()
                    {
                        StyleID = Req.StyleID,
                        BrandID = Req.BrandID,
                        SeasonID = Req.SeasonID,
                    });

                    if (model.OrderID == null)
                    {
                        Req.ErrorMessage = $@"
msg.WithInfo('Cannot found combination Style# {Req.StyleID}, Brand {Req.BrandID}, Season {Req.SeasonID}');
";
                        return View("Index", Req);
                    }
                }

                // 取得資料
                model = _ICFTCommentsService.Get_CFT_OrderComments(model);

                XSSFWorkbook book;
                using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\CFT Comments.xlsx", FileMode.Open, FileAccess.Read))
                {
                    book = new XSSFWorkbook(file);
                    ISheet sheet = book.GetSheetAt(0);


                    for (int RowIdx = 0; RowIdx < model.DataList.Count; RowIdx++)
                    {
                        //注意，要CreateRoll才對
                        sheet.CreateRow(RowIdx + 1);

                        //注意，要CreateCell才對
                        sheet.GetRow(RowIdx + 1).CreateCell(0).SetCellValue(model.DataList[RowIdx].SampleStage);
                        sheet.GetRow(RowIdx + 1).CreateCell(1).SetCellValue(model.DataList[RowIdx].CommentsCategory);
                        sheet.GetRow(RowIdx + 1).CreateCell(2).SetCellValue(model.DataList[RowIdx].Comnments);
                    }


                }


                using (var ms = new MemoryStream())
                {
                    book.Write(ms);

                    Response.AddHeader("Content-Disposition", $"attachment; filename=CFT Comments_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
                    Response.BinaryWrite(ms.ToArray());

                    //== 釋放資源
                    book = null;
                    ms.Close();
                    ms.Dispose();

                    Response.Flush();
                    Response.End();
                }

                

            }
            catch (Exception ex)
            {

                throw ex;
            }
            return View();
        }

        /// <summary>
        /// 使用Microsoft.Office.Interop.Excel的寫法
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "ToExcel")]
        public ActionResult ToExcel(CFTComments_ViewModel Req)
        {
            try
            {
                Req.DataList = new List<CFTComments_Result>();
                CFTComments_ViewModel model = new CFTComments_ViewModel();
                if (Req.QueryType == "OrderID")
                {
                    if (Req.OrderID == null || string.IsNullOrEmpty(Req.OrderID))
                    {

                        Req.ErrorMessage = $@"
msg.WithInfo('SP# cannot be emptry');
";
                        return View("Index", Req);
                    }

                    model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID = Req.OrderID, QueryType = "OrderID" });

                    if (model.OrderID == null)
                    {
                        Req.ErrorMessage = $@"
msg.WithInfo('Cannot found SP# {Req.OrderID}');
";
                        return View("Index", Req);
                    }

                }
                else if (Req.QueryType == "Style")
                {
                    if (string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID)
                        || Req.StyleID == null || Req.BrandID == null || Req.SeasonID == null)
                    {

                        Req.ErrorMessage = $@"
msg.WithInfo('Style#, Brand and Season cannot be emptry');
";
                        return View("Index", Req);
                    }

                    model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel()
                    {
                        StyleID = Req.StyleID,
                        BrandID = Req.BrandID,
                        SeasonID = Req.SeasonID,
                        QueryType = "Style",
                    });

                    if (model.StyleID == null || model.StyleID == string.Empty)
                    {
                        Req.ErrorMessage = $@"
msg.WithInfo('Cannot found combination Style# {Req.StyleID}, Brand {Req.BrandID}, Season {Req.SeasonID}');
";
                        return View("Index", Req);
                    }
                }

                // 1. 在Service層取得資料，生成Excel檔案，放在暫存路徑，回傳檔名

                CFTComments_ViewModel result = _ICFTCommentsService.GetExcel(model);
                string tempFilePath = result.TempFileName;

                // 2. 取得hotst name，串成下載URL ，傳到準備前端下載
                // URL範例：https://misap:1880/TMP/CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx 
                // (暫存檔檔名是CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx)
                tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;

                // 3. 前端下載方式：請參考Index.cshtml的 「window.location.href = '@download'」;

                model = _ICFTCommentsService.Get_CFT_OrderComments(model);
                model.QueryType = Req.QueryType;

                TempData["tempFilePath"] = tempFilePath;
                TempData["Model"] = model;
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return RedirectToAction("Index");
        }
    }
}