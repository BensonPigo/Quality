
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using Ict;
using Sci.Data;
using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sci;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class CFTCommentsController : BaseController
    {
        private ICFTCommentsService _ICFTCommentsService;

        public CFTCommentsController()
        {
            _ICFTCommentsService = new CFTCommentsService();
            this.SelectedMenu = "Sample RFT";

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

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "ToExcel")]
        public ActionResult ToExcel(CFTComments_ViewModel Req)
        {
            this.CheckSession();

            DataTable dt = new DataTable();

            string tempFilePath = _ICFTCommentsService.GetExcel(Req);

            // 檔名命名
            string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

            //下載
            MemoryStream obj_stream = new MemoryStream();

            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFilePath));
            Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            Response.BinaryWrite(obj_stream.ToArray());
            obj_stream.Close();
            obj_stream.Dispose();
            Response.Flush();
            Response.End();

            //等待下載完畢後刪除暫存檔
            System.Threading.Thread.Sleep(3000);
            System.IO.File.Delete(tempFilePath);


            return View();
        }
    }
}