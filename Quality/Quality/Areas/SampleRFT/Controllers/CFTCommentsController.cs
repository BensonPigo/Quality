
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
            /*
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
            */

            Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\CFT Comments.xltx");

            string xltx_name = "CFT Comments.xltx";

            MyUtility.Excel.CopyToXls(dt, string.Empty, xltx_name, 2, false, null, excelApp, wSheet: excelApp.Sheets[1]);


            #region 存檔 > 讀取MemoryStream > 下載 > 刪除
            string fileName = $""CFT Comments{ DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            workbook.SaveAs(filepath);
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            MemoryStream obj_stream = new MemoryStream();
            var tempFile = System.IO.Path.Combine(filepath);
            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFile));
            Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            Response.BinaryWrite(obj_stream.ToArray());
            obj_stream.Close();
            obj_stream.Dispose();
            Response.Flush();
            Response.End();
            System.IO.File.Delete(filepath);
            #endregion

            string filepath = _ICFTCommentsService.ToExcel(Req);

            var tempFile = System.IO.Path.Combine(filepath);

            MemoryStream obj_stream = new MemoryStream();
            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFile));

            string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
            Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            Response.BinaryWrite(obj_stream.ToArray());
            obj_stream.Close();
            obj_stream.Dispose();
            Response.Flush();
            Response.End();
            System.IO.File.Delete(filepath);

            return View();
        }
    }
}