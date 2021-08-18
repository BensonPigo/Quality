using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace Quality.Areas.SampleRFT.Controllers
{
    public class BACriteriaController : BaseController
    {
        private IBACriteriaService _BACriteriaService;

        public BACriteriaController()
        {
            _BACriteriaService = new BACriteriaService();
            this.SelectedMenu = "Sample RFT";

        }


        // GET: SampleRFT/BACriteria
        public ActionResult Index()
        {
            this.CheckSession();
            BACriteria_ViewModel model = new BACriteria_ViewModel();
            model.DataList = new List<DatabaseObject.ResultModel.BACriteria_Result>();
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(BACriteria_ViewModel Req)
        {
            this.CheckSession();


            if (Req == null || (!string.IsNullOrEmpty(Req.OrderID) && !string.IsNullOrEmpty(Req.StyleID)))
            {
                BACriteria_ViewModel e = new BACriteria_ViewModel()
                {
                    ErrorMessage = "Style# and SP# cannot all be empty"
                };
                return View("Index", e);
            }

            BACriteria_ViewModel model = _BACriteriaService.Get_BACriteria_Result(Req);

            if (!model.Result)
            {
                model.ErrorMessage = $@"
msg.WithError('{model.ErrorMessage}');
";
            }

            TempData["Model"] = model;
            return View("Index", model);
        }

        public FileResult ExcelExport()
        {
            if (TempData["Model"] == null)
            {
                return null;
            }

            BACriteria_ViewModel model = (BACriteria_ViewModel)TempData["Model"];
            List<DatabaseObject.ResultModel.BACriteria_Result> dataList = model.DataList;

            XSSFWorkbook book;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\BACriteria.xlsx", FileMode.Open, FileAccess.Read))
            {
                book = new XSSFWorkbook(file);
                var sheet = book.GetSheetAt(0);


                IDataFormat dataFormatCustom = book.CreateDataFormat();
                ICellStyle cellStyleR = book.CreateCellStyle();

                int RowIndex = 2;
                foreach (var data in dataList)
                {
                    var row = sheet.GetRow(RowIndex);

                    row.GetCell(0).SetCellValue(data.OrderID);
                    row.GetCell(1).SetCellValue(data.OrderTypeID);
                    row.GetCell(2).SetCellValue(data.Qty);
                    row.GetCell(3).SetCellValue(data.InspectedQty);
                    row.GetCell(4).SetCellValue(data.BAProduct);
                    row.GetCell(5).SetCellValue(data.BACriteria);

                    RowIndex++;
                }

            }

            using (var ms = new MemoryStream())
            {
                book.Write(ms);

                Response.AddHeader("Content-Disposition", $"attachment; filename=BA Criteria_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
                Response.BinaryWrite(ms.ToArray());

                //== 釋放資源
                book = null;
                ms.Close();
                ms.Dispose();

                Response.Flush();
                Response.End();
            }

            return null;
        }
    }
}