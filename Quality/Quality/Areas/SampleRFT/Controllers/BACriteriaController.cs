using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.SampleRFT.Service;
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
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.BACriteria,,";
            TempData["Model"] = null;
        }


        // GET: SampleRFT/BACriteria
        public ActionResult Index()
        {
            this.CheckSession();
            BACriteria_ViewModel model = new BACriteria_ViewModel();
            model.DataList = new List<DatabaseObject.ResultModel.BACriteria_Result>();
            TempData["Model"] = null;
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(BACriteria_ViewModel Req)
        {
            this.CheckSession();


            if (Req == null || (string.IsNullOrEmpty(Req.OrderID) && string.IsNullOrEmpty(Req.StyleID)))
            {
                BACriteria_ViewModel e = new BACriteria_ViewModel()
                {
                    ErrorMessage = "Style# and SP# cannot all be empty",
                    DataList = new List<DatabaseObject.ResultModel.BACriteria_Result>()
                };
                return View("Index", e);
            }

            BACriteria_ViewModel model = _BACriteriaService.Get_BACriteria_Result(Req);

            if (!model.Result)
            {
                model.ErrorMessage = $@"
msg.WithInfo('{model.ErrorMessage}');
";
            }

            TempData["Model"] = model;
            return View("Index", model);
        }

        public ActionResult ExcelExport()
        {
            if (TempData["Model"] == null)
            {
                return RedirectToAction("Index");
            }

            BACriteria_ViewModel model = (BACriteria_ViewModel)TempData["Model"];
            List<DatabaseObject.ResultModel.BACriteria_Result> dataList = model.DataList;
            TempData["Model"] = model;

            XSSFWorkbook book;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\BACriteria.xlsx", FileMode.Open, FileAccess.Read))
            {
                book = new XSSFWorkbook(file);
                var sheet = book.GetSheetAt(0);

                int RowCount = dataList.Count;
                // 根據Data數量，複製Row
                for (int i = 0; i <= RowCount; i++)
                {
                    var firstRow = sheet.GetRow(1);
                    firstRow.CopyRowTo(i + 2);
                }

                IDataFormat dataFormatCustom = book.CreateDataFormat();
                ICellStyle cellStyleR = book.CreateCellStyle();

                int RowIndex = 1;
                foreach (var data in dataList)
                {
                    var row = sheet.GetRow(RowIndex);


                    var cell0 = row.GetCell(0);
                    cell0.SetCellValue(data.OrderID);

                    var cell1 = row.GetCell(1);
                    cell1.SetCellValue(data.OrderTypeID);

                    var cell2 = row.GetCell(2);
                    cell2.SetCellValue(data.Qty);

                    var cell3 = row.GetCell(3);
                    cell3.SetCellValue(data.InspectedQty);

                    var cell4 = row.GetCell(4);
                    cell4.SetCellValue(data.BAProduct);

                    var cell5 = row.GetCell(5);
                    cell5.SetCellValue(Convert.ToDouble(data.BACriteria));

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