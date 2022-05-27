using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspBySPQueryController : BaseController
    {
        private InspectionBySPService _Service;
        public InspBySPQueryController()
        {
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.InspBySPQuery,,";
            _Service = new InspectionBySPService();
        }
        // GET: SampleRFT/InspBySPQuery
        public ActionResult Index()
        {
            QueryInspectionBySP_ViewModel model = new QueryInspectionBySP_ViewModel() { DataList = new List<QueryInspectionBySP>() };
            //if (model.Result)
            //{
            //    model = _Service.GetQuery(model);
            //}
            //else
            //{
            //    model.ErrorMessage = $@"msg.WithError(""{model.ErrorMessage}"")";
            //}

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(QueryInspectionBySP_ViewModel Req)
        {
            QueryInspectionBySP_ViewModel model = new QueryInspectionBySP_ViewModel();
            if (model.Result)
            {
                model = _Service.GetQuery(Req);
            }
            else
            {
                model.ErrorMessage = $@"msg.WithError(""{model.ErrorMessage}"")";
            }

            TempData["Req"] = Req;
            return View(model);
        }
        public ActionResult Detail(long ID)
        {

            QueryReport model = _Service.GetQueryDetail(ID, this.UserID);

            return View(model);
        }

        public ActionResult IndexToExcel()
        {
            if (TempData["Req"] == null)
            {
                return RedirectToAction("Index");
            }
            QueryInspectionBySP_ViewModel Req = (QueryInspectionBySP_ViewModel)TempData["Req"];

            QueryInspectionBySP_ViewModel model = new QueryInspectionBySP_ViewModel();
            if (model.Result)
            {
                model = _Service.GetQuery(Req);
            }
            else
            {
                model.ErrorMessage = $@"msg.WithError(""{model.ErrorMessage}"")";
            }

            var dataList = model.DataList;

            XSSFWorkbook book;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\InspBySPQuery.xlsx", FileMode.Open, FileAccess.Read))
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
                    cell0.SetCellValue(data.SP);

                    var cell1 = row.GetCell(1);
                    cell1.SetCellValue(data.CustPONO);

                    var cell2 = row.GetCell(2);
                    cell2.SetCellValue(data.StyleID);

                    var cell3 = row.GetCell(3);
                    cell3.SetCellValue(data.SeasonID);

                    var cell4 = row.GetCell(4);
                    cell4.SetCellValue(data.Article);

                    var cell5 = row.GetCell(5);
                    cell5.SetCellValue(data.SampleStage);

                    cell5 = row.GetCell(6);
                    cell5.SetCellValue(data.SewingLineID);

                    cell5 = row.GetCell(7);
                    cell5.SetCellValue(data.InspectionTimes);

                    cell5 = row.GetCell(8);
                    cell5.SetCellValue(data.Inspector);

                    cell5 = row.GetCell(9);
                    cell5.SetCellValue(data.Result);

                    RowIndex++;
                }

            }

            using (var ms = new MemoryStream())
            {
                book.Write(ms);

                Response.AddHeader("Content-Disposition", $"attachment; filename=InspBySPQuery_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
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