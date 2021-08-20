using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class MeasurementController : BaseController
    {
        private IMeasurementService _MeasurementService;
        public MeasurementController()
        {
            _MeasurementService = new MeasurementService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.Measurement,,";
        }

        // GET: SampleRFT/Measurement
        public ActionResult Index()
        {
            ViewBag.Articles = new List<SelectListItem>();
            TempData["Model"] = null;
            return View(new Measurement_ResultModel());
        }

        [HttpPost]
        public ActionResult Index(Measurement_Request request)
        {
            Measurement_Request measurementRequest = _MeasurementService.MeasurementGetPara(request.OrderID, this.FactoryID);
            Measurement_ResultModel measurement = _MeasurementService.MeasurementGet(measurementRequest);
            List<SelectListItem> ArticleList = new FactoryDashBoardWeb.Helper.SetListItem().ItemListBinding(measurementRequest.Articles);
            ViewBag.Articles = ArticleList;
            TempData["Model"] = measurement.JsonBody;
            return View(measurement);
        }

        [HttpPost]
        public ActionResult SpCheck(string SP)
        {
            Measurement_Request request = _MeasurementService.MeasurementGetPara(SP, this.FactoryID);
            return Json(request);
        }

        public ActionResult ExcelExport()
        {
            if (TempData["Model"] == null)
            {
                return RedirectToAction("Index");
            }

            DataTable dt = (DataTable)JsonConvert.DeserializeObject(TempData["Model"].ToString(), (typeof(DataTable)));

            XSSFWorkbook book;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\Measurement.xlsx", FileMode.Open, FileAccess.Read))
            {
                book = new XSSFWorkbook(file);
                var sheet = book.GetSheetAt(0);

                //IDataFormat dataFormatCustom = book.CreateDataFormat();
                //ICellStyle cellStyleR = book.CreateCellStyle();
                //IFont font = book.CreateFont();
                //cellStyleR.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.BlueGrey.Index;
                //cellStyleR.SetFont(font);

                int ColumnIndex = 0;
                int RowIndex = 0;
                

                // 表頭
                foreach (DataColumn dc in dt.Columns)
                {
                    var headRow = sheet.GetRow(RowIndex);
                    var detail = sheet.GetRow(RowIndex + 1);
                    headRow.CreateCell(ColumnIndex, CellType.String);
                    detail.CreateCell(ColumnIndex, CellType.String);
                    var cell = headRow.GetCell(ColumnIndex);
                    cell.SetCellValue(dc.ColumnName);
                    // cell.CellStyle = cellStyleR;
                    ColumnIndex++;
                }

                RowIndex++;
                int RowCount = dt.Rows.Count;
                // 根據Data數量，複製Row
                for (int i = 0; i <= RowCount - 1; i++)
                {
                    var row = sheet.CopyRow(RowIndex, i + 2);
                    for (int j = 0; j <= dt.Columns.Count - 1; j++)
                    {
                        var cell = row.GetCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }

                sheet.RemoveRow(sheet.GetRow(RowIndex));


                for (int i = 0; i <= ColumnIndex - 1; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            using (var ms = new MemoryStream())
            {
                book.Write(ms);

                Response.AddHeader("Content-Disposition", $"attachment; filename=Measurement_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
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