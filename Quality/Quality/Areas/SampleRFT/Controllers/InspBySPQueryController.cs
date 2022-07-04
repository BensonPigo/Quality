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
using Quality.Helper;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using Excel = Microsoft.Office.Interop.Excel;


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
        [SessionAuthorize]
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

            TempData["InspBySPQueryReq"] = Req;
            return View(model);
        }
        public ActionResult Detail(long ID)
        {

            QueryReport model = _Service.GetQueryDetail(ID, this.UserID);

            TempData["AllSize"] = model.DummyFit.ArticleSizeList;
            TempData["ModelQuery"] = model;
            return View(model);
        }

        public ActionResult IndexToExcel()
        {
            if (TempData["InspBySPQueryReq"] == null)
            {
                return RedirectToAction("Index");
            }
            QueryInspectionBySP_ViewModel Req = (QueryInspectionBySP_ViewModel)TempData["InspBySPQueryReq"];

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


        public ActionResult DownLoad()
        {
            GetExcel(true);

            return null;
        }

        public string GetExcel(bool IsDowdload)
        {
            string fileName = string.Empty;

            QueryReport model = (QueryReport)TempData["ModelQuery"];
            TempData["ModelQuery"] = model;

            Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\InspBySPQuery_Detail.xlsx");
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Excel.Worksheet worksheet = excelApp.Sheets[1];

            worksheet.Cells[4, 3] = model.sampleRFTInspection.InspectionTimesText;

            worksheet.Cells[6, 6] = string.Empty;
            worksheet.Cells[6, 10] = model.sampleRFTInspection.FactoryID;
            worksheet.Cells[6, 14] = model.sampleRFTInspection.AddName;

            worksheet.Cells[7, 6] = model.Setting.Model;

            worksheet.Cells[8, 6] = model.sampleRFTInspection.WorkNo;

            worksheet.Cells[9, 6] = model.sampleRFTInspection.POID;
            worksheet.Cells[9, 14] = model.sampleRFTInspection.OrderID;

            worksheet.Cells[10, 6] = model.Setting.Article;

            worksheet.Cells[11, 6] = model.Setting.OrderQty;
            worksheet.Cells[11, 14] = model.sampleRFTInspection.SewingLineID;

            worksheet.Cells[12, 6] = model.Setting.CustPONo;

            worksheet.Cells[13, 6] = string.Empty;

            worksheet.Cells[14, 6] = model.Setting.AQLPlan;
            worksheet.Cells[14, 10] = model.sampleRFTInspection.SampleSize;
            worksheet.Cells[14, 14] = model.sampleRFTInspection.AcceptQty;
            worksheet.Cells[14, 16] = model.sampleRFTInspection.RejectQty;

            worksheet.Cells[15, 6] = model.sampleRFTInspection.CheckFabricApproval ? "V" : string.Empty;

            worksheet.Cells[15, 12] = model.sampleRFTInspection.CheckSealingSampleApproval ? "V" : string.Empty;

            worksheet.Cells[16, 6] = model.sampleRFTInspection.CheckMetalDetection ? "V" : string.Empty;

            worksheet.Cells[18, 6] = model.sampleRFTInspection.CheckColorShade ? "V" : string.Empty;
            worksheet.Cells[19, 6] = model.sampleRFTInspection.CheckAppearance ? "V" : string.Empty;
            worksheet.Cells[20, 6] = model.sampleRFTInspection.CheckHandfeel ? "V" : string.Empty;

            worksheet.Cells[18, 10] = model.sampleRFTInspection.CheckFiberContent ? "V" : string.Empty;
            worksheet.Cells[19, 10] = model.sampleRFTInspection.CheckCountryofOrigin ? "V" : string.Empty;
            worksheet.Cells[20, 10] = model.sampleRFTInspection.CheckCareInstructions ? "V" : string.Empty;
            worksheet.Cells[21, 10] = model.sampleRFTInspection.CheckSizeKey ? "V" : string.Empty;

            worksheet.Cells[18, 14] = model.sampleRFTInspection.CheckDecorativeLabel ? "V" : string.Empty;
            worksheet.Cells[19, 14] = model.sampleRFTInspection.CheckCareLabel ? "V" : string.Empty;
            worksheet.Cells[20, 14] = model.sampleRFTInspection.CheckSecurityLabel ? "V" : string.Empty;
            worksheet.Cells[21, 14] = model.sampleRFTInspection.CheckAdditionalLabel ? "V" : string.Empty;

            worksheet.Cells[18, 18] = model.sampleRFTInspection.CheckOuterCarton ? "V" : string.Empty;
            worksheet.Cells[19, 18] = model.sampleRFTInspection.CheckPackingMode ? "V" : string.Empty;
            worksheet.Cells[20, 18] = model.sampleRFTInspection.CheckPolytagMarketing ? "V" : string.Empty;
            worksheet.Cells[21, 18] = model.sampleRFTInspection.CheckHangtag ? "V" : string.Empty;


            int copyCnt = 0;
            if (model.AddDefect.ListDefectItem.Where(o => o.Qty > 0).Count() > 8)
            {
                copyCnt = model.AddDefect.ListDefectItem.Where(o => o.Qty > 0).Count() - 8;

                Excel.Range range = worksheet.get_Range($"A24:R24").EntireRow;

                for (int i = 0; i < copyCnt; i++)
                {
                    Microsoft.Office.Interop.Excel.Range rgX = worksheet.get_Range($"A25", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rgX.Insert(range.Copy(Type.Missing)); // 貼上
                }

            }

            int idx = 0;
            foreach (var item in model.AddDefect.ListDefectItem.Where(o => o.Qty > 0))
            {
                worksheet.Cells[24 + idx, 3] = item.GarmentDefectCodeID;
                worksheet.Cells[24 + idx, 4] = item.DefectCode;
                worksheet.Cells[24 + idx, 17] = item.Qty;
                idx++;
            }

            // 以下，Row數量要加上去
            worksheet.Cells[36 + copyCnt, 7] = model.sampleRFTInspection.BAQty;

            worksheet.Cells[40 + copyCnt, 7] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C1").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C1").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 8] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C2").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C2").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 9] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C3").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C3").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 10] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C4").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C4").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 11] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C5").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C5").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 12] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C6").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C6").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 13] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C7").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C7").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 14] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C8").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C8").FirstOrDefault().Qty : 0;
            worksheet.Cells[40 + copyCnt, 15] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C9").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C9").FirstOrDefault().Qty : 0;


            worksheet.Cells[46 + copyCnt, 6] = model.Setting.QCInCharge;

            // Pass / Fail  抓到圖形，並插入文字
            worksheet.Cells[46 + copyCnt, 6] = model.Setting.QCInCharge;

            if (model.sampleRFTInspection.Result == "Pass")
            {
                worksheet.Shapes.Item(1).TextFrame2.TextRange.Characters.Text = "V";
            }
            else if (model.sampleRFTInspection.Result == "Fail")
            {
                worksheet.Shapes.Item(3).TextFrame2.TextRange.Characters.Text = "V";
            }



            #region 存檔 > 讀取MemoryStream > 下載 > 刪除
            fileName = $"SampleRFTInspection_{DateTime.Now.ToString("yyyyMMddss")}.xlsx";
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            workbook.SaveAs(filepath);
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            if (IsDowdload)
            {
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
            }
            #endregion

            return fileName;
        }

        [HttpPost]
        [SessionAuthorize]
        public ActionResult SendMail()
        {
            BaseResult result = new BaseResult();
            string FileName = string.Empty;
            string reportPath = string.Empty;
            try
            {
                FileName = GetExcel(false);
                reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
                result.Result = true;
                result.ErrorMessage = string.Empty;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return Json(new { Result = result.Result, reportPath = reportPath, FileName = FileName });
        }
    }
}