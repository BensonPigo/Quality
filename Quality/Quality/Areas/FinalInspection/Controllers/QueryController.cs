

using BusinessLogicLayer;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.FinalInspection;
using FactoryDashBoardWeb.Helper;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Quality.Controllers;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class QueryController : BaseController
    {
        private QueryService Service = new QueryService();
        private IOrdersProvider _IOrdersProvider;
        public IFinalInspection_MeasurementProvider _FinalInspection_MeasurementProvider { get; set; }
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];

        public QueryController()
        {
            this.SelectedMenu = "Final Inspection";
            ViewBag.OnlineHelp = this.OnlineHelp + "FinalInspection.Query,,";
        }

        // GET: FinalInspection/Query
        public ActionResult Index()
        {

            List<string> inspectionlist = new List<string>() {
                "", "Pass","Fail","On-going"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;

            QueryFinalInspection_ViewModel model = new QueryFinalInspection_ViewModel();
            model.DataList = new List<QueryFinalInspection>();

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(QueryFinalInspection_ViewModel model)
        {
            model.DataList = Service.GetFinalinspectionQueryList(model);

            List<string> inspectionlist = new List<string>() {
                "", "Pass","Fail","On-going"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;



            return View(model);
        }

        public ActionResult Detail(string FinalInspectionID)
        {
            QueryReport model = Service.GetFinalInspectionReport(FinalInspectionID);

            TempData["Model"] = model;
            return View(model);
        }

        public ActionResult DownLoad()
        {
            bool test = false;
            if (!test)
            {
                if (TempData["Model"] == null)
                {
                    return RedirectToAction("Index");
                }
            }

            QueryReport model = (QueryReport)TempData["Model"];
            TempData["Model"] = model;

            // GetForFinalInspection 取得 SeasonID
            _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            FinalInspection_Request requestItem = new FinalInspection_Request() { SP = model.SP, CustPONO = model.FinalInspection.CustPONO, FactoryID = model.FinalInspection.FactoryID, StyleID = model.StyleID };
            IList<Orders> orders = _IOrdersProvider.GetOrderForInspection(requestItem);

            Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\FinalInspectionReport.xlsx");
            excelApp.DisplayAlerts = false;
            excelApp.Visible = test;
            Excel.Worksheet worksheet = excelApp.Sheets[1];

            // 先填入< 固定 > Cell資訊
            #region PO information / Inspection Setting
            worksheet.Cells[3, 3] = model.FinalInspection.FactoryID;
            worksheet.Cells[3, 7] = model.FinalInspection.InspectionStage;

            worksheet.Cells[4, 3] = model.FinalInspection.CustPONO;
            //worksheet.Cells[4, 7] = model.Carton;

            worksheet.Cells[5, 3] = model.SP;
            worksheet.Cells[5, 7] = model.FinalInspection.AuditDate.HasValue ? ((DateTime)model.FinalInspection.AuditDate).ToString("yyyy/MM/dd") : string.Empty;

            worksheet.Cells[6, 3] = model.StyleID;
            worksheet.Cells[6, 7] = model.AvailableQty;

            worksheet.Cells[7, 3] = orders.Select(s => s.SeasonID).FirstOrDefault();
            worksheet.Cells[7, 7] = model.AQLPlan;

            worksheet.Cells[8, 3] = model.BrandID;
            worksheet.Cells[8, 7] = (double)(model.FinalInspection.AcceptQty.HasValue ? model.FinalInspection.AcceptQty : 0);

            worksheet.Cells[9, 3] = model.TotalSPQty;
            worksheet.Cells[9, 7] = (double)(model.FinalInspection.AcceptQty.HasValue ? model.FinalInspection.AcceptQty + 1 : 1);
            #endregion

            #region General
            worksheet.Cells[13, 1] = "Fabric Approval: " + (model.FinalInspection.FabricApprovalDoc ? "Y" : "N");
            worksheet.Cells[13, 3] = "Sealing Sample: " + (model.FinalInspection.SealingSampleDoc ? "Y" : "N");
            worksheet.Cells[13, 5] = "Metal Detection: " + (model.FinalInspection.MetalDetectionDoc ? "Y" : "N");
            worksheet.Cells[13, 7] = "Garment Washing Test: " + (model.FinalInspection.GarmentWashingDoc ? "Y" : "N");
            #endregion

            #region Check List
            // Fabric Artwork Check List
            worksheet.Cells[16, 1] = "Close/ Shade: " + (model.FinalInspection.CheckCloseShade ? "Y" : "N");
            worksheet.Cells[16, 3] = "Handfeel: " + (model.FinalInspection.CheckHandfeel ? "Y" : "N");
            worksheet.Cells[16, 5] = "Appearance: " + (model.FinalInspection.CheckAppearance ? "Y" : "N");
            worksheet.Cells[16, 7] = "Print/ Emb Decorations: " + (model.FinalInspection.CheckPrintEmbDecorations ? "Y" : "N");

            // Label
            worksheet.Cells[18, 1] = "Fiber Content: " + (model.FinalInspection.CheckFiberContent ? "Y" : "N");
            worksheet.Cells[18, 3] = "Care Instructions: " + (model.FinalInspection.CheckCareInstructions ? "Y" : "N");
            worksheet.Cells[18, 5] = "Decorative Label: " + (model.FinalInspection.CheckDecorativeLabel ? "Y" : "N");
            worksheet.Cells[18, 7] = "Adicom Label: " + (model.FinalInspection.CheckAdicomLabel ? "Y" : "N");
            worksheet.Cells[19, 1] = "Country of Origion: " + (model.FinalInspection.CheckCountryofOrigion ? "Y" : "N");
            worksheet.Cells[19, 3] = "Size Key: " + (model.FinalInspection.CheckSizeKey ? "Y" : "N");
            worksheet.Cells[19, 5] = "8-Flag Label: " + (model.FinalInspection.Check8FlagLabel ? "Y" : "N");
            worksheet.Cells[19, 7] = "Additional Label: " + (model.FinalInspection.CheckAdditionalLabel ? "Y" : "N");

            // Packaging
            worksheet.Cells[21, 1] = "Shipping Mark: " + (model.FinalInspection.CheckShippingMark ? "Y" : "N");
            worksheet.Cells[21, 3] = "Polytag/ Marketing: " + (model.FinalInspection.CheckPolytagMarketing ? "Y" : "N");
            worksheet.Cells[21, 5] = "Color/ Size/ Qty: " + (model.FinalInspection.CheckColorSizeQty ? "Y" : "N");
            worksheet.Cells[21, 7] = "Hangtag: " + (model.FinalInspection.CheckHangtag ? "Y" : "N");
            #endregion

            #region Others
            worksheet.Cells[44, 3] = (double)model.FinalInspection.ProductionStatus * 0.01;
            worksheet.Cells[45, 3] = model.FinalInspection.OthersRemark;
            #endregion

            #region Result
            worksheet.Cells[48, 3] = model.FinalInspection.CFA;
            worksheet.Cells[48, 7] = (double)(model.FinalInspection.PassQty.HasValue ? model.FinalInspection.PassQty : 0);

            worksheet.Cells[49, 3] = model.FinalInspection.SubmitDate.HasValue ? ((DateTime)model.FinalInspection.SubmitDate).ToString("yyyy/MM/dd") : string.Empty;
            worksheet.Cells[49, 7] = (double)(model.FinalInspection.RejectQty.HasValue ? model.FinalInspection.RejectQty : 0);

            worksheet.Cells[50, 3] = model.FinalInspection.InspectionResult;
            worksheet.Cells[51, 3] = model.FinalInspection.ShipmentStatus;
            #endregion

            #region Measurement 依資料增加
            // 使用 FinalInspection.ID = FinalInspection_Measurement.ID 找到資料，直接展開 FinalInspection_Measurement
            _FinalInspection_MeasurementProvider = new FinalInspectionMeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<QueryReport_Measurement> MeasurementList = _FinalInspection_MeasurementProvider.GetQuery_FinalInspection_Measurement(model.FinalInspection.ID, model.SP).ToList();

            if (MeasurementList == null || MeasurementList.Count == 0)
            {
                for (int i = 0; i < 4; i++)
                {
                    worksheet.Rows[39].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                var groupList = MeasurementList.GroupBy(g => new { g.Time, g.Article, g.SizeCode, g.Location }).ToList();
                int copyCount = groupList.Count();
                Range rngToCopy = worksheet.get_Range("A39:A42").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A39", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                // 開始列39，預設4列
                for (int i = copyCount - 1; i >= 0; i--)
                {
                    int row = i * 4 + 39;
                    string time = groupList[i].Select(s => s.Time).First();
                    string article = groupList[i].Select(s => s.Article).First();
                    string sizeCode = groupList[i].Select(s => s.SizeCode).First();
                    string location = groupList[i].Select(s => s.Location).First();

                    var TimeList = MeasurementList.Where(w => w.Time == time && w.Article == article && w.SizeCode == sizeCode && w.Location == location).ToList();
                    worksheet.Cells[row, 1] = "Time: " + time;
                    worksheet.Cells[row + 1, 2] = article;
                    worksheet.Cells[row + 1, 4] = sizeCode;
                    worksheet.Cells[row + 1, 7] = location;

                    rngToCopy = worksheet.get_Range($"A{row + 3}").EntireRow; // 選取要被複製的資料
                    for (int j = 1; j < TimeList.Count; j++)
                    {
                        Excel.Range rngToInsert = worksheet.get_Range($"A{row + 3}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                        rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                    }

                    for (int j = 0; j < TimeList.Count; j++)
                    {
                        worksheet.Cells[row + 3 + j, 1] = TimeList[j].Description;
                        worksheet.Cells[row + 3 + j, 5] = TimeList[j].Tol2;
                        worksheet.Cells[row + 3 + j, 6] = TimeList[j].Tol1;
                        worksheet.Cells[row + 3 + j, 7] = TimeList[j].SizeSpec;
                        worksheet.Cells[row + 3 + j, 8] = TimeList[j].SizeSpec2;
                        worksheet.Cells[row + 3 + j, 9] = TimeList[j].diff;
                    }
                }
            }
            #endregion

            #region Moisture 依資料增加
            if (model.ListViewMoistureResult == null || model.ListViewMoistureResult.Count == 0)
            {
                for (int i = 0; i < 8; i++)
                {
                    worksheet.Rows[30].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                int copyCount = model.ListViewMoistureResult.Count();
                Range rngToCopy = worksheet.get_Range("A30:A37").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A30", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                // 開始列 30，預設 8 列
                for (int i = 0; i < copyCount; i++)
                {
                    int row = i * 8 + 30;
                    worksheet.Cells[row, 1] = $"Article: {model.ListViewMoistureResult[i].Article}, CTN: {model.ListViewMoistureResult[i].CTNNo}";
                    worksheet.Cells[row + 1, 3] = model.ListViewMoistureResult[i].Instrument;
                    worksheet.Cells[row + 1, 6] = model.ListViewMoistureResult[i].Fabrication;

                    worksheet.Cells[row + 2, 1] = $"Garment Moisture(Standard: {model.ListViewMoistureResult[i].GarmentStandard})";
                    worksheet.Cells[row + 3, 2] = model.ListViewMoistureResult[i].GarmentTop;
                    worksheet.Cells[row + 3, 4] = model.ListViewMoistureResult[i].GarmentMiddle;
                    worksheet.Cells[row + 3, 6] = model.ListViewMoistureResult[i].GarmentBottom;

                    worksheet.Cells[row + 4, 1] = $"Carton Moisture Standard({model.ListViewMoistureResult[i].CTNStandard})";
                    worksheet.Cells[row + 5, 2] = model.ListViewMoistureResult[i].CTNInside;
                    worksheet.Cells[row + 5, 4] = model.ListViewMoistureResult[i].CTNOutside;
                    worksheet.Cells[row + 6, 2] = model.ListViewMoistureResult[i].Result;
                    worksheet.Cells[row + 6, 4] = model.ListViewMoistureResult[i].Action;
                    worksheet.Cells[row + 7, 2] = model.ListViewMoistureResult[i].Remark;
                }
            }
            #endregion

            #region Beautiful Product Audit 依資料增加
            if (model.ListBACriteriaItem == null || model.ListBACriteriaItem.Count == 0)
            {
                for (int i = 0; i < 3; i++)
                {
                    worksheet.Rows[26].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                worksheet.Cells[26, 1] = $"Beautiful Product Qty: {model.FinalInspection.BAQty}";
                int copyCount = model.ListBACriteriaItem.Count;
                Range rngToCopy = worksheet.get_Range("A28").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A28", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                for (int i = 0; i < copyCount; i++)
                {
                    int row = i + 28;
                    worksheet.Cells[row, 1] = model.ListBACriteriaItem[i].BACriteria + ": " + model.ListBACriteriaItem[i].BACriteriaDesc;
                    worksheet.Cells[row, 9] = model.ListBACriteriaItem[i].Qty;
                }
            }
            #endregion
            #region Defect 依資料 依資料增加
            if (model.ListDefectItem == null || model.ListDefectItem.Count == 0)
            {
                for (int i = 0; i < 2; i++)
                {
                    worksheet.Rows[23].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                int copyCount = model.ListDefectItem.Count;
                Range rngToCopy = worksheet.get_Range("A24").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A24", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                for (int i = 0; i < copyCount; i++)
                {
                    int row = i + 24;
                    worksheet.Cells[row, 1] = model.ListDefectItem[i].DefectType;
                    worksheet.Cells[row, 3] = model.ListDefectItem[i].DefectCode;
                    worksheet.Cells[row, 9] = model.ListDefectItem[i].Qty;
                }
            }
            #endregion

            #region 存檔 > 讀取MemoryStream > 下載 > 刪除
            string fileName = $"FinalInspectionReport_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
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

            return null;
        }

        [HttpPost]
        public ActionResult SendMail()
        {
            bool test = IsTest.ToLower() == "true";

            QueryReport model = (QueryReport)TempData["Model"];
            TempData["Model"] = model;
            string WebHost = Request.Url.Scheme + @"://" + Request.Url.Authority + "/";
         
            var result = Service.SendMail(model.FinalInspection.ID, WebHost, test);

            return Json(result);
        }
    }
}