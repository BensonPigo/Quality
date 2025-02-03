

using BusinessLogicLayer;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ResultModel.EtoEFlowChart;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.FinalInspection;
using FactoryDashBoardWeb.Helper;
using Ict;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Quality.Controllers;
using Quality.Helper;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Http.Results;
using System.Web.Mvc;
using System.Web.Services.Description;
using Excel = Microsoft.Office.Interop.Excel;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class QueryController : BaseController
    {
        private QueryService Service = new QueryService();
        private IOrdersProvider _IOrdersProvider;
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
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
                "", "Pass","Fail","On-going","Junk"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;

            QueryFinalInspection_ViewModel model = new QueryFinalInspection_ViewModel()
            {
                ExcludeJunk = true,
            };
            model.DataList = Service.GetFinalinspectionQueryList_Default(model);

            return View(model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult Index(QueryFinalInspection_ViewModel model)
        {
            List<string> inspectionlist = new List<string>() {
                "", "Pass","Fail","On-going"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;

            if (string.IsNullOrEmpty(model.SP) && string.IsNullOrEmpty(model.CustPONO) && string.IsNullOrEmpty(model.StyleID) && (!model.AuditDateStart.HasValue || !model.AuditDateEnd.HasValue)
                && (!model.SubmitDateStart.HasValue || !model.SubmitDateEnd.HasValue))
            {
                model.ErrorMessage = $@"msg.WithError('SP# ,PO# ,Style# , Audit Date, Submit Date can not be empty.');";
                model.DataList = Service.GetFinalinspectionQueryList_Default(model);
                return View(model);
            }

            model.DataList = Service.GetFinalinspectionQueryList(model);

            return View(model);
        }

        public ActionResult Detail(string FinalInspectionID)
        {
            FinalInspectionService fService = new FinalInspectionService();
            QueryReport model = Service.GetFinalInspectionReport(FinalInspectionID);
            model.FinalInspection.GeneralList = fService.GetGeneralByBrand(FinalInspectionID, model.BrandID, model.FinalInspection.InspectionStage);
            model.FinalInspection.CheckListList = fService.GetCheckListByBrand(FinalInspectionID, model.BrandID);


            return View(model);
        }

        public ActionResult DownLoad(string FinalInspectionID)
        {
            bool test = false;
            if (!test)
            {
            }
            FinalInspectionService fService = new FinalInspectionService();

            QueryReport model = Service.GetFinalInspectionReport(FinalInspectionID);
            model.FinalInspection.GeneralList = fService.GetGeneralByBrand(FinalInspectionID, model.BrandID, model.FinalInspection.InspectionStage);
            model.FinalInspection.CheckListList = fService.GetCheckListByBrand(FinalInspectionID, model.BrandID);


            List<FinalInspectionBasicGeneral> AllGeneral = fService.GetAllGeneral();
            List<FinalInspectionBasicCheckList> AllCheckList = fService.GetAllCheckList();

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
            worksheet.Cells[4, 7] = model.Customize4;


            worksheet.Cells[5, 3] = model.SP;
            worksheet.Cells[5, 7] = model.ListCartonInfo != null && model.ListCartonInfo.Any() ? model.ListCartonInfo.Select(o => o.CTNNo).Distinct().JoinToString(",") : "";

            worksheet.Cells[6, 3] = model.StyleID;
            worksheet.Cells[6, 7] = model.FinalInspection.AuditDate.HasValue ? ((DateTime)model.FinalInspection.AuditDate).ToString("yyyy/MM/dd") : string.Empty;

            worksheet.Cells[7, 3] = model.SeasonID;
            worksheet.Cells[7, 7] = model.AvailableQty;

            worksheet.Cells[8, 3] = model.BrandID;
            worksheet.Cells[8, 7] = model.AQLPlan;

            worksheet.Cells[9, 3] = model.TotalSPQty;
            worksheet.Cells[9, 7] = (double)model.FinalInspection.AcceptQty;

            worksheet.Cells[10, 7] = (double)model.FinalInspection.AcceptQty + 1;
            #endregion

            #region Others
            worksheet.Cells[40, 3] = (double)model.FinalInspection.ProductionStatus * 0.01;
            worksheet.Cells[41, 3] = model.FinalInspection.OthersRemark;
            #endregion

            #region Result
            worksheet.Cells[44, 3] = model.FinalInspection.CFA;
            worksheet.Cells[44, 7] = (double)model.FinalInspection.PassQty;

            worksheet.Cells[45, 3] = model.FinalInspection.SubmitDate.HasValue ? ((DateTime)model.FinalInspection.SubmitDate).ToString("yyyy/MM/dd") : string.Empty;
            worksheet.Cells[45, 7] = (double)model.FinalInspection.RejectQty;

            worksheet.Cells[46, 3] = model.FinalInspection.InspectionResult;
            worksheet.Cells[47, 3] = model.FinalInspection.ShipmentStatus;
            #endregion

            #region Signatrue

            int signCtn = model.ListFinalInspectionSignature.Where(o => o.Signature != null).Count();

            int lasrRowIdx = 55;
            int bonusRowCtn = (signCtn / 3) - 2 + (signCtn % 3 > 0 ? 1 : 0);
            if (bonusRowCtn > 0)
            {
                Range rngToCopy = worksheet.get_Range("A50:A52").EntireRow; // 選取要被複製的資料
                for (int i = 0; i < bonusRowCtn; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A56", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                    lasrRowIdx += 3;
                }
            }

            Excel.PageSetup pageSetup = worksheet.PageSetup;
            // 设置Print Area
            pageSetup.PrintArea = $"A1:I{lasrRowIdx}";

            int RowIdx = 50;
            int ColIdx = 1;

            if (model.ListFinalInspectionSignature.Any(o => o.JobTitle == "QCManager"))
            {
                foreach (FinalInspectionSignature signatureData in model.ListFinalInspectionSignature.Where(o => o.JobTitle == "QCManager"))
                {
                    if (signatureData.Signature == null)
                    {
                        continue;
                    }

                    worksheet.Cells[RowIdx, ColIdx] = "QC CManager： " + signatureData.UserID + " - " + signatureData.UserName;
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[RowIdx + 1, ColIdx];

                    byte[] tmp = (byte[])signatureData.Signature; // 圖片的 byte[]

                    MemoryStream ms = new MemoryStream(tmp);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    if (ColIdx >= 7)
                    {
                        ColIdx = 1;
                        RowIdx += 3;
                    }
                    else
                    {
                        ColIdx += 3;
                    }
                }

            }
            if (model.ListFinalInspectionSignature.Any(o => o.JobTitle == "ProductionManager"))
            {
                worksheet.Cells[RowIdx, ColIdx] = "Production Manager";

                foreach (FinalInspectionSignature signatureData in model.ListFinalInspectionSignature.Where(o => o.JobTitle == "ProductionManager"))
                {
                    if (signatureData.Signature == null)
                    {
                        continue;
                    }
                    worksheet.Cells[RowIdx, ColIdx] = "Production CManager： " + signatureData.UserID + " - " + signatureData.UserName;
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[RowIdx + 1, ColIdx];

                    byte[] tmp = (byte[])signatureData.Signature; // 圖片的 byte[]

                    MemoryStream ms = new MemoryStream(tmp);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    if (ColIdx >= 7)
                    {
                        ColIdx = 1;
                        RowIdx += 3;
                    }
                    else
                    {
                        ColIdx += 3;
                    }
                }
            }
            if (model.ListFinalInspectionSignature.Any(o => o.JobTitle == "TSD"))
            {
                foreach (FinalInspectionSignature signatureData in model.ListFinalInspectionSignature.Where(o => o.JobTitle == "TSD"))
                {
                    if (signatureData.Signature == null)
                    {
                        continue;
                    }
                    worksheet.Cells[RowIdx, ColIdx] = "TSD： " + signatureData.UserID + " - " + signatureData.UserName;
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[RowIdx + 1, ColIdx];

                    byte[] tmp = (byte[])signatureData.Signature; // 圖片的 byte[]

                    MemoryStream ms = new MemoryStream(tmp);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    if (ColIdx >= 7)
                    {
                        ColIdx = 1;
                        RowIdx += 3;
                    }
                    else
                    {
                        ColIdx += 3;
                    }
                }
            }
            if (model.ListFinalInspectionSignature.Any(o => o.JobTitle == "QC"))
            {
                foreach (FinalInspectionSignature signatureData in model.ListFinalInspectionSignature.Where(o => o.JobTitle == "QC"))
                {
                    if (signatureData.Signature == null)
                    {
                        continue;
                    }
                    worksheet.Cells[RowIdx, ColIdx] = "QC： " + signatureData.UserID + " - " + signatureData.UserName;
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[RowIdx + 1, ColIdx];

                    byte[] tmp = (byte[])signatureData.Signature; // 圖片的 byte[]

                    MemoryStream ms = new MemoryStream(tmp);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    if (ColIdx >= 7)
                    {
                        ColIdx = 1;
                        RowIdx += 3;
                    }
                    else
                    {
                        ColIdx += 3;
                    }
                }
            }
            if (model.ListFinalInspectionSignature.Any(o => o.JobTitle == "Production"))
            {
                foreach (FinalInspectionSignature signatureData in model.ListFinalInspectionSignature.Where(o => o.JobTitle == "Production"))
                {
                    if (signatureData.Signature == null)
                    {
                        continue;
                    }
                    worksheet.Cells[RowIdx, ColIdx] = "Production： " + signatureData.UserID + " - " +signatureData.UserName;
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[RowIdx + 1, ColIdx];

                    byte[] tmp = (byte[])signatureData.Signature; // 圖片的 byte[]

                    MemoryStream ms = new MemoryStream(tmp);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    if (ColIdx >= 7)
                    {
                        ColIdx = 1;
                        RowIdx += 3;
                    }
                    else
                    {
                        ColIdx += 3;
                    }
                }
            }
            #endregion

            #region Measurement 依資料增加
            // 使用 FinalInspection.ID = FinalInspection_Measurement.ID 找到資料，直接展開 FinalInspection_Measurement
            _FinalInspection_MeasurementProvider = new FinalInspectionMeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<QueryReport_Measurement> MeasurementList = _FinalInspection_MeasurementProvider.GetQuery_FinalInspection_Measurement(model.FinalInspection.ID, model.SP).ToList();

            if (MeasurementList == null || MeasurementList.Count == 0)
            {
                for (int i = 0; i < 4; i++)
                {
                    worksheet.Rows[35].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                var groupList = MeasurementList.GroupBy(g => new { g.Time, g.Article, g.SizeCode, g.Location }).ToList();
                int copyCount = groupList.Count();
                Range rngToCopy = worksheet.get_Range("A35:A38").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A35", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                // 開始列41，預設4列
                for (int i = copyCount - 1; i >= 0; i--)
                {
                    int row = i * 4 + 35;
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
                    worksheet.Rows[26].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                int copyCount = model.ListViewMoistureResult.Count();
                Range rngToCopy = worksheet.get_Range("A26:A33").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A26", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                // 開始列 31，預設 8 列
                for (int i = 0; i < copyCount; i++)
                {
                    int row = i * 8 + 26;
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
                    worksheet.Rows[22].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                worksheet.Cells[22, 1] = $"Beautiful Product Qty: {model.FinalInspection.BAQty}";
                int copyCount = model.ListBACriteriaItem.Count;
                Range rngToCopy = worksheet.get_Range("A24").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A24", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                for (int i = 0; i < copyCount; i++)
                {
                    int row = i + 24;
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
                    worksheet.Rows[19].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
            else
            {
                int copyCount = model.ListDefectItem.Count;
                Range rngToCopy = worksheet.get_Range("A20").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < copyCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A20", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                for (int i = 0; i < copyCount; i++)
                {
                    int row = i + 20;
                    worksheet.Cells[row, 1] = model.ListDefectItem[i].DefectTypeDesc;
                    worksheet.Cells[row, 3] = model.ListDefectItem[i].DefectCodeDesc;
                    worksheet.Cells[row, 9] = model.ListDefectItem[i].Qty;
                }
            }
            #endregion

            #region Check List 依資料增加

            if (model.FinalInspection.CheckListListDic.Count > 0)
            {
                var CheckListType = AllCheckList.Select(o => o.Type).Distinct().ToList();
                int CheckListRowCount = CheckListType.Count;

                foreach (var t in CheckListType)
                {
                    var sameType = AllCheckList.Where(o => o.Type == t);
                    int typCount = (sameType.Count() / 4) + 1;
                    CheckListRowCount += typCount;
                }

                // 根據Type數量複製Row
                Range rngToCopy = worksheet.get_Range("A16:A17").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < CheckListType.Count; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A16", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                int eachTypeRowIdx = 16;
                foreach (var type in CheckListType)
                {
                    // 填入Check List的Type
                    worksheet.Cells[eachTypeRowIdx, 1] = type;

                    var sameTypeData = model.FinalInspection.CheckListList.Where(o => o.Type == type).ToList();
                    int typeRowCount = (sameTypeData.Count % 4) > 0 ? (sameTypeData.Count / 4) + 1 : (sameTypeData.Count / 4);


                    // 根據這個Type有多少 CheckList 數量複製Row
                    Range rngToCopy2 = worksheet.get_Range($@"A{eachTypeRowIdx + 1}:A{eachTypeRowIdx + 1}").EntireRow; // 選取要被複製的資料
                    for (int i = 1; i < typeRowCount; i++)
                    {
                        Excel.Range rngToInsert = worksheet.get_Range($"A{eachTypeRowIdx + 1}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                        rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy2.Copy(Type.Missing)); // 貼上
                    }


                    int columnCount = 0;
                    int checkRowCount = 0;
                    foreach (var checkItem in sameTypeData)
                    {
                        string isSelect = model.FinalInspection.CheckListListDic[checkItem.CheckListColName] ? "Y" : "N";

                        int col = 1 + (columnCount * 2);
                        worksheet.Cells[eachTypeRowIdx + 1 + checkRowCount, col] = $@"{checkItem.ItemName}: {isSelect}";
                        columnCount++;

                        // 一個Row只有四筆資料，超過換行
                        if (columnCount > 3)
                        {
                            columnCount = 0;
                            checkRowCount += 1;
                        }
                    }

                    // 每個Type底下有幾Row的Check List
                    eachTypeRowIdx += typeRowCount;
                    // 換一個Type
                    eachTypeRowIdx++;
                }

            }

            #endregion

            #region General 依資料增加

            if (model.FinalInspection.GeneralDic.Count > 0)
            {
                int GeneralRowCount = (AllGeneral.Count % 4) > 0 ? (AllGeneral.Count / 4) + 1 : (AllGeneral.Count / 4);

                Range rngToCopy = worksheet.get_Range("A14:A14").EntireRow; // 選取要被複製的資料
                for (int i = 1; i < GeneralRowCount; i++)
                {
                    Excel.Range rngToInsert = worksheet.get_Range("A14", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上
                }

                int columnCount = 0;
                GeneralRowCount = 0;
                foreach (var g in AllGeneral)
                {
                    int col = 1 + (columnCount * 2);
                    string isSelect = model.FinalInspection.GeneralDic[g.GeneralColName] ? "Y" : "N";

                    worksheet.Cells[14 + GeneralRowCount, col] = $@"{g.ItemName}: {isSelect}";
                    columnCount++;

                    // 一個Row只有四筆資料，超過換行
                    if (columnCount > 3)
                    {
                        columnCount = 0;
                        GeneralRowCount += 1;
                    }
                }
            }
            #endregion

            #region Title
            string FactoryNameEN = fService.GetFactoryNameEN(model.SP, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
            worksheet.Rows["1"].Insert();
            // 1. 合併欄位 (B1:K1)
            Microsoft.Office.Interop.Excel.Range cellA1 = worksheet.Range["A1", "I1"];
            cellA1.Merge();
            // 設置字體樣式

            // 2. 設置文字和樣式
            cellA1.Value = FactoryNameEN; // 替換為你的 FactoryNameEN 變數
            cellA1.Font.Name = "Calibri";      // 設置字體類型
            cellA1.Font.Size = 18;          // 設置字體大小
            cellA1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // 水平置中
            cellA1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;   // 垂直置中
            cellA1.Font.Bold = true;        // 設置字體加粗

            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            if (usedRange != null)
            {
                // 獲取最後一行
                int lastRow = usedRange.Row + usedRange.Rows.Count - 1;

                // 清除已有的列印範圍
                worksheet.PageSetup.PrintArea = string.Empty;

                // 設定新的列印範圍
                string printArea = $"A1:I{lastRow}";
                worksheet.PageSetup.PrintArea = printArea;
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

        public ActionResult DownloadTable(QueryFinalInspection_ViewModel model)
        {
            this.CheckSession();
            Report_Result report_Result = Service.QueryReport(model);
            string tempFilePath = report_Result.TempFileName;
            tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;
            if (!report_Result.Result)
            {
                report_Result.ErrorMessage = report_Result.ErrorMessage.ToString();
            }

            return Json(new { Result = report_Result.Result, ErrorMessage = report_Result.ErrorMessage, reportPath = tempFilePath, FileName = report_Result.TempFileName });
        }

        [HttpPost]
        public ActionResult SendMail(string FinalInspectionID)
        {
            bool test = IsTest.ToLower() == "true";

            string WebHost = Request.Url.Scheme + @"://" + Request.Url.Authority + "/";

            var result = Service.SendMail(FinalInspectionID, WebHost, test);

            return Json(result);
        }

        [HttpPost]
        public ActionResult ClickJunk(string ID)
        {
            DatabaseObject.BaseResult result = Service.ClickJunk(ID);
            return Json(result);
        }

        public JsonResult GetSignatureUserList(string userIDs)
        {
            try
            {
                userIDs = userIDs == null ? string.Empty : userIDs;
                FinalInspectionService fService = new FinalInspectionService();
                List<FinalInspectionSignature> list = Service.GetSignatureUserList(userIDs);

                return new JsonResult
                {
                    Data = list,
                    MaxJsonLength = int.MaxValue,/*重點在這行*/
                    JsonRequestBehavior = JsonRequestBehavior.AllowGet
                };
            }
            catch (Exception ex)
            {
                return Json(new { Result = false, ErrorMessage = ex.Message });
            }

        }

        [HttpPost]
        public JsonResult PostSignatureUserList(string FinalInspectionID, string JobTitle, List<FinalInspectionSignature> allData)
        {
            try
            {
                if (allData == null)
                {
                    allData = new List<FinalInspectionSignature>();
                }
                var result = Service.InsertFinalInspectionSignatureUser(FinalInspectionID, JobTitle, allData);

                if (result.Result)
                {
                    return Json(new { Result = true, });
                }
                else
                {
                    return Json(new { Result = false, ErrorMessage = result.ErrorMessage });
                }
            }
            catch (Exception ex)
            {
                return Json(new { Result = false, ErrorMessage = ex.Message });
            }

        }

        [HttpPost]
        public JsonResult GetSignature(FinalInspectionSignature Req)
        {
            try
            {
                QueryReport model = Service.GetFinalInspectionSignature(Req);

                return Json(new { Result = true, Data = model.ListFinalInspectionSignature.FirstOrDefault() });

            }
            catch (Exception ex)
            {
                return Json(new { Result = false, ErrorMessage = ex.Message });
            }

        }
        [HttpPost]
        public JsonResult InsertSignature(FinalInspectionSignature Req)
        {
            try
            {
                Req.Signature = ImageHelper.ImageCompress(Req.Signature);
                Req.AddName = this.UserID;
                QueryReport model = Service.InsertFinalInspectionSignature(Req);


                if (model.Result)
                {
                    return Json(new { Result = true, Data = model.ListFinalInspectionSignature.FirstOrDefault() });
                }
                else
                {
                    return Json(new { Result = false, ErrorMessage = model.ErrorMessage });
                }

            }
            catch (Exception ex)
            {
                return Json(new { Result = false, ErrorMessage = ex.Message });
            }

        }
    }
}