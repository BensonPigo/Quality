using ADOHelper.Utility;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.SampleRFT;
using FactoryDashBoardWeb.Helper;
using Ionic.Zip;
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
        private CFTCommentsService _CFTCommentsService;
        public InspBySPQueryController()
        {
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.InspBySPQuery,,";
            _Service = new InspectionBySPService();
            _CFTCommentsService = new CFTCommentsService();
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


        public ActionResult DownLoad(long ID)
        {
            QueryReport model = _Service.GetQueryDetail(ID, this.UserID);
            string ExcelFileName = GetExcel(true, model);

            string CFTExcelFileName = GetCFTComment(true, model);

            List<string> dummyFitPicList = GetDummtPic(true, model);

            ZipFile zip = new ZipFile();
            string zipName = $@"Query Report_{DateTime.Now.ToString("yyyyMMddHHmmss")}.zip";
            string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);

            zip.AddFile(ExcelFileName, string.Empty);
            zip.AddFile(CFTExcelFileName, string.Empty);
            foreach (var dummyFileName in dummyFitPicList)
            {
                zip.AddFile(dummyFileName, string.Empty);
            }
            zip.Save(zipPath);

            MemoryStream obj_stream = new MemoryStream();
            var tempFile = System.IO.Path.Combine(zipName);
            obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(zipPath));
            Response.AddHeader("Content-Disposition", $"attachment; filename={zipName}");
            Response.BinaryWrite(obj_stream.ToArray());
            obj_stream.Close();
            obj_stream.Dispose();
            Response.Flush();
            Response.End();

            return null;
        }

        public string GetExcel(bool IsDowdload, QueryReport model)
        {
            string fileName = string.Empty;
            TempData["ModelQuery"] = model;

            Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\InspBySPQuery_Detail.xlsx");
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Excel.Worksheet worksheet = excelApp.Sheets[1];

            worksheet.Cells[4, 3] = model.sampleRFTInspection.InspectionTimesText;

            worksheet.Cells[6, 6] = "Sintex";
            worksheet.Cells[6, 10] = model.sampleRFTInspection.FactoryID;
            worksheet.Cells[6, 15] = model.sampleRFTInspection.CFTNameText;

            worksheet.Cells[7, 6] = model.Setting.Model;

            worksheet.Cells[8, 6] = model.sampleRFTInspection.WorkNo;

            // Sample Type
            worksheet.Cells[9, 6] = model.Setting.SampleStage;

            worksheet.Cells[10, 6] = model.sampleRFTInspection.POID;
            worksheet.Cells[10, 14] = model.sampleRFTInspection.OrderID;

            worksheet.Cells[11, 6] = model.Setting.Article;

            worksheet.Cells[12, 6] = model.Setting.OrderQty;
            worksheet.Cells[12, 14] = model.sampleRFTInspection.SewingLineID + (string.IsNullOrEmpty(model.sampleRFTInspection.SewingLine2ndID) ? string.Empty : $", {model.sampleRFTInspection.SewingLine2ndID}");

            worksheet.Cells[13, 6] = model.Setting.CustPONo;

            worksheet.Cells[14, 6] = model.sampleRFTInspection.BrandFTYCode;

            if (model.Setting.InspectionStage == "100%")
            {
                worksheet.Cells[15, 3] = "100%";
            }
            worksheet.Cells[15, 6] = model.Setting.AQLPlan;
            worksheet.Cells[15, 10] = model.Setting.InspectionStage == "100%" ? model.Setting.OrderQty : model.sampleRFTInspection.SampleSize;
            worksheet.Cells[15, 13] = model.sampleRFTInspection.AcceptQty;
            worksheet.Cells[15, 17] = model.Setting.InspectionStage == "100%" ? 1 : model.sampleRFTInspection.AcceptQty + 1;

            worksheet.Cells[16, 6] = model.sampleRFTInspection.CheckFabricApproval ? "V" : string.Empty;

            worksheet.Cells[16, 12] = model.sampleRFTInspection.CheckSealingSampleApproval ? "V" : string.Empty;

            worksheet.Cells[17, 6] = model.sampleRFTInspection.CheckMetalDetection ? "V" : string.Empty;

            worksheet.Cells[19, 6] = model.sampleRFTInspection.CheckColorShade ? "V" : string.Empty;
            worksheet.Cells[20, 6] = model.sampleRFTInspection.CheckAppearance ? "V" : string.Empty;
            worksheet.Cells[21, 6] = model.sampleRFTInspection.CheckHandfeel ? "V" : string.Empty;
            worksheet.Cells[22, 6] = model.sampleRFTInspection.CheckEMB ? "V" : string.Empty;
            worksheet.Cells[23, 6] = model.sampleRFTInspection.CheckPrintEmbDecorations ? "V" : string.Empty;
            worksheet.Cells[24, 6] = model.sampleRFTInspection.CheckHT ? "V" : string.Empty;
            worksheet.Cells[25, 6] = model.sampleRFTInspection.CheckBadge ? "V" : string.Empty;

            worksheet.Cells[19, 10] = model.sampleRFTInspection.CheckFiberContent ? "V" : string.Empty;
            worksheet.Cells[20, 10] = model.sampleRFTInspection.CheckCountryofOrigin ? "V" : string.Empty;
            worksheet.Cells[21, 10] = model.sampleRFTInspection.CheckCareInstructions ? "V" : string.Empty;
            worksheet.Cells[22, 10] = model.sampleRFTInspection.CheckSizeKey ? "V" : string.Empty;

            worksheet.Cells[19, 14] = model.sampleRFTInspection.CheckDecorativeLabel ? "V" : string.Empty;
            worksheet.Cells[20, 14] = model.sampleRFTInspection.CheckCareLabel ? "V" : string.Empty;
            worksheet.Cells[22, 14] = model.sampleRFTInspection.CheckSecurityLabel ? "V" : string.Empty;
            worksheet.Cells[22, 14] = model.sampleRFTInspection.CheckAdditionalLabel ? "V" : string.Empty;

            worksheet.Cells[19, 18] = model.sampleRFTInspection.CheckOuterCarton ? "V" : string.Empty;
            worksheet.Cells[20, 18] = model.sampleRFTInspection.CheckPackingMode ? "V" : string.Empty;
            worksheet.Cells[21, 18] = model.sampleRFTInspection.CheckPolytagMarketing ? "V" : string.Empty;
            worksheet.Cells[22, 18] = model.sampleRFTInspection.CheckHangtag ? "V" : string.Empty;


            int copyCnt = 0;
            if (model.AddDefect.ListDefectItem.Where(o => o.Qty > 0).Count() > 8)
            {
                copyCnt = model.AddDefect.ListDefectItem.Where(o => o.Qty > 0).Count() - 8;

                Excel.Range range = worksheet.get_Range($"A28:R28").EntireRow;

                for (int i = 0; i < copyCnt; i++)
                {
                    Microsoft.Office.Interop.Excel.Range rgX = worksheet.get_Range($"A29", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rgX.Insert(range.Copy(Type.Missing)); // 貼上
                }

            }

            int idx = 0;
            foreach (var item in model.AddDefect.ListDefectItem.Where(o => o.Qty > 0))
            {
                worksheet.Cells[28 + idx, 3] = item.GarmentDefectCodeID;
                worksheet.Cells[28 + idx, 4] = item.DefectCode;
                worksheet.Cells[28 + idx, 17] = item.Qty;
                idx++;
            }

            // 以下，Row數量要加上去
            worksheet.Cells[40 + copyCnt, 7] = model.sampleRFTInspection.BAQty;

            worksheet.Cells[44 + copyCnt, 7] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C1").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C1").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 8] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C2").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C2").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 9] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C3").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C3").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 10] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C4").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C4").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 11] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C5").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C5").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 12] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C6").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C6").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 13] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C7").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C7").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 14] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C8").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C8").FirstOrDefault().Qty : 0;
            worksheet.Cells[44 + copyCnt, 15] = model.BA.ListBACriteria.Where(o => o.BACriteria == "C9").Any() ? model.BA.ListBACriteria.Where(o => o.BACriteria == "C9").FirstOrDefault().Qty : 0;


            worksheet.Cells[47 + copyCnt, 6] = model.Setting.QCInCharge;

            // Pass / Fail  抓到圖形，並插入文字
            worksheet.Cells[50 + copyCnt, 6] = model.Setting.QCInCharge;

            worksheet.Cells[55 + copyCnt, 10] = model.sampleRFTInspection.SubmitDate.HasValue ? model.sampleRFTInspection.SubmitDate.Value.ToShortDateString() : string.Empty;

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

            //if (IsDowdload)
            //{
            //    MemoryStream obj_stream = new MemoryStream();
            //    var tempFile = System.IO.Path.Combine(filepath);
            //    obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFile));
            //    Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            //    Response.BinaryWrite(obj_stream.ToArray());
            //    obj_stream.Close();
            //    obj_stream.Dispose();
            //    Response.Flush();
            //    Response.End();
            //    System.IO.File.Delete(filepath);
            //}
            #endregion

            fileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);
            return fileName;
        }

        public string GetCFTComment(bool IsDowdload, QueryReport model)
        {
            string fileName = string.Empty;
            TempData["ModelQuery"] = model;

            CFTComments_ViewModel qModel = new CFTComments_ViewModel();

            qModel = _CFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID = model.sampleRFTInspection.OrderID, QueryType = "OrderID" });
            qModel.ReleasedBy = this.UserID;

            CFTComments_ViewModel result = _CFTCommentsService.GetExcel2(qModel);
            fileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.TempFileName);


            return fileName;
        }

        public List<string> GetDummtPic(bool IsDowdload, QueryReport model)
        {
            List<string> resultPathList = new List<string>();
            TempData["ModelQuery"] = model;

            var detailList = model.DummyFit.DetailList;

            foreach (var detail in detailList)
            {
                string frontImgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", detail.FrontImageName);
                string backImgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", detail.BackImageName);
                string leftImgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", detail.LeftImageName);
                string rightImgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", detail.RightImageName);

                if (detail.Front != null && detail.Front.Length > 0)
                {
                    using (var imageFile = new FileStream(frontImgPath, FileMode.Create))
                    {
                        imageFile.Write(detail.Front, 0, detail.Front.Length);
                        imageFile.Flush();
                    }
                    resultPathList.Add(frontImgPath);
                }
                if (detail.Back != null && detail.Back.Length > 0)
                {
                    using (var imageFile = new FileStream(backImgPath, FileMode.Create))
                    {
                        imageFile.Write(detail.Back, 0, detail.Back.Length);
                        imageFile.Flush();
                    }
                    resultPathList.Add(backImgPath);
                }
                if (detail.Left != null && detail.Left.Length > 0)
                {
                    using (var imageFile = new FileStream(leftImgPath, FileMode.Create))
                    {
                        imageFile.Write(detail.Left, 0, detail.Left.Length);
                        imageFile.Flush();
                    }
                    resultPathList.Add(leftImgPath);
                }
                if (detail.Right != null && detail.Right.Length > 0)
                {
                    using (var imageFile = new FileStream(rightImgPath, FileMode.Create))
                    {
                        imageFile.Write(detail.Right, 0, detail.Right.Length);
                        imageFile.Flush();
                    }
                    resultPathList.Add(rightImgPath);
                }
            }


            return resultPathList;
        }

        [HttpPost]
        [SessionAuthorize]
        public ActionResult SendMail(long ID)
        {
            BaseResult result = new BaseResult();            
            string FileName = string.Empty;
            string zipName = $@"Query Report_{DateTime.Now.ToString("yyyyMMddHHmmss")}.zip";
            string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);
            string reportPath = string.Empty;
            try
            {
                QueryReport model = _Service.GetQueryDetail(ID, this.UserID);
                string ExcelFileName = GetExcel(true, model);
                string CFTExcelFileName = GetCFTComment(true, model);
                List<string> dummyFitPicList = GetDummtPic(true, model);

                ZipFile zip = new ZipFile();

                zip.AddFile(ExcelFileName, string.Empty);
                zip.AddFile(CFTExcelFileName, string.Empty);
                foreach (var dummyFileName in dummyFitPicList)
                {
                    zip.AddFile(dummyFileName, string.Empty);
                }
                zip.Save(zipPath);


                reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + zipName;
                result.Result = true;
                result.ErrorMessage = string.Empty;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            TempData["InspBySPQueryID"] = ID;
            return Json(new { Result = result.Result, reportPath = reportPath, FileName = zipName });
        }

        public ActionResult SendMailer(string TO, string CC, string subject, string body, string file)
        {
            SendMail_Request request = new SendMail_Request()
            {
                To = TO,
                CC = CC,
                Subject = subject,
                Body = body,
                FileonServer = new List<string>() { file }
            };

            ViewBag.JS = "";
            ViewBag.InspBySPQueryID = TempData["InspBySPQueryID"];


            return View(request);
        }

        [HttpPost]
        public ActionResult SendMailer(SendMail_Request _Request, long InspBySPQueryID)
        {
            QueryReport model = _Service.GetQueryDetail(InspBySPQueryID, this.UserID);
            var defectData = model.AddDefect.ListDefectItem.Where(o => o.Qty > 0);

            string defectHtml = $@"";
            foreach (var item in defectData)
            {
                defectHtml += $@"
<tr>
<td style=""width: 16.6667%;"">{item.DefectTypeDesc}</td>
<td style=""width: 16.6667%;"">{item.DefectCodeDesc}</td>
<td style=""width: 16.6667%;"">{item.AreaCodes}</td>
<td style=""width: 16.6667%;"">{item.Qty}</td>
<td style=""width: 16.6667%;"">{item.Responsibility}</td>
<td style=""width: 16.6667%;"">{item.AIComment}</td>
</tr>
";
            }

            string html = $@"
<table style=""border-collapse: collapse; width: 100%;"" border=""1"">
<tbody>
<tr>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">Defect Type</td>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">Defect Code</td>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">Defect Area</td>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">Defect Qty&nbsp;</td>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">Reponsibility</td>
<td style=""width: 16.6667%; background-color: lightgray; text-align: center;"">AI Comment</td>
</tr>
{defectHtml}
</tbody>
</table>
";
            _Request.Body = _Request.Body + "</br>" + html;

            SendMail_Result result = MailTools.SendMail(_Request);

            string js = "";

            js += "<script src='/ThirdParty/SciCustom/js/jquery-3.4.1.min.js'></script> ";
            js += "<script>  $(function () { ";
            if (result.result)
            {
                js += "alert('Success'); ";
            }
            else
            {
                js += "alert('" + result.resultMsg + "'); ";
            }

            js += "window.close(); ";
            js += " }); </script>";

            return Content(js, "text/html");
        }
    }
}