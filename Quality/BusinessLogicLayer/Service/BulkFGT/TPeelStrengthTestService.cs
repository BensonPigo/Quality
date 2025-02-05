using ADOHelper.Utility;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Web;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class TPeelStrengthTestService
    {
        private string UploadFileRootPath = ConfigurationManager.AppSettings["UploadFileRootPath"];
        private string UploadFilePath = $@"{ConfigurationManager.AppSettings["UploadFileRootPath"]}BulkFGT\TPeelStrengthTest\";
        private TPeelStrengthTestProvider _Provider;
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public TPeelStrengthTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            TPeelStrengthTest_ViewModel model = new TPeelStrengthTest_ViewModel()
            {
                Request = new TPeelStrengthTest_Request(),
                Main = new TPeelStrengthTest_Main()
                {
                    // ISP20230288 寫死4
                    MachineReport = "",

                },
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<TPeelStrengthTest_Detail>(),

                MachineNo_Source = new List<System.Web.Mvc.SelectListItem>()
                {
                    // ISP20230288 寫死37
                    new System.Web.Mvc.SelectListItem(){Text="Fabric",Value="Fabric"},
                    new System.Web.Mvc.SelectListItem(){Text="Accessory",Value="Accessory"},
                    new System.Web.Mvc.SelectListItem(){Text="Printing",Value="Printing"},
                },
            };


            if (!System.IO.Directory.Exists($@"{UploadFileRootPath}\BulkFGT\"))
            {
                System.IO.Directory.CreateDirectory($@"{UploadFileRootPath}\BulkFGT\");
            }

            if (!System.IO.Directory.Exists($@"{UploadFileRootPath}\BulkFGT\TPeelStrengthTest"))
            {
                System.IO.Directory.CreateDirectory($@"{UploadFileRootPath}\BulkFGT\TPeelStrengthTest");
            }

            return model;
        }

        /// <summary>
        /// 取得查詢結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public TPeelStrengthTest_ViewModel GetData(TPeelStrengthTest_Request Req)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<TPeelStrengthTest_Main> tmpList = new List<TPeelStrengthTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new TPeelStrengthTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new TPeelStrengthTest_Request()
                    {
                        BrandID = Req.BrandID,
                        SeasonID = Req.SeasonID,
                        StyleID = Req.StyleID,
                        Article = Req.Article,
                    });
                }


                if (tmpList.Any())
                {
                    // 取得 ReportNo下拉選單
                    foreach (var item in tmpList)
                    {
                        model.ReportNo_Source.Add(new System.Web.Mvc.SelectListItem()
                        {
                            Text = item.ReportNo,
                            Value = item.ReportNo,
                        });
                    }

                    // 取得表身
                    // 若"有"傳入ReportNo，則可以直接找出表頭表身明細
                    if (!string.IsNullOrEmpty(Req.ReportNo) && tmpList.Where(o => o.ReportNo == Req.ReportNo).Any())
                    {
                        // 取得表頭資料
                        model.Main = tmpList.Where(o => o.ReportNo == Req.ReportNo).FirstOrDefault();
                    }
                    // 若"沒"傳入ReportNo，抓下拉選單第一筆顯示
                    else
                    {
                        // 根據下拉選單第一筆，取得表頭資料
                        model.Main = tmpList.Where(o => o.ReportNo == model.ReportNo_Source.FirstOrDefault().Value).FirstOrDefault();

                    }

                    // 取得Article 下拉選單
                    _Provider = new TPeelStrengthTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new TPeelStrengthTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new TPeelStrengthTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"T-Peel Strength Test/{model.Main.OrderID}/" +
                        $"{model.Main.StyleID}/" +
                        $"{model.Main.FabricRefNo}/" +
                        $"{model.Main.FabricColor}/" +
                        $"{model.Main.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                    model.Main.MailSubject = Subject;

                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "Data not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public TPeelStrengthTest_ViewModel GetOrderInfo(string OrderID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new TPeelStrengthTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new TPeelStrengthTest_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 取得表頭SP#相關欄位
                    model.Main.OrderID = tmpOrders.FirstOrDefault().ID;
                    model.Main.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                    model.Main.BrandID = tmpOrders.FirstOrDefault().BrandID;
                    model.Main.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                    model.Main.StyleID = tmpOrders.FirstOrDefault().StyleID;

                    // 取得Article 下拉選單
                    foreach (var oriData in tmpOrders)
                    {
                        System.Web.Mvc.SelectListItem Article = new System.Web.Mvc.SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "SP# not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public TPeelStrengthTest_ViewModel NewSave(TPeelStrengthTest_ViewModel Req, string MDivision, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                if (Req.Main.MachineReportFile != null)
                {
                    this.CreateMachineReportFile(Req);
                }

                // 新增，並取得ReportNo
                _Provider.Insert_TPeelStrengthTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // TPeelStrengthTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                _Provider.Processe_TPeelStrengthTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                    ReportNo = Req.Main.ReportNo,
                });

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }

        public TPeelStrengthTest_ViewModel EditSave(TPeelStrengthTest_ViewModel Req, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                if (Req.Main.MachineReportFile != null)
                {
                    this.CreateMachineReportFile(Req);
                }

                // 更新表頭
                _Provider.Update_TPeelStrengthTest(Req, UserID);

                // 更新表身
                _Provider.Processe_TPeelStrengthTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                    ReportNo = Req.Main.ReportNo,
                });

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        public TPeelStrengthTest_ViewModel Delete(TPeelStrengthTest_ViewModel Req)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_TPeelStrengthTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new TPeelStrengthTest_Request()
                {
                    ReportNo = model.Main.ReportNo,
                    BrandID = model.Main.BrandID,
                    SeasonID = model.Main.SeasonID,
                    StyleID = model.Main.StyleID,
                    Article = model.Main.Article,
                };

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        public TPeelStrengthTest_ViewModel EncodeAmend(TPeelStrengthTest_ViewModel Req, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_TPeelStrengthTest(Req.Main, UserID);

                _ISQLDataTransaction.Commit();

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        private void AddImageToWorksheet(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
        {
            if (imageData != null)
            {
                using (var stream = new MemoryStream(imageData))
                {
                    worksheet.AddPicture(stream)
                             .MoveTo(worksheet.Cell(row, col), 5, 5)
                             .WithSize(width, height);
                }
            }
        }
        public TPeelStrengthTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            TPeelStrengthTest_ViewModel result = new TPeelStrengthTest_ViewModel();

            try
            {
                // 初始化變數
                string baseFileName = "TPeelStrengthTest";
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templatePath = Path.Combine(baseFilePath, "XLT", $"{baseFileName}.xltx");
                string tmpName;

                // 取得報表資料
                _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
                TPeelStrengthTest_ViewModel model = this.GetData(new TPeelStrengthTest_Request { ReportNo = ReportNo });
                DataTable reportTechnician = _Provider.GetReportTechnician(new TPeelStrengthTest_Request { ReportNo = ReportNo });
                var testCode = _QualityBrandTestCodeProvider.Get(model.Main.BrandID, "T-Peel Strength Test");

                // 處理檔案名稱
                tmpName = $"T-Peel Strength Test_{model.Main.OrderID}_{model.Main.StyleID}_" +
                          $"{model.Main.FabricRefNo}_{model.Main.FabricColor}_{model.Main.Result}_{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[/:?""<>|*%]", string.Empty);
                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                string outputDirectory = Path.Combine(baseFilePath, "TMP");
                string filePath = Path.Combine(outputDirectory, $"{tmpName}.xlsx");
                string pdfPath = Path.Combine(outputDirectory, $"{tmpName}.pdf");

                // 開啟模板並填寫資料
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填寫標題
                    if (testCode.Any())
                    {
                        worksheet.Cell("A1").Value = $"T-peel strength test({testCode.FirstOrDefault().TestCode})";
                    }

                    worksheet.Cell("B3").Value = model.Main.ReportNo;
                    worksheet.Cell("B4").Value = model.Main.SubmitDateText;
                    worksheet.Cell("E4").Value = model.Main.ReportDateText;
                    worksheet.Cell("B5").Value = model.Main.OrderID;
                    worksheet.Cell("E5").Value = model.Main.BrandID;
                    worksheet.Cell("B6").Value = model.Main.StyleID;
                    worksheet.Cell("E6").Value = model.Main.SeasonID;
                    worksheet.Cell("B7").Value = model.Main.Article;
                    worksheet.Cell("E7").Value = model.Main.MachineNo;
                    worksheet.Cell("B8").Value = model.Main.FabricRefNo;
                    worksheet.Cell("E8").Value = model.Main.FabricColor;

                    // Technician 資料處理
                    if (reportTechnician.Rows.Count > 0)
                    {
                        worksheet.Cell("F20").Value = reportTechnician.Rows[0]["Technician"]?.ToString();
                        AddImageToWorksheet(worksheet, reportTechnician.Rows[0]["TechnicianSignture"] as byte[], 19, 6, 100, 24);
                    }

                    // TestAfterPicture 圖片
                    AddImageToWorksheet(worksheet, model.Main.TestAfterPicture, 17, 2, 200, 300);

                    // 表身處理
                    if (model.DetailList.Any())
                    {
                        string allRemark = string.Join(Environment.NewLine, model.DetailList.Select(o => o.Remark));
                        worksheet.Cell("B14").Value = allRemark;

                        // 複製欄位
                        int copyCount = model.DetailList.Count - 1;
                        for (int i = 0; i < copyCount; i++)
                        {
                            // 1. 複製第 12 列
                            var rowToCopy = worksheet.Row(12);

                            // 2. 插入一列，將第 12 和第 13 列之間騰出空間
                            worksheet.Row(13).InsertRowsAbove(1);

                            // 3. 複製內容與格式到新插入的第 13 列
                            var newRow = worksheet.Row(13);
                            rowToCopy.CopyTo(newRow);
                        }

                        int rowIdx = 12;
                        foreach (var detailData in model.DetailList)
                        {
                            worksheet.Cell(rowIdx, 1).Value = $"Wash {detailData.EvaluationItem}";
                            worksheet.Cell(rowIdx, 3).Value = detailData.WarpValue;
                            worksheet.Cell(rowIdx, 4).Value = detailData.WarpResult;
                            worksheet.Cell(rowIdx, 5).Value = detailData.WeftValue;
                            worksheet.Cell(rowIdx, 6).Value = detailData.WeftResult;
                            rowIdx++;
                        }
                    }
                    #region Title
                    string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:F1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    mergedCell.Style.Font.Italic = false;

                    // 自動檢測使用範圍
                    var usedRange = worksheet.RangeUsed();
                    var lastRow = worksheet.CellsUsed().Max(cell => cell.Address.RowNumber);
                    // 確認範圍不為空
                    if (usedRange != null)
                    {
                        // 清除所有已有的列印範圍
                        worksheet.PageSetup.PrintAreas.Clear();

                        // 設定列印範圍為使用範圍
                        worksheet.PageSetup.PrintAreas.Add($"A1:F{lastRow + 10}");
                    }
                    #endregion
                    // 儲存 Excel 檔案
                    workbook.SaveAs(filePath);
                }

                // 轉 PDF
                if (isPDF)
                {
                    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    //officeService.ConvertExcelToPdf(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    ConvertToPDF.ExcelToPDF(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", pdfPath));
                    result.TempFileName = $"{tmpName}.pdf";
                }
                else
                {
                    result.TempFileName = $"{tmpName}.xlsx";
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        // 定義一個方法來將頁面從一個 PDF 複製到另一個 PDF
        public void CopyPages(PdfReader sourcePdf, Document destinationPdf, PdfWriter writer)
        {
            PdfContentByte contentByte = writer.DirectContent;
            for (int pageIndex = 1; pageIndex <= sourcePdf.NumberOfPages; pageIndex++)
            {
                PdfImportedPage importedPage = writer.GetImportedPage(sourcePdf, pageIndex);
                destinationPdf.NewPage();
                contentByte.AddTemplate(importedPage, 0, 0);
            }
        }

        public void CreateMachineReportFile(TPeelStrengthTest_ViewModel Req)
        {
            Req.Main.MachineReport = Req.Main.MachineReportFile.FileName;
            //string FileName= Req.Main.MachineReportFile.FileName;
            string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(Req.Main.MachineReportFile.FileName);
            string FileExtension = Path.GetExtension(Req.Main.MachineReportFile.FileName);
            string FullFileName = $@"{this.UploadFilePath}\{Req.Main.MachineReport}";

            int count = 1;
            string FileName = FileNameWithoutExtension;
            while (File.Exists(FullFileName))
            {
                FileName = $"{FileNameWithoutExtension} ({count})";
                FullFileName = Path.Combine(this.UploadFilePath, $"{FileName}{FileExtension}");
                count++;
            }

            using (var fileStream = File.Create(FullFileName))
            {
                Req.Main.MachineReportFile.InputStream.Seek(0, SeekOrigin.Begin);
                Req.Main.MachineReportFile.InputStream.CopyTo(fileStream);
            }

            Req.Main.MachineReport = FileName + FileExtension;
        }
        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            TPeelStrengthTest_ViewModel model = this.GetData(new TPeelStrengthTest_Request() { ReportNo = ReportNo });
            string name = $"T-Peel Strength Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            TPeelStrengthTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"T-Peel Strength Test/{model.Main.OrderID}/" +
                    $"{model.Main.StyleID}/" +
                    $"{model.Main.FabricRefNo}/" +
                    $"{model.Main.FabricColor}/" +
                    $"{model.Main.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                To = TO,
                CC = CC,
                Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                FileUploader = Files,
                IsShowAIComment = true,
            };

            if (!string.IsNullOrEmpty(Subject))
            {
                sendMail_Request.Subject = Subject;
            }

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
            sendMail_Request.Body = Body + Environment.NewLine + sendMail_Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
