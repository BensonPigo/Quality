using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class WickingHeightTestService
    {
        private WickingHeightTestProvider _Provider;
        private MailToolsService _MailService;

        public WickingHeightTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            WickingHeightTest_ViewModel model = new WickingHeightTest_ViewModel()
            {
                Request = new WickingHeightTest_Request(),
                Main = new WickingHeightTest_Main(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<WickingHeightTest_Detail>() 
                { 
                    new WickingHeightTest_Detail() { EvaluationType = "Before Wash" },
                    new WickingHeightTest_Detail() { EvaluationType = "After Wash" },
                },
                DetaiItemlList = new List<WickingHeightTest_Detail_Item>()
                {
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 1", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 2", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 3", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 1", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 2", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 3", WarpTime = 30, WeftTime = 30 }
                },
            };

            try
            {
                _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }
            return model;
        }

        /// <summary>
        /// 取得查詢結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public WickingHeightTest_ViewModel GetData(WickingHeightTest_Request Req)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<WickingHeightTest_Main> tmpList = new List<WickingHeightTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WickingHeightTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WickingHeightTest_Request()
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
                    if (!string.IsNullOrEmpty(Req.ReportNo))
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
                    _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new WickingHeightTest_Request() { OrderID = model.Main.OrderID });

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new WickingHeightTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    model.DetaiItemlList = _Provider.GetDetaiItemlList(new WickingHeightTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Wicking Height Test/{model.Main.OrderID}/" +
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

        public WickingHeightTest_ViewModel GetOrderInfo(string OrderID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new WickingHeightTest_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 取得表頭SP#相關欄位
                    model.Main.OrderID = tmpOrders.FirstOrDefault().ID;
                    model.Main.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                    model.Main.BrandID = tmpOrders.FirstOrDefault().BrandID;
                    model.Main.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                    model.Main.StyleID = tmpOrders.FirstOrDefault().StyleID;
                    model.Main.Article = tmpOrders.FirstOrDefault().Article;
                    model.Main.FabricType = tmpOrders.FirstOrDefault().FabricType;
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

        public WickingHeightTest_ViewModel NewSave(WickingHeightTest_ViewModel Req, string MDivision, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.WarpResult == "Fail" || o.WeftResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_WickingHeightTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // WickingHeightTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }

                // WickingHeightTest_Detail 新增 or 修改
                if (Req.DetaiItemlList == null)
                {
                    Req.DetaiItemlList = new List<WickingHeightTest_Detail_Item>();
                }

                _Provider.Processe_WickingHeightTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel EditSave(WickingHeightTest_ViewModel Req, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.WarpResult == "Fail" || o.WeftResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_WickingHeightTest(Req, UserID);

                // 更新表身
                _Provider.Processe_WickingHeightTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel Delete(WickingHeightTest_ViewModel Req)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_WickingHeightTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel EncodeAmend(WickingHeightTest_ViewModel Req, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_WickingHeightTest(Req.Main, UserID);

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

        public WickingHeightTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            WickingHeightTest_ViewModel result = new WickingHeightTest_ViewModel();
            _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                string baseFileName = "WickingHeightTest";
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templatePath = Path.Combine(baseFilePath, "XLT", $"{baseFileName}.xltx");

                // 取得報表資料
                WickingHeightTest_ViewModel model = this.GetData(new WickingHeightTest_Request { ReportNo = ReportNo });
                DataTable reportTechnician = _Provider.GetReportTechnician(new WickingHeightTest_Request { ReportNo = ReportNo });

                // 生成檔案名稱
                string tmpName = $"Wicking Height Test_{model.Main.OrderID}_{model.Main.StyleID}_{model.Main.FabricRefNo}_" +
                                 $"{model.Main.FabricColor}_{model.Main.Result}_{DateTime.Now:yyyyMMddHHmmss}";
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                string outputPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填寫主要報表內容
                    worksheet.Cell("C2").Value = model.Main.ReportNo?.ToString();
                    worksheet.Cell("F2").Value = model.Main.ReportDate?.ToString("yyyy/MM/dd");

                    worksheet.Cell("C3").Value = model.Main.SubmitDate?.ToString("yyyy/MM/dd");
                    worksheet.Cell("F3").Value = model.Main.BrandID?.ToString();

                    worksheet.Cell("C4").Value = model.Main.SeasonID?.ToString();
                    worksheet.Cell("F4").Value = model.Main.OrderID?.ToString();

                    worksheet.Cell("C5").Value = model.Main.StyleID?.ToString();
                    worksheet.Cell("F5").Value = "Bulk";

                    worksheet.Cell("C6").Value = model.Main.Article?.ToString();
                    worksheet.Cell("F6").Value = model.Main.FabricType?.ToString();

                    worksheet.Cell("B7").Value = model.Main.Result?.ToString();
                    worksheet.Cell("E7").Value = model.Main.FabricDescription?.ToString();

                    // Technician 資料與簽名圖片
                    if (reportTechnician.Rows.Count > 0)
                    {
                        worksheet.Cell("C25").Value = reportTechnician.Rows[0]["Technician"]?.ToString();
                        AddImageToWorksheet(worksheet, reportTechnician.Rows[0]["TechnicianSignture"] as byte[], 26, 3, 100, 24);
                    }

                    // TestWarpPicture 圖片
                    AddImageToWorksheet(worksheet, model.Main.TestWarpPicture, 23, 1, 200, 300);

                    // TestWeftPicture 圖片
                    AddImageToWorksheet(worksheet, model.Main.TestWeftPicture, 23, 4, 200, 300);

                    // 表身處理
                    if (model.DetailList.Any() && model.DetailList.Count >= 1)
                    {
                        var detailBefore = model.DetailList.FirstOrDefault(x => x.EvaluationType == "Before Wash");
                        var detailAfter = model.DetailList.FirstOrDefault(x => x.EvaluationType == "After Wash");

                        worksheet.Cell("B15").Value = detailBefore?.Remark?.ToString();
                        worksheet.Cell("B20").Value = detailAfter?.Remark?.ToString();

                        // Before Wash 資料
                        long detailUkey = detailBefore.Ukey;
                        int i = 0;
                        foreach (var item in model.DetaiItemlList.Where(x => x.WickingHeightTestDetailUkey == detailUkey).OrderBy(x => x.Ukey))
                        {
                            worksheet.Cell(14, 2 + i).Value = $"{item.WarpValues}mm/{item.WarpTime}min";
                            worksheet.Cell(14, 5 + i).Value = $"{item.WeftValues}mm/{item.WeftTime}min";
                            i++;
                        }

                        // After Wash 資料
                        detailUkey = detailAfter.Ukey;
                        i = 0;
                        foreach (var item in model.DetaiItemlList.Where(x => x.WickingHeightTestDetailUkey == detailUkey).OrderBy(x => x.Ukey))
                        {
                            worksheet.Cell(19, 2 + i).Value = $"{item.WarpValues}mm/{item.WarpTime}min";
                            worksheet.Cell(19, 5 + i).Value = $"{item.WeftValues}mm/{item.WeftTime}min";
                            i++;
                        }
                    }
                    #region Title
                    string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:G1").Merge();
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
                        worksheet.PageSetup.PrintAreas.Add($"A1:G{lastRow + 10}");
                    }
                    #endregion


                    // 儲存 Excel 檔案
                    workbook.SaveAs(outputPath);
                }

                result.TempFileName = $"{tmpName}.xlsx";
                result.Result = true;

                //// 轉換為 PDF
                //if (isPDF)
                //{
                //    //string pdfPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.pdf");

                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //    //officeService.ConvertExcelToPdf(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //    ConvertToPDF.ExcelToPDF(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{tmpName}.pdf"));

                //    result.TempFileName = $"{tmpName}.pdf";
                //}
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }

            return result;
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
        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            WickingHeightTest_ViewModel model = this.GetData(new WickingHeightTest_Request() { ReportNo = ReportNo });
            string name = $"Wicking Height Test_{model.Main.OrderID}_" +
                    $"{model.Main.StyleID}_" +
                    $"{model.Main.FabricRefNo}_" +
                    $"{model.Main.FabricColor}_" +
                    $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            WickingHeightTest_ViewModel report = this.GetReport(ReportNo, false,AssignedFineName : name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Wicking Height Test/{model.Main.OrderID}/" +
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
