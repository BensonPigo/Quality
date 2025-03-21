using ADOHelper.Utility;
using ClosedXML.Excel;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using static Sci.MyUtility;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class SalivaFastnessTestService
    {
        private SalivaFastnessTestProvider _Provider;
        private MailToolsService _MailService;
        public SalivaFastnessTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            SalivaFastnessTest_ViewModel model = new SalivaFastnessTest_ViewModel()
            {
                Request = new SalivaFastnessTest_Request(),
                Main = new SalivaFastnessTest_Main()
                {
                    // ISP20230288 寫死4
                    Time = 4,
                    Temperature = 37,
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                Scale_Source = new List<System.Web.Mvc.SelectListItem>(),
                Temperature_Source = new List<System.Web.Mvc.SelectListItem>()
                {
                    // ISP20230288 寫死37
                    new System.Web.Mvc.SelectListItem(){Text="37",Value="37"},
                    //new System.Web.Mvc.SelectListItem(){Text="0",Value="0"},
                    //new System.Web.Mvc.SelectListItem(){Text="30",Value="30"},
                    //new System.Web.Mvc.SelectListItem(){Text="40",Value="40"},
                    //new System.Web.Mvc.SelectListItem(){Text="50",Value="50"},
                    //new System.Web.Mvc.SelectListItem(){Text="60",Value="60"},
                },
                DetailList = new List<SalivaFastnessTest_Detail>(),
            };

            try
            {
                _Provider = new SalivaFastnessTestProvider(Common.ProductionDataAccessLayer);
                model.Scale_Source = _Provider.GetScales();

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
        public SalivaFastnessTest_ViewModel GetData(SalivaFastnessTest_Request Req)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new SalivaFastnessTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SalivaFastnessTest_Main> tmpList = new List<SalivaFastnessTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new SalivaFastnessTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new SalivaFastnessTest_Request()
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
                    _Provider = new SalivaFastnessTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new SalivaFastnessTest_Request() { OrderID = model.Main.OrderID });

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

                    _Provider = new SalivaFastnessTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new SalivaFastnessTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Saliva Fastness Test/{model.Main.OrderID}/" +
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
        public SalivaFastnessTest_ViewModel GetOrderInfo(string OrderID)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new SalivaFastnessTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new SalivaFastnessTest_Request() { OrderID = OrderID });


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
        public SalivaFastnessTest_ViewModel NewSave(SalivaFastnessTest_ViewModel Req, string MDivision, string UserID)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new SalivaFastnessTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<SalivaFastnessTest_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_SalivaFastnessTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // SalivaFastnessTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<SalivaFastnessTest_Detail>();
                }

                _Provider.Processe_SalivaFastnessTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new SalivaFastnessTest_Request()
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
        public SalivaFastnessTest_ViewModel EditSave(SalivaFastnessTest_ViewModel Req, string UserID)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new SalivaFastnessTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<SalivaFastnessTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_SalivaFastnessTest(Req, UserID);

                // 更新表身
                _Provider.Processe_SalivaFastnessTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new SalivaFastnessTest_Request()
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
        public SalivaFastnessTest_ViewModel Delete(SalivaFastnessTest_ViewModel Req)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new SalivaFastnessTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_SalivaFastnessTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new SalivaFastnessTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new SalivaFastnessTest_Request()
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
        public SalivaFastnessTest_ViewModel EncodeAmend(SalivaFastnessTest_ViewModel Req, string UserID)
        {
            SalivaFastnessTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new SalivaFastnessTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_SalivaFastnessTest(Req.Main, UserID);

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

        public SalivaFastnessTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            SalivaFastnessTest_ViewModel result = new SalivaFastnessTest_ViewModel();
            string basefileName = "SalivaFastnessTest";
            string tmpName = string.Empty;

            try
            {
                // 取得報表資料
                _Provider = new SalivaFastnessTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                SalivaFastnessTest_ViewModel model = this.GetData(new SalivaFastnessTest_Request() { ReportNo = ReportNo });
                DataTable ReportTechnician = _Provider.GetReportTechnician(new SalivaFastnessTest_Request() { ReportNo = ReportNo });

                tmpName = $"Saliva Fastness Test_{model.Main.OrderID}_" +
                          $"{model.Main.StyleID}_{model.Main.FabricRefNo}_{model.Main.FabricColor}_" +
                          $"{model.Main.Result}_{DateTime.Now:yyyyMMddHHmmss}";

                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string fileName = $"{tmpName}.xlsx";
                string fullExcelFileName = Path.Combine(baseFilePath, "TMP", fileName);
                string filePdfName = $"{tmpName}.pdf";
                string fullPdfFileName = Path.Combine(baseFilePath, "TMP", filePdfName);

                string templatePath = Path.Combine(baseFilePath, "XLT", "SalivaFastnessTest.xltx");
                // 建立 Excel 檔案
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入基本資料
                    worksheet.Cell("B3").Value = model.Main.ReportNo;
                    worksheet.Cell(3, 6).Value = model.Main.OrderID;

                    worksheet.Cell(4, 2).Value = model.Main.FactoryID;
                    worksheet.Cell(4, 6).Value = model.Main.SubmitDateText;

                    worksheet.Cell(5, 2).Value = model.Main.StyleID;
                    worksheet.Cell(5, 6).Value = model.Main.ReportDateText;

                    worksheet.Cell(6, 2).Value = model.Main.Article;
                    worksheet.Cell(7, 2).Value = model.Main.SeasonID;
                    worksheet.Cell(8, 2).Value = model.Main.FabricRefNo;
                    worksheet.Cell(9, 2).Value = model.Main.FabricColor;
                    worksheet.Cell(10, 2).Value = model.Main.FabricDescription;

                    worksheet.Cell(11, 2).Value = model.Main.TypeOfPrint;
                    worksheet.Cell(11, 6).Value = model.Main.PrintColor;
                    
                    // BrandID
                    if (model.Main.BrandID.ToUpper() == "ADIDAS")
                    {
                        worksheet.Cell("B2").Value = "☑  ADIDAS";
                    }
                    if (model.Main.BrandID.ToUpper() == "REEBOK")
                    {
                        worksheet.Cell("D2").Value = "☑  REEBOK";
                    }


                    // ITEM TESTED
                    if (model.Main.ItemTested.ToUpper() == "FABRIC")
                    {
                        worksheet.Cell("E7").Value = "FABRIC:  ☑";
                    }
                    if (model.Main.ItemTested.ToUpper() == "ACCESSORIES")
                    {
                        worksheet.Cell("F7").Value = "ACCESSORIES:  ☑";
                    }
                    if (model.Main.ItemTested.ToUpper() == "PRINTING")
                    {
                        worksheet.Cell("G7").Value = "PRINTING:  ☑";
                    }


                    // 簽名圖片
                    if (ReportTechnician.Rows.Count > 0 && ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {
                        byte[] signature = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"];
                        using (var stream = new MemoryStream(signature))
                        {
                            var image = worksheet.AddPicture(stream)
                                .MoveTo(worksheet.Cell(37, 6), 5, 5)
                                .WithSize(100, 24);
                        }
                        worksheet.Cell(37, 5).Value = ReportTechnician.Rows[0]["Technician"]?.ToString();
                    }

                    // TestBeforePicture、TestAfterPicture 圖片
                    if (model.Main.TestBeforePicture != null && model.Main.TestBeforePicture.Length > 1)
                    {
                        AddImageToWorksheet(worksheet, model.Main.TestBeforePicture, 23, 1, 400, 400);
                    }
                    if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                    {
                        AddImageToWorksheet(worksheet, model.Main.TestAfterPicture, 23, 6, 400, 400);
                    }
                    // 處理 DetailList
                    int startRow = 14;
                    foreach (var detail in model.DetailList)
                    {

                        if (startRow == 14)
                        {
                            // 1. 複製格式
                            var sourceRange = worksheet.Range("A14:I19"); // 複製的範圍（6 行模板）

                            // 2. 插入
                            worksheet.Row(startRow).InsertRowsAbove(6);

                            // 3. 複製格式 貼上
                            var destinationRange = worksheet.Range($"A{startRow}:I{startRow + 5}");
                            sourceRange.CopyTo(destinationRange);

                            // 4. 手動複製行高
                            for (int i = 0; i < sourceRange.RowCount(); i++)
                            {
                                var sourceRow = sourceRange.FirstRow().RowNumber() + i;
                                var targetRow = startRow + i;

                                worksheet.Row(targetRow).Height = worksheet.Row(sourceRow).Height; // 複製行高
                            }
                        }

                        worksheet.Cell(startRow, 4).Value = detail.AcetateScale;
                        worksheet.Cell(startRow, 6).Value = detail.AcetateResult;
                        worksheet.Cell(startRow + 1, 4).Value = detail.CottonScale;
                        worksheet.Cell(startRow + 1, 6).Value = detail.CottonResult;
                        worksheet.Cell(startRow + 2, 4).Value = detail.NylonScale;
                        worksheet.Cell(startRow + 2, 6).Value = detail.NylonResult;

                        worksheet.Cell(startRow + 3, 4).Value = detail.PolyesterScale;
                        worksheet.Cell(startRow + 3, 6).Value = detail.PolyesterResult;
                        worksheet.Cell(startRow + 4, 4).Value = detail.AcrylicScale;
                        worksheet.Cell(startRow + 4, 6).Value = detail.AcrylicResult;
                        worksheet.Cell(startRow + 5, 4).Value = detail.WoolScale;
                        worksheet.Cell(startRow + 5, 6).Value = detail.WoolResult;

                        // 結果標記
                        if (detail.AllResult == "Pass")
                        {
                            worksheet.Cell(startRow, 8).Value = "☑";
                        }
                        else
                        {
                            worksheet.Cell(startRow + 3, 8).Value = "☑";
                        }
                        startRow += 6; // 移動到下一組
                    }

                    #region Title
                    string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:I1").Merge();
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
                        worksheet.PageSetup.PrintAreas.Add($"A1:I{lastRow + 10}");
                    }
                    #endregion
                    // 儲存 Excel 檔案
                    workbook.SaveAs(fullExcelFileName);
                }
                result.TempFileName = fileName;

                //// 若需要轉 PDF
                //if (isPDF)
                //{
                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //    //officeService.ConvertExcelToPdf(fullExcelFileName, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //    ConvertToPDF.ExcelToPDF(fullExcelFileName, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName));
                //    result.TempFileName = filePdfName;
                //}

                result.Result = true;
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
            _Provider = new SalivaFastnessTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            SalivaFastnessTest_ViewModel model = this.GetData(new SalivaFastnessTest_Request() { ReportNo = ReportNo });
            string name = $"Saliva Fastness Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_ " +
                $"{model.Main.FabricColor}_ " +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            SalivaFastnessTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Saliva Fastness Test/{model.Main.OrderID}/" +
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
