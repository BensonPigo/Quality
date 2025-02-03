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
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class RandomTumblePillingTestService
    {
        private RandomTumblePillingTestProvider _Provider;
        private MailToolsService _MailService;
        public RandomTumblePillingTest_ViewModel GetDefaultModel(bool isNew = false)
        {
            RandomTumblePillingTest_ViewModel model = new RandomTumblePillingTest_ViewModel()
            {
                Request = new RandomTumblePillingTest_Request(),
                Main = new RandomTumblePillingTest_Main()
                {
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                TestStandard_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<RandomTumblePillingTest_Detail>(),
            };

            try
            {
                _Provider = new RandomTumblePillingTestProvider(Common.ProductionDataAccessLayer);
                //model.Scale_Source = _Provider.GetScales();

                if (isNew)
                {
                    // 預設每個Side 各三筆
                    foreach (var side in model.DefaultSide)
                    {
                        int idx = 1;
                        for (int i = 0; i <= 2; i++)
                        {
                            RandomTumblePillingTest_Detail nData = new RandomTumblePillingTest_Detail()
                            {
                                Side = side,
                                EvaluationItem = $"Specimen {idx}",
                                FirstScale = "1",
                                SecondScale = "1",
                                Result = "Fail",
                            };
                            model.DetailList.Add(nData);
                            idx++;
                        }
                    }
                }

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
        public RandomTumblePillingTest_ViewModel GetData(RandomTumblePillingTest_Request Req)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new RandomTumblePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<RandomTumblePillingTest_Main> tmpList = new List<RandomTumblePillingTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new RandomTumblePillingTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new RandomTumblePillingTest_Request()
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
                    _Provider = new RandomTumblePillingTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new RandomTumblePillingTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    string Subject = $"Random Tumble Pilling Test/{model.Main.OrderID}/" +
                         $"{model.Main.StyleID}/" +
                         $"{model.Main.FabricRefNo}/" +
                         $"{model.Main.FabricColor}/" +
                         $"{model.Main.Result}/" +
                         $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    model.Main.MailSubject = Subject;
                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new RandomTumblePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new RandomTumblePillingTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

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
        public RandomTumblePillingTest_ViewModel GetOrderInfo(string OrderID)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new RandomTumblePillingTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new RandomTumblePillingTest_Request() { OrderID = OrderID });


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
        public RandomTumblePillingTest_ViewModel NewSave(RandomTumblePillingTest_ViewModel Req, string MDivision, string UserID)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new RandomTumblePillingTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<RandomTumblePillingTest_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_RandomTumblePillingTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // RandomTumblePillingTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<RandomTumblePillingTest_Detail>();
                }

                _Provider.Processe_RandomTumblePillingTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new RandomTumblePillingTest_Request()
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
        public RandomTumblePillingTest_ViewModel EditSave(RandomTumblePillingTest_ViewModel Req, string UserID)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new RandomTumblePillingTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<RandomTumblePillingTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_RandomTumblePillingTest(Req, UserID);

                // 更新表身
                _Provider.Processe_RandomTumblePillingTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new RandomTumblePillingTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                    ReportNo = Req.Main.ReportNo,
                });

                model.Request.Article = model.Main.Article;

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
        public RandomTumblePillingTest_ViewModel Delete(RandomTumblePillingTest_ViewModel Req)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new RandomTumblePillingTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_RandomTumblePillingTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new RandomTumblePillingTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new RandomTumblePillingTest_Request()
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
        public RandomTumblePillingTest_ViewModel EncodeAmend(RandomTumblePillingTest_ViewModel Req, string UserID)
        {
            RandomTumblePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new RandomTumblePillingTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_RandomTumblePillingTest(Req.Main, UserID);

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
        public RandomTumblePillingTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            RandomTumblePillingTest_ViewModel result = new RandomTumblePillingTest_ViewModel();
            string tmpName = string.Empty;

            try
            {
                // 取得報表資料
                _Provider = new RandomTumblePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                RandomTumblePillingTest_ViewModel model = this.GetData(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });
                DataTable ReportTechnician = _Provider.GetReportTechnician(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });

                tmpName = $"Random Tumble Pilling Test_{model.Main.OrderID}_{model.Main.StyleID}_" +
                          $"{model.Main.FabricRefNo}_{model.Main.FabricColor}_{model.Main.Result}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[\/:*?""<>|]", ""); // 移除非法字元

                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templateFilePath = Path.Combine(baseFilePath, "XLT", "RandomTumblePillingTest.xltx");
                string fileName = $"{tmpName}.xlsx";
                string fullExcelFileName = Path.Combine(baseFilePath, "TMP", fileName);
                string filePdfName = $"{tmpName}.pdf";
                string fullPdfFileName = Path.Combine(baseFilePath, "TMP", filePdfName);

                // 開啟範本檔案進行操作
                using (var workbook = new XLWorkbook(templateFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 固定欄位填寫
                    worksheet.Cell(3, 2).Value = model.Main.ReportNo;
                    worksheet.Cell(3, 6).Value = model.Main.OrderID;

                    worksheet.Cell(4, 2).Value = model.Main.FactoryID;
                    worksheet.Cell(4, 6).Value = model.Main.SubmitDateText;

                    worksheet.Cell(5, 2).Value = model.Main.StyleID;
                    worksheet.Cell(5, 6).Value = model.Main.ReportDateText;

                    worksheet.Cell(6, 2).Value = model.Main.Article;
                    worksheet.Cell(6, 6).Value = model.Main.SeasonID;

                    worksheet.Cell(7, 2).Value = model.Main.SeasonID;
                    worksheet.Cell(7, 6).Value = model.Main.FabricColor;

                    // BrandID
                    if (model.Main.BrandID == "ADIDAS")
                    {
                        worksheet.Cell("B2").Value = "☑  ADIDAS";
                    }
                    if (model.Main.BrandID == "REEBOK")
                    {
                        worksheet.Cell("D2").Value = "☑  REEBOK";
                    }

                    // TestStandard
                    if (model.Main.TestStandard == "Fleece/Polar Fleece")
                    {
                        worksheet.Cell("B8").Value = "☑  Fleece/Polar Fleece";
                    }
                    if (model.Main.TestStandard == "French Terry")
                    {
                        worksheet.Cell("E8").Value = "☑  French Terry ";
                    }
                    if (model.Main.TestStandard == "Normal Fabric")
                    {
                        worksheet.Cell("F8").Value = "☑  Normal Fabric ";
                    }


                    // 結果標記
                    if (model.Main.Result == "Pass")
                        worksheet.Cell("I10").Value = "☑";
                    else
                        worksheet.Cell("I13").Value = "☑";


                    // 簽名圖片
                    if (ReportTechnician.Rows.Count > 0 && ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {
                        byte[] signature = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"];
                        AddImageToWorksheet(worksheet, signature, 48, 6, 100, 24);
                        worksheet.Cell(48, 5).Value = ReportTechnician.Rows[0]["Technician"]?.ToString();
                    }

                    // TestBeforePicture 和 TestAfterPicture
                    AddImageToWorksheet(worksheet, model.Main.TestFaceSideBeforePicture, 21, 1, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestFaceSideAfterPicture, 21, 6, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestBackSideBeforePicture, 34, 1, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestBackSideAfterPicture, 34, 6, 200, 300);

                    // 處理 DetailList
                    int rowIdx = 10;
                    foreach (var item in model.DetailList)
                    {
                        List<string> WorstScale = new List<string>();

                        // 每個 Side 開始填入時，先填 Worst Scale
                        if (rowIdx == 10 || rowIdx == 13)
                        {
                            var worstFirst = model.DetailList
                                                 .Where(o => o.Side == item.Side)
                                                 .OrderBy(o => o.FirstScale)
                                                 .FirstOrDefault()?.FirstScale;

                            var worstSecond = model.DetailList
                                                  .Where(o => o.Side == item.Side)
                                                  .OrderBy(o => o.SecondScale)
                                                  .FirstOrDefault()?.SecondScale;

                            if (!string.IsNullOrEmpty(worstFirst)) WorstScale.Add(worstFirst);
                            if (!string.IsNullOrEmpty(worstSecond)) WorstScale.Add(worstSecond);

                            worksheet.Cell(rowIdx, 7).Value = WorstScale.OrderBy(o => o).FirstOrDefault();
                        }

                        worksheet.Cell(rowIdx, 2).Value = item.EvaluationItem;
                        worksheet.Cell(rowIdx, 4).Value = item.FirstScale;
                        worksheet.Cell(rowIdx, 6).Value = item.SecondScale;

                        rowIdx++;
                    }

                    #region Title
                    string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(2).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A2:I2").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A2");
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

                // 轉換為 PDF（若需要）
                if (isPDF)
                {
                    LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    officeService.ConvertExcelToPdf(fullExcelFileName, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    result.TempFileName = filePdfName;
                }
                else
                {
                    result.TempFileName = fileName;
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }

            return result;
        }

        // 新增圖片的共用方法
        private void AddImageToWorksheet(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
        {
            if (imageData != null)
            {
                using (var stream = new MemoryStream(imageData))
                {
                    worksheet.AddPicture(stream)
                             .MoveTo(worksheet.Cell(row, col), 5, 5) // 微調圖片位置
                             .WithSize(width, height);
                }
            }
        }

        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new RandomTumblePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            RandomTumblePillingTest_ViewModel model = this.GetData(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });
            string name = $"Random Tumble Pilling Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            RandomTumblePillingTest_ViewModel report = this.GetReport(ReportNo, false , name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Random Tumble Pilling Test/{model.Main.OrderID}/" +
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
