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
    public class WaterAbsorbencyService
    {
        private WaterAbsorbencyProvider _Provider;
        private MailToolsService _MailService;
        public WaterAbsorbency_ViewModel GetDefaultModel(bool iNew = false)
        {
            WaterAbsorbency_ViewModel model = new WaterAbsorbency_ViewModel()
            {
                Request = new WaterAbsorbency_Request(),
                Main = new WaterAbsorbency_Main(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<WaterAbsorbency_Detail>()
                {
                    new WaterAbsorbency_Detail() { EvaluationItem = "Before Wash" },
                    new WaterAbsorbency_Detail() { EvaluationItem = "After 5 Cycle Wash" },
                },
            };

            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }
            return model;
        }

        public WaterAbsorbency_ViewModel GetData(WaterAbsorbency_Request Req)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<WaterAbsorbency_Main> tmpList = new List<WaterAbsorbency_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WaterAbsorbency_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WaterAbsorbency_Request()
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
                    _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new WaterAbsorbency_Request() { OrderID = model.Main.OrderID });

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new WaterAbsorbency_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Water Absorbency Test/{model.Main.OrderID}/" +
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

        public WaterAbsorbency_ViewModel GetOrderInfo(string OrderID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new WaterAbsorbency_Request() { OrderID = OrderID });


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

        public WaterAbsorbency_ViewModel NewSave(WaterAbsorbency_ViewModel Req, string MDivision, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_WaterAbsorbency(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // WaterAbsorbency_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }

                _Provider.Processe_WaterAbsorbency_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
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
        public WaterAbsorbency_ViewModel EditSave(WaterAbsorbency_ViewModel Req, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_WaterAbsorbency(Req, UserID);

                // 更新表身
                _Provider.Processe_WaterAbsorbency_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
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
        public WaterAbsorbency_ViewModel Delete(WaterAbsorbency_ViewModel Req)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_WaterAbsorbency(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new WaterAbsorbency_Request()
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
        public WaterAbsorbency_ViewModel EncodeAmend(WaterAbsorbency_ViewModel Req, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_WaterAbsorbency(Req.Main, UserID);

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
        public WaterAbsorbency_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            WaterAbsorbency_ViewModel result = new WaterAbsorbency_ViewModel();
            _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                string baseFileName = "WaterAbsorbencyTest";
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templatePath = Path.Combine(baseFilePath, "XLT", $"{baseFileName}.xltx");

                // 取得報表資料
                WaterAbsorbency_ViewModel model = this.GetData(new WaterAbsorbency_Request { ReportNo = ReportNo });
                DataTable reportTechnician = _Provider.GetReportTechnician(new WaterAbsorbency_Request { ReportNo = ReportNo });

                // 檔案名稱處理
                string tmpName = $"Water Absorbency Test_{model.Main.OrderID}_{model.Main.StyleID}_" +
                                 $"{model.Main.FabricRefNo}_{model.Main.FabricColor}_{model.Main.Result}_{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[/:?""<>|*%]", string.Empty);
                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                string outputDirectory = Path.Combine(baseFilePath, "TMP");
                string filePath = Path.Combine(outputDirectory, $"{tmpName}.xlsx");
                string pdfPath = Path.Combine(outputDirectory, $"{tmpName}.pdf");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填寫主要報表內容
                    worksheet.Cell("D4").Value = model.Main.ReportNo;
                    worksheet.Cell("G4").Value = model.Main.OrderID;

                    worksheet.Cell("D5").Value = model.Main.FactoryID;
                    worksheet.Cell("G5").Value = model.Main.SubmitDateText;

                    worksheet.Cell("D6").Value = model.Main.StyleID;
                    worksheet.Cell("G6").Value = model.Main.ReportDateText;

                    worksheet.Cell("D7").Value = model.Main.Article;
                    worksheet.Cell("D8").Value = model.Main.SeasonID;
                    worksheet.Cell("D9").Value = model.Main.FabricRefNo;
                    worksheet.Cell("D10").Value = model.Main.FabricColor;
                    worksheet.Cell("E11").Value = model.Main.FabricDescription;

                    // BrandID
                    if (model.Main.BrandID.ToUpper() == "ADIDAS")
                    {
                        worksheet.Cell("C3").Value = "☑ ADIDAS";
                    }
                    if (model.Main.BrandID.ToUpper() == "REEBOK")
                    {
                        worksheet.Cell("F3").Value = "☑ REEBOK";
                    }

                    // Technician 資料處理
                    if (reportTechnician.Rows.Count > 0)
                    {
                        worksheet.Cell("F69").Value = reportTechnician.Rows[0]["Technician"]?.ToString();
                        AddImageToWorksheet(worksheet, reportTechnician.Rows[0]["TechnicianSignture"] as byte[], 69, 7, 100, 24);
                    }

                    // 插入圖片
                    AddImageToWorksheet(worksheet, model.Main.TestBeforePicture, 21, 2, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestBeforePicture, 42, 2, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestBeforeWashPicture, 21, 7, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestAfterPicture, 42, 7, 200, 300);

                    // 表身處理
                    if (model.DetailList.Any())
                    {
                        var detailBefore = model.DetailList.FirstOrDefault(x => x.EvaluationItem == "Before Wash");
                        if (detailBefore != null)
                        {
                            worksheet.Cell("D15").Value = detailBefore.NoOfDrops;
                            if (model.Main.FabricType == "KNIT")
                                worksheet.Cell("E15").Value = detailBefore.Values;
                            else if (model.Main.FabricType == "WOVEN")
                                worksheet.Cell("G15").Value = detailBefore.Values;
                            worksheet.Cell("H15").Value = detailBefore.Result;
                            worksheet.Cell("I15").Value = model.Main.Result;
                        }

                        var detailAfter = model.DetailList.FirstOrDefault(x => x.EvaluationItem == "After 5 Cycle Wash");
                        if (detailAfter != null)
                        {
                            worksheet.Cell("D17").Value = detailAfter.NoOfDrops;
                            if (model.Main.FabricType == "KNIT")
                                worksheet.Cell("E17").Value = detailAfter.Values;
                            else if (model.Main.FabricType == "WOVEN")
                                worksheet.Cell("G17").Value = detailAfter.Values;
                            worksheet.Cell("H17").Value = detailAfter.Result;
                        }
                    }

                    // 儲存 Excel 檔案
                    workbook.SaveAs(filePath);
                }

                // PDF 轉換
                if (isPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(filePath, pdfPath))
                    {
                        result.TempFileName = $"{tmpName}.pdf";
                    }
                    else
                    {
                        result.Result = false;
                        result.ErrorMessage = "Convert To PDF Fail";
                        return result;
                    }
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

        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);

            WaterAbsorbency_ViewModel model = this.GetData(new WaterAbsorbency_Request() { ReportNo = ReportNo });
            string name = $"Water Absorbency Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            WaterAbsorbency_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Water Absorbency Test/{model.Main.OrderID}/" +
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
