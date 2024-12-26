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
using System.Web;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class MartindalePillingTestService
    {
        private MartindalePillingTestProvider _Provider;
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        public MartindalePillingTest_ViewModel GetDefaultModel(bool isNew = false)
        {
            MartindalePillingTest_ViewModel model = new MartindalePillingTest_ViewModel()
            {
                Request = new MartindalePillingTest_Request(),
                Main = new MartindalePillingTest_Main()
                {
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                TestStandard_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<MartindalePillingTest_Detail>(),
            };

            try
            {
                _Provider = new MartindalePillingTestProvider(Common.ProductionDataAccessLayer);
                model.Scale_Source = _Provider.GetScales();

                if (isNew)
                {
                    // 預設每個rubTimes 各三筆
                    int idx = 1;
                    foreach (var rubTimes in model.DefaultRubTimes)
                    {
                        for (int i = 0; i <= 2; i++)
                        {
                            MartindalePillingTest_Detail nData = new MartindalePillingTest_Detail()
                            {
                                RubTimes = rubTimes,
                                EvaluationItem = $"Specimen {idx}",
                                Scale = "3-4",
                                Result = "Pass",
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
        public MartindalePillingTest_ViewModel GetData(MartindalePillingTest_Request Req)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new MartindalePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<MartindalePillingTest_Main> tmpList = new List<MartindalePillingTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new MartindalePillingTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new MartindalePillingTest_Request()
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
                    _Provider = new MartindalePillingTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new MartindalePillingTest_Request() { OrderID = model.Main.OrderID });

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

                    _Provider = new MartindalePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new MartindalePillingTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Martindale Pilling Test Evaporation Rate Test/{model.Main.OrderID}/" +
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
        public MartindalePillingTest_ViewModel GetOrderInfo(string OrderID)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new MartindalePillingTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new MartindalePillingTest_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 取得表頭SP#相關欄位
                    model.Main.OrderID = tmpOrders.FirstOrDefault().ID;
                    model.Main.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                    model.Main.BrandID = tmpOrders.FirstOrDefault().BrandID;
                    model.Main.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                    model.Main.StyleID = tmpOrders.FirstOrDefault().StyleID;
                    model.Main.FabricType = tmpOrders.FirstOrDefault().FabricType;

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
        public MartindalePillingTest_ViewModel NewSave(MartindalePillingTest_ViewModel Req, string MDivision, string UserID)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new MartindalePillingTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<MartindalePillingTest_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_MartindalePillingTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // MartindalePillingTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<MartindalePillingTest_Detail>();
                }

                _Provider.Processe_MartindalePillingTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new MartindalePillingTest_Request()
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
        public MartindalePillingTest_ViewModel EditSave(MartindalePillingTest_ViewModel Req, string UserID)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new MartindalePillingTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<MartindalePillingTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_MartindalePillingTest(Req, UserID);

                // 更新表身
                _Provider.Processe_MartindalePillingTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new MartindalePillingTest_Request()
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
        public MartindalePillingTest_ViewModel Delete(MartindalePillingTest_ViewModel Req)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new MartindalePillingTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_MartindalePillingTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new MartindalePillingTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new MartindalePillingTest_Request()
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
        public MartindalePillingTest_ViewModel EncodeAmend(MartindalePillingTest_ViewModel Req, string UserID)
        {
            MartindalePillingTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new MartindalePillingTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_MartindalePillingTest(Req.Main, UserID);

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
        public MartindalePillingTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            MartindalePillingTest_ViewModel result = new MartindalePillingTest_ViewModel();

            string basefileName = "MartindalePillingTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            if (!System.IO.File.Exists(openfilepath))
            {
                throw new FileNotFoundException("Excel template not found", openfilepath);
            }

            string tmpName = string.Empty;

            try
            {
                _Provider = new MartindalePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

                MartindalePillingTest_ViewModel model = this.GetData(new MartindalePillingTest_Request() { ReportNo = ReportNo });
                var testCode = _QualityBrandTestCodeProvider.Get(model.Main.BrandID, "T-Peel Strength Test");
                DataTable ReportTechnician = _Provider.GetReportTechnician(new MartindalePillingTest_Request() { ReportNo = ReportNo });

                tmpName = $"Martindale Pilling Test Evaporation Rate Test_{model.Main.OrderID}_" +
                          $"{model.Main.StyleID}_" +
                          $"{model.Main.FabricRefNo}_" +
                          $"{model.Main.FabricColor}_" +
                          $"{model.Main.Result}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";

                using (var workbook = new XLWorkbook(openfilepath))
                {
                    var worksheet = workbook.Worksheet(1);

                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 1).Value = $"T-Peel Strength Test ({testCode.FirstOrDefault().TestCode})";
                    }

                    worksheet.Cell(3, 2).Value = model.Main.ReportNo;
                    worksheet.Cell(3, 6).Value = model.Main.OrderID;
                    worksheet.Cell(4, 2).Value = model.Main.FactoryID;
                    worksheet.Cell(4, 6).Value = model.Main.SubmitDateText;
                    worksheet.Cell(5, 2).Value = model.Main.StyleID;
                    worksheet.Cell(5, 6).Value = model.Main.ReportDateText;
                    worksheet.Cell(6, 2).Value = model.Main.Article;
                    worksheet.Cell(6, 6).Value = model.Main.FabricRefNo;
                    worksheet.Cell(7, 2).Value = model.Main.SeasonID;
                    worksheet.Cell(7, 6).Value = model.Main.FabricColor;

                    if (model.Main.FabricType == "WOVEN")
                    {
                        worksheet.Cell("A8").Value = "FABRIC TYPE:               ☑";
                    }
                    if (model.Main.FabricType == "KNIT")
                    {
                        worksheet.Cell("D8").Value = "☑";
                    }

                    if (model.Main.Result == "Pass")
                    {
                        worksheet.Cell("I10").Value = "☑";
                    }
                    else
                    {
                        worksheet.Cell("I13").Value = "☑";
                    }

                    AddImageToWorksheet(worksheet, model.Main.TestBeforePicture, 18, 1, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.TestBeforePicture, 31, 1, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.Test500AfterPicture, 18, 6, 200, 300);
                    AddImageToWorksheet(worksheet, model.Main.Test2000AfterPicture, 31, 6, 200, 300);

                    if (ReportTechnician.Rows.Count > 0)
                    {
                        string technicianName = ReportTechnician.Rows[0]["Technician"].ToString();
                        worksheet.Cell(45, 5).Value = technicianName;

                        if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                        {
                            byte[] signature = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"];
                            AddImageToWorksheet(worksheet, signature, 45, 6, 100, 24);
                        }
                    }

                    if (model.DetailList.Any())
                    {
                        int rowIdx = 0;
                        foreach (var rubTimes in model.DefaultRubTimes)
                        {
                            worksheet.Cell(10 + rowIdx, 1).Value = rubTimes;

                            MartindalePillingTest_Detail WorstDetail = model.DetailList.Where(o => o.RubTimes == rubTimes).OrderBy(o => o.Scale).FirstOrDefault();
                            worksheet.Cell(10 + rowIdx, 6).Value = $"{model.Main.TestStandard} {WorstDetail?.Scale}";

                            rowIdx += 3;
                        }

                        rowIdx = 0;
                        foreach (var detailData in model.DetailList)
                        {
                            worksheet.Cell(10 + rowIdx, 2).Value = detailData.EvaluationItem;
                            worksheet.Cell(10 + rowIdx, 4).Value = $"{model.Main.TestStandard} {detailData.Scale}";
                            rowIdx++;
                        }
                    }

                    tmpName = RemoveInvalidFileNameChars(tmpName);

                    string fileName = $"{tmpName}.xlsx";
                    string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);
                    string filePdfName = $"{tmpName}.pdf";
                    string fullPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName);

                    workbook.SaveAs(fullExcelFileName);
                    if (isPDF)
                    {
                        ConvertToPDF.ExcelToPDF(fullExcelFileName, fullPdfFileName);
                        result.TempFileName = filePdfName;
                    }
                    else
                    {
                        result.TempFileName = fileName;
                    }
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

        private string RemoveInvalidFileNameChars(string input)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                input = input.Replace(c.ToString(), "");
            }
            return input;
        }



        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new MartindalePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            MartindalePillingTest_ViewModel model = this.GetData(new MartindalePillingTest_Request() { ReportNo = ReportNo });

            string name = $"Martindale Pilling Test Evaporation Rate Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
            MartindalePillingTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Martindale Pilling Test Evaporation Rate Test/{model.Main.OrderID}/" +
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
