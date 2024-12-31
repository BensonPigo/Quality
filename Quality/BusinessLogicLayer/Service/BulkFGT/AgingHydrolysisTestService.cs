using ADOHelper.Utility;
using BusinessLogicLayer.Helper;
using ClosedXML.Excel;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Ict;
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

//using static Sci.MyUtility;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AgingHydrolysisTestService
    {
        public AgingHydrolysisTest_Provider _Provider;
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;

        public AgingHydrolysisTest_ViewModel GetDefaultModel()
        {
            return new AgingHydrolysisTest_ViewModel()
            {
                Request = new AgingHydrolysisTest_Request(),
                MainData = new AgingHydrolysisTest_Main()
                {
                },
                OrderID_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                TimeUnit_Source = new List<System.Web.Mvc.SelectListItem>()
                {
                    new SelectListItem(){Text = "",Value=""},
                    new SelectListItem(){Text = "Hour",Value="Hour"},
                    new SelectListItem(){Text = "Day",Value="Day"}
                },
                MainList = new List<AgingHydrolysisTest_Main>(),
                DetailList = new List<AgingHydrolysisTest_Detail>(),
            };
        }

        public AgingHydrolysisTest_Detail_ViewModel GetDefaulDeailtModel()
        {
            _Provider = new AgingHydrolysisTest_Provider(Common.ProductionDataAccessLayer);
            return new AgingHydrolysisTest_Detail_ViewModel()
            {
                // 預設產生五筆
                MockupList = new List<AgingHydrolysisTest_Detail_Mockup>()
                {
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen1" ,ChangeScaleStandard="4-5",StainingScaleStandard="4" ,ChangeScale="4-5",StainingScale="4",ChangeResult="Pass",StainingResult="Pass"},
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen2" ,ChangeScaleStandard="4-5",StainingScaleStandard="4" ,ChangeScale="4-5",StainingScale="4",ChangeResult="Pass",StainingResult="Pass"},
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen3" ,ChangeScaleStandard="4-5",StainingScaleStandard="4" ,ChangeScale="4-5",StainingScale="4",ChangeResult="Pass",StainingResult="Pass"},
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen4" ,ChangeScaleStandard="4-5",StainingScaleStandard="4" ,ChangeScale="4-5",StainingScale="4",ChangeResult="Pass",StainingResult="Pass"},
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen5" ,ChangeScaleStandard="4-5",StainingScaleStandard="4" ,ChangeScale="4-5",StainingScale="4",ChangeResult="Pass",StainingResult="Pass"},
                },
                Scale_Source = _Provider.GetScales()
            };
        }

        /// <summary>
        /// 取得Order ID的資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public AgingHydrolysisTest_ViewModel GetOrderInfo(string OrderID)
        {
            AgingHydrolysisTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new AgingHydrolysisTest_Provider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new AgingHydrolysisTest_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 檢查SP# 是否已經存在AgingHydrolysisTest
                    _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);
                    var existsData = _Provider.GetMainList(new AgingHydrolysisTest_Request()
                    {
                        OrderID = tmpOrders.FirstOrDefault().ID,
                        BrandID = tmpOrders.FirstOrDefault().BrandID,
                        SeasonID = tmpOrders.FirstOrDefault().SeasonID,
                        StyleID = tmpOrders.FirstOrDefault().StyleID,
                    });

                    if (existsData.Any())
                    {
                        model.Result = false;
                        model.ErrorMessage = "Data has existed.";
                    }
                    else
                    {

                        model.MainData.OrderID = tmpOrders.FirstOrDefault().ID;
                        model.MainData.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                        model.MainData.BrandID = tmpOrders.FirstOrDefault().BrandID;
                        model.MainData.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                        model.MainData.StyleID = tmpOrders.FirstOrDefault().StyleID;

                        // 取得Article 下拉選單
                        foreach (var oriData in tmpOrders)
                        {
                            SelectListItem Article = new SelectListItem()
                            {
                                Text = oriData.Article,
                                Value = oriData.Article,
                            };
                            model.Article_Source.Add(Article);
                        }

                        // 預設選取第一筆
                        model.MainData.Article = model.Article_Source.FirstOrDefault().Value;
                        model.Result = true;
                    }
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

        /// <summary>
        /// 取得主要畫面
        /// </summary>
        /// <param name="Req">必需：Brand、Season、Style、Article、選填：OrderID</param>
        /// <returns></returns>
        public AgingHydrolysisTest_ViewModel GetMainPage(AgingHydrolysisTest_Request Req)
        {
            AgingHydrolysisTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && string.IsNullOrEmpty(Req.OrderID))
                {
                    model.Result = false;
                    model.ErrorMessage = "Please inout filter coluimn.";
                    return model;
                }

                // 取得符合條件的主表
                List<AgingHydrolysisTest_Main> tmpList = _Provider.GetMainList(new AgingHydrolysisTest_Request()
                {
                    BrandID = Req.BrandID,
                    SeasonID = Req.SeasonID,
                    StyleID = Req.StyleID,
                    Article = Req.Article,
                    //OrderID = Req.OrderID,
                    AgingHydrolysisTestID = Req.AgingHydrolysisTestID,
                });

                if (tmpList.Any())
                {
                    // 取得下拉選單
                    tmpList = _Provider.GetMainList(new AgingHydrolysisTest_Request()
                    {
                        BrandID = tmpList.FirstOrDefault().BrandID,
                        SeasonID = tmpList.FirstOrDefault().SeasonID,
                        StyleID = tmpList.FirstOrDefault().StyleID,
                        Article = tmpList.FirstOrDefault().Article,
                    });

                    model.OrderID_Source = new SetListItem().ItemListBinding(tmpList.Select(o => o.OrderID).Distinct().ToList());

                    // 若"有"傳入OrderID，則可以直接找出表頭表身明細
                    if (!string.IsNullOrEmpty(Req.OrderID) && tmpList.Any(o => o.OrderID == Req.OrderID))
                    {
                        // 取得表頭資料
                        model.MainData = tmpList.Where(o => o.OrderID == Req.OrderID).FirstOrDefault();
                    }
                    // 若"沒"傳入OrderID，抓下拉選單第一筆顯示
                    else
                    {
                        // 根據下拉選單第一筆，取得表頭資料
                        model.MainData = tmpList.Where(o => o.OrderID == model.OrderID_Source.FirstOrDefault().Value).FirstOrDefault();

                    }

                    model.Request.BrandID = model.MainData.BrandID;
                    model.Request.SeasonID = model.MainData.SeasonID;
                    model.Request.StyleID = model.MainData.StyleID;
                    model.Request.Article = model.MainData.Article;
                    model.Request.OrderID = model.MainData.OrderID;
                    model.Request.AgingHydrolysisTestID = model.MainData.ID;

                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new AgingHydrolysisTest_Request()
                    {
                        AgingHydrolysisTestID = model.MainData.ID
                    });

                    _Provider = new AgingHydrolysisTest_Provider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new AgingHydrolysisTest_Request() { OrderID = model.MainData.OrderID });

                    // 取得Article 下拉選單
                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
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

        /// <summary>
        /// AgingHydrolysisTest 新增/修改
        /// </summary>
        /// <param name="Req"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public AgingHydrolysisTest_ViewModel SaveMainPage(AgingHydrolysisTest_ViewModel Req, string UserID)
        {
            AgingHydrolysisTest_ViewModel model = this.GetDefaultModel();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new AgingHydrolysisTest_Provider(_ISQLDataTransaction);

                // 新增or修改，並取得AgingHydrolysisTest主檔
                AgingHydrolysisTest_Main main = _Provider.InsertUpdate_AgingHydrolysisTest(Req, UserID);


                // AgingHydrolysisTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<AgingHydrolysisTest_Detail>();
                }
                Req.MainData = main;
                _Provider.Processe_AgingHydrolysisTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();
                // 重新查詢資料
                model = this.GetMainPage(new AgingHydrolysisTest_Request()
                {
                    BrandID = main.BrandID,
                    SeasonID = main.SeasonID,
                    StyleID = main.StyleID,
                    Article = main.Article,
                    OrderID = main.OrderID,
                    AgingHydrolysisTestID = main.ID,
                });
                model.Request.BrandID = Req.MainData.BrandID;
                model.Request.SeasonID = Req.MainData.SeasonID;
                model.Request.StyleID = Req.MainData.StyleID;
                model.Request.Article = Req.MainData.Article;
                model.Request.OrderID = Req.MainData.OrderID;
                model.Request.AgingHydrolysisTestID = Req.MainData.ID;

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
        /// <summary>
        /// AgingHydrolysisTest 刪除，一併刪除底下的AgingHydrolysisTest_Detail、AgingHydrolysisTest_Detail_Mockup、AgingHydrolysisTest_Image
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public AgingHydrolysisTest_ViewModel DeleteMainPage(AgingHydrolysisTest_ViewModel Req)
        {
            AgingHydrolysisTest_ViewModel model = this.GetDefaultModel();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new AgingHydrolysisTest_Provider(_ISQLDataTransaction);

                if (Req.MainData.ID == 0)
                {
                    model.Result = false;
                    model.ErrorMessage = "There is no data ID, please notice MIS Team.";
                    return model;
                }

                _Provider.Delete_AgingHydrolysisTest(Req);

                _ISQLDataTransaction.Commit();
                model.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }

        /// <summary>
        ///  取得AgingHydrolysisTest_Detail頁面
        /// </summary>
        /// <param name="Req">必需：AgingHydrolysisTestID、選填：ReportNo</param>
        /// <returns></returns>
        public AgingHydrolysisTest_Detail_ViewModel GetDetailPage(AgingHydrolysisTest_Request Req)
        {
            AgingHydrolysisTest_Detail_ViewModel model = this.GetDefaulDeailtModel();

            try
            {
                _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);


                List<AgingHydrolysisTest_Detail> tmpList = _Provider.GetDetailList(new AgingHydrolysisTest_Request()
                {
                    AgingHydrolysisTestID = Req.AgingHydrolysisTestID,
                });

                // 取得相同AgingHydrolysisTestID的Detail資料
                model.MainDetailDataList = tmpList;

                if (tmpList.Any())
                {
                    // ReportNo 不為空，則帶出該筆的明細
                    if (!string.IsNullOrEmpty(Req.ReportNo))
                    {
                        model.MainDetailData = tmpList.Where(o => o.ReportNo == Req.ReportNo).FirstOrDefault();

                        // 取得現有Mockup List
                        var nowMockupList = _Provider.GetDetailMockupList(new AgingHydrolysisTest_Request()
                        {
                            ReportNo = Req.ReportNo,
                        });

                        // 有現存的資料
                        if (nowMockupList.Any())
                        {
                            model.MockupList = nowMockupList;
                        }
                        else
                        {
                            // 其餘則使用宣告時預設的五筆
                        }

                        string Subject = $"Accelerated Aging by Hydrolysis Test_{model.MainDetailData.OrderID}_" +
                            $"{model.MainDetailData.StyleID}_" +
                            $"{model.MainDetailData.FabricRefNo}_" +
                            $"{model.MainDetailData.FabricColor}_" +
                            $"{model.MainDetailData.Result}_" +
                            $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                        model.MainDetailData.MailSubject = Subject;
                    }

                }
                else
                {

                    model.Result = false;
                    model.ErrorMessage = "Data not found.";
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
        /// AgingHydrolysisTest_Detail頁面Save
        /// </summary>
        /// <param name="Req">必填：AgingHydrolysisTest_Detail，選填：MockupList</param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public AgingHydrolysisTest_Detail_ViewModel SaveDetailPage(AgingHydrolysisTest_Detail_ViewModel Req, string UserID)
        {
            AgingHydrolysisTest_Detail_ViewModel model = this.GetDefaulDeailtModel();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new AgingHydrolysisTest_Provider(_ISQLDataTransaction);

                // 修改AgingHydrolysisTest_Detail，這個地方不可能會有新增，因此不需要放AgingHydrolysisTest_Main.MDivisionID
                _Provider.Processe_AgingHydrolysisTest_Detail(new AgingHydrolysisTest_ViewModel()
                {
                    MainData = new AgingHydrolysisTest_Main() { ID = Req.MainDetailData.AgingHydrolysisTestID },
                    DetailList = new List<AgingHydrolysisTest_Detail>() { Req.MainDetailData },
                }
                , UserID: UserID, isSaveDetailPage: true);

                // AgingHydrolysisTest_Detail_Mockup 異動
                _Provider.Process_AgingHydrolysisTest_Detail_Mockup(Req, UserID);

                // 重新查詢資料
                model = this.GetDetailPage(new AgingHydrolysisTest_Request()
                {
                    AgingHydrolysisTestID = Req.MainDetailData.AgingHydrolysisTestID,
                    ReportNo = Req.MainDetailData.ReportNo,
                });

                _ISQLDataTransaction.Commit();
                model.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }

        /// <summary>
        /// AgingHydrolysisTest_Detail頁面Encode/Amend
        /// </summary>
        /// <param name="Req"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public AgingHydrolysisTest_ViewModel EncodeAmend(AgingHydrolysisTest_Detail Req, string UserID)
        {
            AgingHydrolysisTest_ViewModel model = this.GetDefaultModel();
            try
            {
                _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);
                _Provider.EncodeAmend_AgingHydrolysisTest_Detail(Req, UserID);

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
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
        public AgingHydrolysisTest_Detail_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            AgingHydrolysisTest_Detail_ViewModel result = new AgingHydrolysisTest_Detail_ViewModel();
            _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                string baseFileName = "AgingHydrolysisTest";
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string templatePath = Path.Combine(baseFilePath, "XLT", $"{baseFileName}.xltx");
                string tmpName = string.Empty;

                // 取得報表資料
                DataSet reportDataSet = _Provider.GetReport(new AgingHydrolysisTest_Request { ReportNo = ReportNo });
                DataTable reportTechnician = _Provider.GetReportTechnician(new AgingHydrolysisTest_Request { ReportNo = ReportNo });
                DataTable detailTable = reportDataSet.Tables[0];
                DataTable mockupTable = reportDataSet.Tables.Count > 1 ? reportDataSet.Tables[1] : null;

                if (detailTable.Rows.Count == 0)
                {
                    result.ErrorMessage = "Report No. not found.";
                    result.Result = false;
                    return result;
                }

                // 生成檔案名稱
                tmpName = $"Accelerated Aging by Hydrolysis Test_{detailTable.Rows[0]["OrderID"]}_{detailTable.Rows[0]["StyleID"]}_{detailTable.Rows[0]["FabricRefNo"]}_{detailTable.Rows[0]["FabricColor"]}_{detailTable.Rows[0]["Result"]}_{DateTime.Now:yyyyMMddHHmmss}";
                tmpName = Regex.Replace(tmpName, @"[/:?""<>|*%\t]", string.Empty);
                if (!string.IsNullOrWhiteSpace(AssignedFineName)) tmpName = AssignedFineName;

                string outputPath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    if (detailTable.Rows[0]["BrandID"]?.ToString() == "ADIDAS")
                    {
                        worksheet.Cell("A2").Value = "☑";
                    }
                    else if (detailTable.Rows[0]["BrandID"]?.ToString() == "REEBOK")
                    {
                        worksheet.Cell("C2").Value = "☑";
                    }

                    if (detailTable.Rows[0]["Result"].ToString().ToUpper() == "PASS")
                    {
                        worksheet.Cell("I11").Value = "☑";
                    }
                    else
                    {
                        worksheet.Cell("I13").Value = "☑";
                    }

                    // 填寫主要報表內容
                    worksheet.Cell("B3").Value = detailTable.Rows[0]["ReportNo"]?.ToString();
                    worksheet.Cell("F3").Value = detailTable.Rows[0]["OrderID"]?.ToString();
                    worksheet.Cell("B4").Value = detailTable.Rows[0]["FactoryID"]?.ToString();
                    worksheet.Cell("F4").Value = detailTable.Rows[0]["ReceivedDate"] == DBNull.Value ? string.Empty : Convert.ToDateTime(detailTable.Rows[0]["ReceivedDate"]).ToString("yyyy/MM/dd");
                    worksheet.Cell("B5").Value = detailTable.Rows[0]["StyleID"]?.ToString();
                    worksheet.Cell("F5").Value = detailTable.Rows[0]["ReportDate"] == DBNull.Value ? string.Empty : Convert.ToDateTime(detailTable.Rows[0]["ReportDate"]).ToString("yyyy/MM/dd");

                    worksheet.Cell("B6").Value = detailTable.Rows[0]["Article"]?.ToString();
                    worksheet.Cell("F6").Value = detailTable.Rows[0]["FabricRefNo"]?.ToString();
                    worksheet.Cell("B7").Value = detailTable.Rows[0]["SeasonID"]?.ToString();
                    worksheet.Cell("F7").Value = detailTable.Rows[0]["FabricColor"]?.ToString();
                    worksheet.Cell("B8").Value = detailTable.Rows[0]["MaterialType"]?.ToString();
                    worksheet.Cell("F8").Value = detailTable.Rows[0]["Comment"]?.ToString();

                    // 插入圖片
                    AddImageToWorksheet(worksheet, detailTable.Rows[0]["TestBeforePicture"] as byte[], 19, 1, 360, 540);
                    AddImageToWorksheet(worksheet, detailTable.Rows[0]["TestAfterPicture"] as byte[], 19, 5, 360, 540);

                    // 技術人員與簽名
                    if (reportTechnician.Rows.Count > 0)
                    {
                        worksheet.Cell("A29").Value = reportTechnician.Rows[0]["Technician"].ToString();
                        AddImageToWorksheet(worksheet, reportTechnician.Rows[0]["TechnicianSignture"] as byte[], 31, 1, 100, 24);

                        worksheet.Cell("E29").Value = reportTechnician.Rows[0]["ApproverName"].ToString();
                        AddImageToWorksheet(worksheet, reportTechnician.Rows[0]["ApproverSignture"] as byte[], 31, 5, 100, 24);
                    }

                    // 填寫 Mockup 資料
                    if (mockupTable != null)
                    {
                        for (int i = 0; i < mockupTable.Rows.Count && i < 5; i++)
                        {
                            worksheet.Cell(11 + i, 1).Value = mockupTable.Rows[i]["SpecimenName"]?.ToString();
                            worksheet.Cell(11 + i, 2).Value = mockupTable.Rows[i]["ChangeScale"]?.ToString();
                            worksheet.Cell(11 + i, 4).Value = mockupTable.Rows[i]["StainingScale"]?.ToString();
                            worksheet.Cell(11 + i, 5).Value = mockupTable.Rows[i]["Comment"]?.ToString();
                        }
                    }

                    // 儲存 Excel 檔案
                    workbook.SaveAs(outputPath);
                }

                result.TempFileName = $"{tmpName}.xlsx";
                result.Result = true;

                // PDF 轉換
                if (isPDF)
                {
                    LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    officeService.ConvertExcelToPdf(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));

                    result.TempFileName = $"{tmpName}.pdf";
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }

            return result;
        }
        public SendMail_Result FailSendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new AgingHydrolysisTest_Provider(Common.ManufacturingExecutionDataAccessLayer);

            var detail = this.GetDetailPage(new AgingHydrolysisTest_Request() { ReportNo = ReportNo });
            AgingHydrolysisTest_Main MainData = _Provider.GetMainList(new AgingHydrolysisTest_Request()
            {
                AgingHydrolysisTestID = detail.MainDetailData.AgingHydrolysisTestID,
            }).FirstOrDefault();

            string name = $"Accelerated Aging by Hydrolysis Test_{MainData.OrderID}_" +
                $"{MainData.StyleID}_" +
                $"{detail.MainDetailData.FabricRefNo}_" +
                $"{detail.MainDetailData.FabricColor}_" +
                $"{detail.MainDetailData.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            AgingHydrolysisTest_Detail_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName) ;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Accelerated Aging by Hydrolysis Test/{MainData.OrderID}/" +
                    $"{MainData.StyleID}/" +
                    $"{detail.MainDetailData.FabricRefNo}/" +
                    $"{detail.MainDetailData.FabricColor}/" +
                    $"{detail.MainDetailData.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                To = TO,
                CC = CC,
                Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                FileUploader = Files,
                IsShowAIComment = true,
                AICommentType = "Accelerated Aging by Hydrolysis",
                StyleID = MainData.StyleID,
                SeasonID = MainData.SeasonID,
                BrandID = MainData.BrandID,
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
