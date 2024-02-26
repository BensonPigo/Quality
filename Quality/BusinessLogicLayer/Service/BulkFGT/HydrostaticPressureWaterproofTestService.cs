using ADOHelper.Utility;
using DatabaseObject.ViewModel.BulkFGT;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Web.Mvc;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Provider.MSSQL;
using System.Net.Mail;
using System.Web.UI.WebControls;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class HydrostaticPressureWaterproofTestService
    {
        private HydrostaticPressureWaterproofTestProvider _Provider;
        private MailToolsService _MailService;
        public HydrostaticPressureWaterproofTest_ViewModel GetDefaultModel(bool IsNew = false)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = new HydrostaticPressureWaterproofTest_ViewModel()
            {
                Request = new HydrostaticPressureWaterproofTest_Request(),
                Main = new HydrostaticPressureWaterproofTest_Main()
                {
                    // ISP20230288 寫死4
                    //MachineReport = "",
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<HydrostaticPressureWaterproofTest_Detail>(),

            };

            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                model.Standard_Source = _Provider.GetStandards();

                if (IsNew)
                {
                    model.Main.Temperature = 40;
                    model.Main.WashCycles = 10;
                    model.Main.DryingCondition = "Tumble dry low";

                    // 取得預設表身選項
                    var Standards = _Provider.GetStandards();

                    // 取得分組資料
                    var groupData = Standards.Select(o => new { o.EvaluationType, o.EvaluationItem }).Distinct();

                    // 產生表身，預設填入值為Pass
                    foreach (var data in groupData)
                    {
                        var asReceivedData = Standards.Where(o => o.EvaluationType == data.EvaluationType && o.EvaluationItem == data.EvaluationItem && o.Phase == "As received" && o.Result == "Pass").FirstOrDefault();
                        var afterWashData = Standards.Where(o => o.EvaluationType == data.EvaluationType && o.EvaluationItem == data.EvaluationItem && o.Phase == "After wash" && o.Result == "Pass").FirstOrDefault();

                        HydrostaticPressureWaterproofTest_Detail d = new HydrostaticPressureWaterproofTest_Detail()
                        {
                            EvaluationType = data.EvaluationType,
                            EvaluationItem = data.EvaluationItem,
                            AsReceivedValue = asReceivedData.Standard,
                            AsReceivedResult = asReceivedData.Result,
                            AfterWashValue = afterWashData.Standard,
                            AfterWashResult = afterWashData.Result,
                        };
                        model.DetailList.Add(d);
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
        public HydrostaticPressureWaterproofTest_ViewModel GetData(HydrostaticPressureWaterproofTest_Request Req)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<HydrostaticPressureWaterproofTest_Main> tmpList = new List<HydrostaticPressureWaterproofTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new HydrostaticPressureWaterproofTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new HydrostaticPressureWaterproofTest_Request()
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

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new HydrostaticPressureWaterproofTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    // 取得Article 下拉選單
                    _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new HydrostaticPressureWaterproofTest_Request() { OrderID = model.Main.OrderID });

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

        public HydrostaticPressureWaterproofTest_ViewModel GetOrderInfo(string OrderID)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new HydrostaticPressureWaterproofTest_Request() { OrderID = OrderID });


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
        public HydrostaticPressureWaterproofTest_ViewModel NewSave(HydrostaticPressureWaterproofTest_ViewModel Req, string MDivision, string UserID)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasAsReceivedFail = Req.DetailList.Where(o => o.AsReceivedResult == "Fail").Any();
                    bool HasAfterWashFail = Req.DetailList.Where(o => o.AfterWashResult == "Fail").Any();
                    Req.Main.Result = HasAsReceivedFail || HasAfterWashFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<HydrostaticPressureWaterproofTest_Detail>();
                }

                // 新增，並取得ReportNo
                _Provider.Insert_HydrostaticPressureWaterproofTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // HydrostaticPressureWaterproofTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<HydrostaticPressureWaterproofTest_Detail>();
                }

                _Provider.Processe_HydrostaticPressureWaterproofTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new HydrostaticPressureWaterproofTest_Request()
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
        public HydrostaticPressureWaterproofTest_ViewModel EditSave(HydrostaticPressureWaterproofTest_ViewModel Req, string UserID)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasAsReceivedFail = Req.DetailList.Where(o => o.AsReceivedResult == "Fail").Any();
                    bool HasAfterWashFail = Req.DetailList.Where(o => o.AfterWashResult == "Fail").Any();
                    Req.Main.Result = HasAsReceivedFail  || HasAfterWashFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<HydrostaticPressureWaterproofTest_Detail>();
                }

                //if (Req.Main.MachineReportFile != null)
                //{
                //    this.CreateMachineReportFile(Req);
                //}

                // 更新表頭，並取得ReportNo
                _Provider.Update_HydrostaticPressureWaterproofTest(Req, UserID);

                // 更新表身
                _Provider.Processe_HydrostaticPressureWaterproofTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new HydrostaticPressureWaterproofTest_Request()
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
        public HydrostaticPressureWaterproofTest_ViewModel Delete(HydrostaticPressureWaterproofTest_ViewModel Req)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);


            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.Delete_HydrostaticPressureWaterproofTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new HydrostaticPressureWaterproofTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new HydrostaticPressureWaterproofTest_Request()
                {
                    ReportNo = model.Main.ReportNo,
                    BrandID = model.Main.BrandID,
                    SeasonID = model.Main.SeasonID,
                    StyleID = model.Main.StyleID,
                    Article = model.Main.Article,
                };
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
        public HydrostaticPressureWaterproofTest_ViewModel EncodeAmend(HydrostaticPressureWaterproofTest_ViewModel Req, string UserID)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new HydrostaticPressureWaterproofTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_HydrostaticPressureWaterproofTest(Req.Main, UserID);

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
        public HydrostaticPressureWaterproofTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            HydrostaticPressureWaterproofTest_ViewModel result = new HydrostaticPressureWaterproofTest_ViewModel();

            string basefileName = "HydrostaticPressureWaterproofTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            string tmpName = string.Empty;

            try
            {

                // 取得報表資料

                HydrostaticPressureWaterproofTest_ViewModel model = this.GetData(new HydrostaticPressureWaterproofTest_Request() { ReportNo = ReportNo });

                _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                DataTable ReportTechnician = _Provider.GetReportTechnician(new HydrostaticPressureWaterproofTest_Request() { ReportNo = ReportNo });

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表
                tmpName = $"Hydrostatic Pressure Waterproof Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                string reportNo = model.Main.ReportNo;
                //string machineReport = string.IsNullOrEmpty(model.Main.MachineReport) ? string.Empty : model.Main.MachineReport;

                worksheet.Cells[3, 2] = model.Main.ReportNo;

                worksheet.Cells[4, 2] = model.Main.SubmitDateText;
                worksheet.Cells[4, 5] = model.Main.ReportDateText;

                worksheet.Cells[5, 2] = model.Main.OrderID;
                worksheet.Cells[5, 5] = model.Main.BrandID;

                worksheet.Cells[6, 2] = model.Main.StyleID;
                worksheet.Cells[6, 5] = model.Main.SeasonID;

                worksheet.Cells[7, 2] = model.Main.Article;
                worksheet.Cells[7, 5] = model.Main.FabricColor;

                worksheet.Cells[9, 2] = model.Main.Temperature;
                worksheet.Cells[9, 4] = model.Main.DryingCondition;
                worksheet.Cells[9, 6] = model.Main.WashCycles;

                worksheet.Cells[65, 2] = model.Main.Remark;

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[80, 5] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[79, 5];
                    if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    }
                }

                // TestBeforePicture 圖片
                if (model.Main.TestBeforePicture != null && model.Main.TestBeforePicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[67, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // TestAfterPicture 圖片
                if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[67, 4];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // 表身處理
                if (model.DetailList.Any() && model.DetailList.Count > 1)
                {
                    foreach (var evaluationType in model.EvaluationTypeList)
                    {
                        var sameEvaluationType = model.DetailList.Where(o => o.EvaluationType == evaluationType).OrderBy(o => o.EvaluationItemSeq);

                        if (evaluationType == "Seam procedure(mm)")
                        {
                            int idx = 0;
                            foreach (var item in sameEvaluationType)
                            {
                                // As received
                                worksheet.Cells[12 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[12 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "As received").Standard;
                                worksheet.Cells[12 + idx, 4] = item.AsReceivedValue;
                                worksheet.Cells[12 + idx, 6] = item.AsReceivedResult;

                                // Aftr wash
                                worksheet.Cells[23 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[23 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "After wash").Standard;
                                worksheet.Cells[23 + idx, 4] = item.AfterWashValue;
                                worksheet.Cells[23 + idx, 6] = item.AfterWashResult;
                                idx++;
                            }
                        }
                        else if (evaluationType == "Fabric procedure")
                        {
                            int idx = 0;
                            foreach (var item in sameEvaluationType)
                            {
                                // As received
                                worksheet.Cells[33 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[33 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "As received").Standard;
                                worksheet.Cells[33 + idx, 4] = item.AsReceivedValue;
                                worksheet.Cells[33 + idx, 6] = item.AsReceivedResult;

                                // Aftr wash
                                worksheet.Cells[36 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[36 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "After wash").Standard;
                                worksheet.Cells[36 + idx, 4] = item.AfterWashValue;
                                worksheet.Cells[36 + idx, 6] = item.AfterWashResult;
                                idx++;
                                // 因為只有一列資料，寫完就跳出
                                break;
                            }
                        }
                        else if (evaluationType == "Heat transfer/ logo procedure")
                        {
                            int idx = 0;
                            foreach (var item in sameEvaluationType)
                            {
                                // As received
                                worksheet.Cells[39 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[39 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "As received").Standard;
                                worksheet.Cells[39 + idx, 4] = item.AsReceivedValue;
                                worksheet.Cells[39 + idx, 6] = item.AsReceivedResult;

                                // Aftr wash
                                worksheet.Cells[43 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[43 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "After wash").Standard;
                                worksheet.Cells[43 + idx, 4] = item.AfterWashValue;
                                worksheet.Cells[43 + idx, 6] = item.AfterWashResult;
                                idx++;
                                // 因為只有一列資料，寫完就跳出
                                break;
                            }
                        }
                        else if (evaluationType == "Seam procedure(min)")
                        {
                            int idx = 0;
                            foreach (var item in sameEvaluationType)
                            {
                                // As received
                                worksheet.Cells[46 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[46 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "As received").Standard;
                                worksheet.Cells[46 + idx, 4] = item.AsReceivedValue;
                                worksheet.Cells[46 + idx, 6] = item.AsReceivedResult;

                                // Aftr wash
                                worksheet.Cells[57 + idx, 1] = item.EvaluationItem;
                                worksheet.Cells[57 + idx, 3] = model.GetStandard(evaluationType, item.EvaluationItem, "After wash").Standard;
                                worksheet.Cells[57 + idx, 4] = item.AfterWashValue;
                                worksheet.Cells[57 + idx, 6] = item.AfterWashResult;
                                idx++;
                            }
                        }
                    }


                }

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                string fileName = $"{tmpName}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"{tmpName}.pdf";
                string fullPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName);


                Microsoft.Office.Interop.Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(fullExcelFileName);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);

                // 轉PDF再繼續進行以下
                if (isPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(fullExcelFileName, fullPdfFileName))
                    {
                        result.TempFileName = filePdfName;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }
                }
                else
                {
                    result.TempFileName = fileName;
                    result.Result = true;
                }

            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }
            finally
            {
                Marshal.ReleaseComObject(excel);
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


        public SendMail_Result SendMail(string ReportNo, string TO, string CC)
        {
            _Provider = new HydrostaticPressureWaterproofTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            HydrostaticPressureWaterproofTest_ViewModel model = this.GetData(new HydrostaticPressureWaterproofTest_Request() { ReportNo = ReportNo });

            string name = $"Hydrostatic Pressure Waterproof Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            HydrostaticPressureWaterproofTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Hydrostatic Pressure Waterproof Test/{model.Main.OrderID}/" +
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
                IsShowAIComment = true,
            };

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
            sendMail_Request.Body = sendMail_Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
