using ADOHelper.Utility;
using BusinessLogicLayer.Helper;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Tsp;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using static Sci.MyUtility;
using static System.Net.Mime.MediaTypeNames;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class RandomTumblePillingTestService
    {
        private RandomTumblePillingTestProvider _Provider;
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
                model.Scale_Source = _Provider.GetScales();

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

        public RandomTumblePillingTest_ViewModel GetReport(string ReportNo, bool isPDF)
        {
            RandomTumblePillingTest_ViewModel result = new RandomTumblePillingTest_ViewModel();

            string basefileName = "RandomTumblePillingTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {
                _Provider = new RandomTumblePillingTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料

                RandomTumblePillingTest_ViewModel model = this.GetData(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });

                DataTable ReportTechnician = _Provider.GetReportTechnician(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                // 取得工作表上所有圖形物件
                Microsoft.Office.Interop.Excel.Shapes shapes = worksheet.Shapes;

                // 根據名稱，搜尋文字方塊物件
                Microsoft.Office.Interop.Excel.Shape Woven_TextBox = shapes.Item("Woven_TextBox");
                Microsoft.Office.Interop.Excel.Shape Knit_TextBox = shapes.Item("Knit_TextBox");

                // 固定欄位

                // FabricType
                //if (model.Main.FabricType.ToUpper() == "WOVEN")
                //{
                //    Woven_TextBox.TextFrame.Characters().Text = "V";
                //}
                //if (model.Main.FabricType.ToUpper() == "KNIT")
                //{
                //    Knit_TextBox.TextFrame.Characters().Text = "V";
                //}

                string reportNo = model.Main.ReportNo;
                worksheet.Cells[3, 2] = model.Main.ReportNo;
                worksheet.Cells[3, 6] = model.Main.OrderID;

                worksheet.Cells[4, 2] = model.Main.FactoryID;
                worksheet.Cells[4, 6] = model.Main.SubmitDateText;

                worksheet.Cells[5, 2] = model.Main.StyleID;
                worksheet.Cells[5, 6] = model.Main.ReportDateText;

                worksheet.Cells[6, 2] = model.Main.Article;
                worksheet.Cells[6, 2] = model.Main.SeasonID;

                worksheet.Cells[7, 2] = model.Main.SeasonID;
                worksheet.Cells[7, 6] = model.Main.FabricColor;

                if (model.Main.Result == "Pass")
                {
                    worksheet.Cells[11, 9] = "V";
                }
                else
                {
                    worksheet.Cells[14, 9] = "V";
                }

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[45, 5] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[45, 6];
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

                // TestBeforePicture、TestAfterPicture 圖片
                if (model.Main.TestFaceSideBeforePicture != null && model.Main.TestFaceSideBeforePicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[18, 1];
                    Microsoft.Office.Interop.Excel.Range cell2 = worksheet.Cells[31, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestFaceSideBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell2.Left + 5, cell2.Top + 5, 200, 300);
                }
                if (model.Main.TestFaceSideAfterPicture != null && model.Main.TestFaceSideAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[18, 6];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestFaceSideAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }
                if (model.Main.TestBackSideBeforePicture != null && model.Main.TestBackSideBeforePicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[31, 6];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBackSideBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }
                if (model.Main.TestBackSideAfterPicture != null && model.Main.TestBackSideAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[31, 6];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBackSideAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // 表身處理
                if (model.DetailList.Any() && model.DetailList.Count > 1)
                {
                    int rowIdx = 0;
                    //foreach (var side in model.DefaultSide)
                    //{
                    //    // Rub Times
                    //    worksheet.Cells[10 + rowIdx, 1] = side;

                    //    RandomTumblePillingTest_Detail WorstDetail = model.DetailList.Where(o => o.Side == side).OrderBy(o => o.Scale).FirstOrDefault();

                    //    //
                    //    worksheet.Cells[10 + rowIdx, 6] = $@"{model.Main.TestStandard} {WorstDetail.Scale}";

                    //    rowIdx += 3;
                    //}
                    //rowIdx = 0;
                    //foreach (var detailData in model.DetailList)
                    //{
                    //    worksheet.Cells[10 + rowIdx, 2] = detailData.EvaluationItem;
                    //    worksheet.Cells[10 + rowIdx, 4] = $@"{model.Main.TestStandard} {detailData.Scale}";
                    //    rowIdx += 1;
                    //}

                }

                string fileName = $"RandomTumblePillingTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"RandomTumblePillingTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
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
    }
}
