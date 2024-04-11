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
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Web.Mvc;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Web;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class StickerTestService
    {
        private StickerTestProvider _Provider;
        private MailToolsService _MailService;
        public StickerTest_ViewModel GetDefaultModel(bool IsNew = false)
        {
            StickerTest_ViewModel model = new StickerTest_ViewModel()
            {
                Request = new StickerTest_Request(),
                Main = new StickerTest_Main()
                {
                    // ISP20230288 寫死4
                    //MachineReport = "",
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                Scale_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<StickerTest_Detail>(),
                DetailItemList = new List<StickerTest_Detail_Item>(),

            };

            try
            {
                _Provider = new StickerTestProvider(Common.ProductionDataAccessLayer);
                model.Scale_Source = _Provider.GetScales();

                _Provider = new StickerTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                model.Item_Source = _Provider.GetTestItems();

                if (IsNew)
                {
                    // 取得預設表身選項
                    var TestItems = _Provider.GetTestItems().Where(o => o.Result == "Pass").OrderBy(o => o.EvaluationItemSeq).ThenBy(o => o.EvaluationItemDescSeq);

                    // 產生StickerTest_Detail，預設填入值為Pass
                    foreach (var data in TestItems)
                    {
                        // 塞第二層 StickerTest_Detail
                        if (!model.DetailList.Any(o => o.EvaluationItem == data.EvaluationItem))
                        {

                            StickerTest_Detail detail = new StickerTest_Detail()
                            {
                                EvaluationItem = data.EvaluationItem,
                                Scale = "5",
                                Result = "Pass",
                            };
                            model.DetailList.Add(detail);
                        }

                        // 塞第三層 StickerTest_Detail_Item
                        StickerTest_Detail_Item detailItem = new StickerTest_Detail_Item()
                        {
                            EvaluationItem = data.EvaluationItem,
                            EvaluationItemDesc = data.EvaluationItemDesc,
                            Value = data.Value,
                            Result = data.Result,
                        };

                        model.DetailItemList.Add(detailItem);
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
        public StickerTest_ViewModel GetData(StickerTest_Request Req)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new StickerTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<StickerTest_Main> tmpList = new List<StickerTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new StickerTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new StickerTest_Request()
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

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new StickerTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    // 取得表身資料
                    model.DetailItemList = _Provider.GetDetailItemList(new StickerTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });


                    // 取得Article 下拉選單
                    _Provider = new StickerTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new StickerTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    string Subject = $"Residue/Ageing Test for Sticker Test/{model.Main.OrderID}/" +
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

        public StickerTest_ViewModel GetOrderInfo(string OrderID)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new StickerTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new StickerTest_Request() { OrderID = OrderID });


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
        public StickerTest_ViewModel NewSave(StickerTest_ViewModel Req, string MDivision, string UserID)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new StickerTestProvider(_ISQLDataTransaction);

                // 以EvaluationItem分組，相同EvaluationItem底下有任一Fail，則該群組 = Fail
                foreach (var item in Req.DetailList)
                {
                    string evaluationItem = item.EvaluationItem;

                    if (Req.DetailItemList.Where(o => o.EvaluationItem == evaluationItem && o.Result == "Fail").Any())
                    {
                        item.Result = "Fail";
                    }
                }

                // 判斷表頭Result，表身有任一Detail = Fail，即視作Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<StickerTest_Detail>();
                }

                // 新增，並取得ReportNo
                _Provider.Insert_StickerTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // StickerTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<StickerTest_Detail>();
                }

                _Provider.Processe_StickerTest_Detail(Req, UserID);
                _Provider.Processe_StickerTest_Detail_Item(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new StickerTest_Request()
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
        public StickerTest_ViewModel EditSave(StickerTest_ViewModel Req, string UserID)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new StickerTestProvider(_ISQLDataTransaction);

                // 以EvaluationItem分組，相同EvaluationItem底下有任一Fail，則該群組 = Fail
                foreach (var item in Req.DetailList)
                {
                    string evaluationItem = item.EvaluationItem;

                    if (Req.DetailItemList.Where(o => o.EvaluationItem == evaluationItem && o.Result == "Fail").Any())
                    {
                        item.Result = "Fail";
                    }
                }

                // 判斷表頭Result，表身有任一Detail = Fail，即視作Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<StickerTest_Detail>();
                }

                // 更新表頭，並取得ReportNo
                _Provider.Update_StickerTest(Req, UserID);

                // 更新表身
                _Provider.Processe_StickerTest_Detail(Req, UserID);
                _Provider.Processe_StickerTest_Detail_Item(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new StickerTest_Request()
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
        public StickerTest_ViewModel Delete(StickerTest_ViewModel Req)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);


            try
            {
                _Provider = new StickerTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.Delete_StickerTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new StickerTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                if (!string.IsNullOrEmpty(model.Main.ReportNo))
                {
                    model.Request = new StickerTest_Request()
                    {
                        ReportNo = model.Main.ReportNo,
                        BrandID = model.Main.BrandID,
                        SeasonID = model.Main.SeasonID,
                        StyleID = model.Main.StyleID,
                        Article = model.Main.Article,
                    };
                }
                else
                {
                    model.Result = true;
                }
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
        public StickerTest_ViewModel EncodeAmend(StickerTest_ViewModel Req, string UserID)
        {
            StickerTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new StickerTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_StickerTest(Req.Main, UserID);

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
        public StickerTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            StickerTest_ViewModel result = new StickerTest_ViewModel();

            string basefileName = "StickerTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
            string tmpName = string.Empty;

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {
                // 取得報表資料

                StickerTest_ViewModel model = this.GetData(new StickerTest_Request() { ReportNo = ReportNo });

                tmpName = $"Residue ,Ageing Test for Sticker Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                _Provider = new StickerTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                model.Item_Source = _Provider.GetTestItems();

                DataTable ReportTechnician = _Provider.GetReportTechnician(new StickerTest_Request() { ReportNo = ReportNo });

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                // 取得工作表上所有圖形物件
                Microsoft.Office.Interop.Excel.Shapes shapes = worksheet.Shapes;

                // 根據名稱，搜尋文字方塊物件
                Microsoft.Office.Interop.Excel.Shape ADIDAS_TextBox = shapes.Item("ADIDAS_TextBox");
                Microsoft.Office.Interop.Excel.Shape REEBOK_TextBox = shapes.Item("REEBOK_TextBox");
                Microsoft.Office.Interop.Excel.Shape PASS_TextBox = shapes.Item("PASS_TextBox");
                Microsoft.Office.Interop.Excel.Shape FAIL_TextBox = shapes.Item("FAIL_TextBox");

                // BrandID
                if (model.Main.BrandID.ToUpper() == "ADIDAS")
                {
                    ADIDAS_TextBox.TextFrame.Characters().Text = "V";
                }
                if (model.Main.BrandID.ToUpper() == "REEBOK")
                {
                    REEBOK_TextBox.TextFrame.Characters().Text = "V";
                }

                // Result
                if (model.Main.Result.ToUpper() == "PASS")
                {
                    PASS_TextBox.TextFrame.Characters().Text = "V";
                }
                else
                {
                    FAIL_TextBox.TextFrame.Characters().Text = "V";
                }

                string reportNo = model.Main.ReportNo;

                worksheet.Cells[1, 1] = $" PHX-AP0434 {model.Main.TestStandard} Test for Sticker";

                worksheet.Cells[3, 2] = model.Main.ReportNo;
                worksheet.Cells[3, 7] = model.Main.OrderID;

                worksheet.Cells[4, 2] = model.Main.FactoryID;
                worksheet.Cells[4, 7] = model.Main.SubmitDateText;

                worksheet.Cells[5, 2] = model.Main.StyleID;
                worksheet.Cells[5, 7] = model.Main.ReportDateText;

                worksheet.Cells[6, 2] = model.Main.Article;
                worksheet.Cells[6, 7] = model.Main.FabricRefNo;

                worksheet.Cells[7, 2] = model.Main.SeasonID;
                worksheet.Cells[7, 7] = model.Main.FabricColor;

                worksheet.Cells[8, 2] = model.Main.FabricDescription;

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[49, 4] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[48, 4];
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
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[39, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // TestAfterPicture 圖片
                if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[39, 5];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                int detailIdx = 0;

                // 表身處理 單筆和多筆分開
                if (model.DetailList.Any() && model.DetailList.Count >= 1)
                {
                    foreach (var detail in model.DetailList)
                    {
                        worksheet.Cells[11 + detailIdx, 1] = detail.EvaluationItem;
                        worksheet.Cells[11 + detailIdx, 2] = detail.Scale;

                        var DetailItemList = model.DetailItemList.Where(o => o.EvaluationItem == detail.EvaluationItem).ToList();

                        int detailItemIdx = 0;
                        foreach (var detail_Item in DetailItemList)
                        {                            
                            worksheet.Cells[11 + detailIdx + detailItemIdx, 4] = detail_Item.EvaluationItemDesc;

                            var ItemList = model.Item_Source.Where(o=> o.EvaluationItem == detail.EvaluationItem && o.EvaluationItemDesc == detail_Item.EvaluationItemDesc);

                            worksheet.Cells[11 + detailIdx + detailItemIdx, 6] = ItemList.Where(o => o.Result == "Pass").FirstOrDefault().Value;
                            worksheet.Cells[11 + detailIdx + detailItemIdx, 8] = ItemList.Where(o => o.Result == "Fail").FirstOrDefault().Value;

                            if (detail_Item.Result == "Pass")
                            {
                                worksheet.Cells[11 + detailIdx + detailItemIdx, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[11 + detailIdx + detailItemIdx, 9] = "V";
                            }
                            detailItemIdx++;
                        }

                        detailIdx += 5;
                    }
                }

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();

                foreach (char invalidChar in invalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
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
        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new StickerTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            StickerTest_ViewModel model = this.GetData(new StickerTest_Request() { ReportNo = ReportNo });
            string name = $"Residue ,Ageing Test for Sticker Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            StickerTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Residue/Ageing Test for Sticker Test/{model.Main.OrderID}/" +
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
