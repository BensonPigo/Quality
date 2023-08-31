using ADOHelper.Utility;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class WickingHeightTestService
    {
        private WickingHeightTestProvider _Provider;

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

        public WickingHeightTest_ViewModel GetReport(string ReportNo, bool isPDF)
        {
            WickingHeightTest_ViewModel result = new WickingHeightTest_ViewModel();

            string basefileName = "WickingHeightTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            // Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            return result;
        }
    }
}
