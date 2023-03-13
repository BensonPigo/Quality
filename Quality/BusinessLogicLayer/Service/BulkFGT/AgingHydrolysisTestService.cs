using ADOHelper.Utility;
using ADOHelper.Utility.Interface;
using BusinessLogicLayer.Helper;
using DatabaseObject.ViewModel.BulkFGT;
using Ict;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AgingHydrolysisTestService
    {
        public AgingHydrolysisTest_Provider _Provider;

        public AgingHydrolysisTest_ViewModel GetDefaultModel()
        {
            return new AgingHydrolysisTest_ViewModel()
            {
                Request = new AgingHydrolysisTest_Request(),
                MainData = new AgingHydrolysisTest_Main(),
                OrderID_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                MainList = new List<AgingHydrolysisTest_Main>(),
                DetailList = new List<AgingHydrolysisTest_Detail>(),
            };
        }

        public AgingHydrolysisTest_Detail_ViewModel GetDefaulDeailtModel()
        {
            _Provider = new AgingHydrolysisTest_Provider(Common.ProductionDataAccessLayer);
            return new AgingHydrolysisTest_Detail_ViewModel()
            {
                // 預設產生四筆
                MockupList = new List<AgingHydrolysisTest_Detail_Mockup>()
                {
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen1" },
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen2" },
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen3" },
                    new AgingHydrolysisTest_Detail_Mockup() { SpecimenName = "Specimen4" },
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
                    var existsData =_Provider.GetMainList(new AgingHydrolysisTest_Request() 
                    { 
                        OrderID = tmpOrders.FirstOrDefault().ID ,
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

                List<AgingHydrolysisTest_Main> tmpList = _Provider.GetMainList(new AgingHydrolysisTest_Request()
                {
                    BrandID = Req.BrandID,
                    SeasonID = Req.SeasonID,
                    StyleID = Req.StyleID,
                    Article = Req.Article,
                    OrderID = Req.OrderID,
                });

                if (tmpList.Any())
                {
                    // 取得下拉選單
                    model.OrderID_Source = new SetListItem().ItemListBinding(tmpList.Select(o => o.OrderID).Distinct().ToList());

                    // 若"有"傳入OrderID，則可以直接找出表頭表身明細
                    if (!string.IsNullOrEmpty(Req.OrderID))
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
                    AgingHydrolysisTestID = main.ID,
                });
                model.Request.BrandID = Req.MainData.BrandID;
                model.Request.SeasonID = Req.MainData.SeasonID;
                model.Request.StyleID = Req.MainData.StyleID;
                model.Request.Article = Req.MainData.Article;
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

                        // 有現存的、且MaterialType = "Mockup"，則帶現存四筆
                        if (nowMockupList.Any() && model.MainDetailData.MaterialType == "Mockup")
                        {
                            model.MockupList = nowMockupList;
                        }
                        // MaterialType != "Mockup"，清空
                        else if (model.MainDetailData.MaterialType != "Mockup")
                        {
                            model.MockupList = new List<AgingHydrolysisTest_Detail_Mockup>();
                        }
                        else
                        {
                            // 其餘則使用宣告時預設的四筆
                        }
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
                }, UserID);

                // 如果 AgingHydrolysisTest_Detail.MaterialType != "Mockup"，則刪除AgingHydrolysisTest_Detail_Mockup
                if (Req.MainDetailData.MaterialType != "Mockup")
                {
                    Req.MockupList = new List<AgingHydrolysisTest_Detail_Mockup>();
                }

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
    }
}
