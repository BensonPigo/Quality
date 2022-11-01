using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class HeatTransferWashService
    {
        public HeatTransferWashProvider _Provider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        public BaseResult Create(HeatTransferWash_ViewModel model, string MDivision, string userid, out string NewReportNo)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);
            NewReportNo = string.Empty;
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<HeatTransferWash_Detail_Result>() : model.Details;

                if (model.Details.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                {
                    model.Main.Result = "Fail";
                }
                else
                {
                    model.Main.Result = "Pass";
                }

                ctn = _Provider.Insert_HeatTransferWash(model.Main, MDivision, userid, out NewReportNo);

                if (ctn == 0)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrorMessage = "Create data fail.";
                    return result;
                }
                model.Main.ReportNo = NewReportNo;

                foreach (var detail in model.Details)
                {
                    detail.ReportNo = NewReportNo;
                    ctn = _Provider.Insert_HeatTransferWash_Detail(detail);
                    if (ctn == 0)
                    {
                        _ISQLDataTransaction.RollBack();
                        result.Result = false;
                        result.ErrorMessage = "Create detail Fail.";
                        return result;
                    }
                }

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public HeatTransferWash_ViewModel GetHeatTransferWash(HeatTransferWash_Request Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
            };

            try
            {
                _Provider = new HeatTransferWashProvider(Common.ProductionDataAccessLayer);
                List<SelectListItem> ReportNoList = _Provider.GetReportNo_Source(Req);

                if (!string.IsNullOrEmpty(Req.ReportNo))
                {
                    model.Main = _Provider.GetMainData(Req);
                    model.Details = _Provider.GetDetailData(Req.ReportNo).ToList();
                }
                else if (string.IsNullOrEmpty(Req.ReportNo) && ReportNoList.Any())
                {
                    model.Main = _Provider.GetMainData(new HeatTransferWash_Request()
                    {
                        ReportNo = ReportNoList.FirstOrDefault().Value,
                    });
                    model.Details = _Provider.GetDetailData(ReportNoList.FirstOrDefault().Value).ToList();
                }
                model.Request.ReportNo = model.Main.ReportNo;
                model.Request.BrandID = model.Main.BrandID;
                model.Request.SeasonID = model.Main.SeasonID;
                model.Request.StyleID = model.Main.StyleID;
                model.Request.Article = model.Main.Article;

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public List<Orders> GetOrders(Orders orders)
        {
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            try
            {
                orders.Category = "B";
                return _OrdersProvider.Get(orders).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<Order_Qty> GetDistinctArticle(Order_Qty order_Qty)
        {
            _OrderQtyProvider = new OrderQtyProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _OrderQtyProvider.GetDistinctArticle(order_Qty).ToList();
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
