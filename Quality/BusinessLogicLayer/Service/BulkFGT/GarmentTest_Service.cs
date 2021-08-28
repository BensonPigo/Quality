using BusinessLogicLayer.Interface.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ADOHelper.Utility;
using System.Data;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class GarmentTest_Service : IGarmentTest_Service
    {
        private IGarmentTestProvider _IGarmentTestProvider;
        private IGarmentTestDetailApperanceProvider _IGarmentTestDetailApperanceProvider;
        private IGarmentTestDetailFGPTProvider _IGarmentTestDetailFGPTProvider;
        private IGarmentTestDetailFGWTProvider _IGarmentTestDetailFGWTProvider;
        private IGarmentTestDetailProvider _IGarmentTestDetailProvider;
        private IGarmentTestDetailShrinkageProvider _IGarmentTestDetailShrinkageProvider;
        private IGarmentDetailSpiralityProvider _IGarmentDetailSpiralityProvider;

        public enum SelectType
        {
            OrderID,
            StyleID,
            Article,
            Season,
            Brand,
        }

        public GarmentTest_ViewModel GetSelectItemData(GarmentTest_ViewModel garmentTest_ViewModel, SelectType type)
        {
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            try
            {
                switch (type)
                {
                    case SelectType.Article:
                        result.Article_List = _IGarmentTestProvider.GetArticle(garmentTest_ViewModel).Select(x => x.Article).ToList();
                        break;
                    case SelectType.StyleID:
                        result.StyleID_Lsit = _IGarmentTestProvider.GetStyleID().Select(x => x.ID).ToList();
                        break;
                    case SelectType.Brand:
                        result.Brand_List = _IGarmentTestProvider.GetBrandID().Select(x => x.ID).ToList();
                        break;
                    case SelectType.Season:
                        result.Season_List = _IGarmentTestProvider.GetSeasonID().Select(x => x.ID).ToList();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public GarmentTest_Result GetGarmentTest(GarmentTest_Request garmentTest_ViewModel)
        {
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            GarmentTest_Result result = new GarmentTest_Result();
            try
            {
                // 抓取 garmentTest_ViewModel.Factory 撈取 M，並傳入Get_GarmentTest
                IFactoryProvider factoryProvider = new FactoryProvider(Common.ProductionDataAccessLayer);
                garmentTest_ViewModel.MDivisionid = factoryProvider.GetMDivisionID(garmentTest_ViewModel.Factory).FirstOrDefault().MDivisionID;

                if (string.IsNullOrEmpty(garmentTest_ViewModel.MDivisionid))
                {
                    result.Result = false;
                    result.ErrMsg = "Mdivision is empty.";
                    return result;
                }

                var query = _IGarmentTestProvider.Get_GarmentTest(garmentTest_ViewModel);
                if (!query.Any() || query.Count() == 0)
                {
                    throw new Exception("data not found!");
                }

                result.garmentTest = query.FirstOrDefault();

                // Detail
                result.garmentTest_Details = _IGarmentTestDetailProvider.Get_GarmentTestDetail(
                    new GarmentTest_ViewModel
                    {
                        ID = result.garmentTest.ID
                    }).ToList();

                result.SizeCodes = Get_SizeCode(result.garmentTest.OrderID, result.garmentTest.Article);
                result.Result = true;

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrMsg = ex.Message.ToString();
            }

            return result;
        }

        public List<string> Get_SizeCode(string OrderID, string Article)
        {
            List<string> result = new List<string>();
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailProvider.GetSizeCode(OrderID, Article).Select(x => x.SizeCode).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<GarmentTest_Detail_Shrinkage> Get_Shrinkage(Int64 ID, string No)
        {
            IList<GarmentTest_Detail_Shrinkage> result = new List<GarmentTest_Detail_Shrinkage>();
            _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailShrinkageProvider.Get_GarmentTest_Detail_Shrinkage(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<Garment_Detail_Spirality> Get_Spirality(Int64 ID, string No)
        {
            IList<Garment_Detail_Spirality> result = new List<Garment_Detail_Spirality>();
            _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentDetailSpiralityProvider.Get_Garment_Detail_Spirality(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<GarmentTest_Detail_Apperance_ViewModel> Get_Apperance(Int64 ID, string No)
        {
            IList<GarmentTest_Detail_Apperance_ViewModel> result = new List<GarmentTest_Detail_Apperance_ViewModel>();
            _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailApperanceProvider.Get_GarmentTest_Detail_Apperance(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<GarmentTest_Detail_FGWT_ViewModel> Get_FGWT(Int64 ID, string No)
        {
            IList<GarmentTest_Detail_FGWT_ViewModel> result = new List<GarmentTest_Detail_FGWT_ViewModel>();
            _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailFGWTProvider.Get_GarmentTest_Detail_FGWT(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<GarmentTest_Detail_FGPT_ViewModel> Get_FGPT(Int64 ID, string No)
        {
            IList<GarmentTest_Detail_FGPT_ViewModel> result = new List<GarmentTest_Detail_FGPT_ViewModel>();
            _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailFGPTProvider.Get_GarmentTest_Detail_FGPT(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public GarmentTest_ViewModel Save_GarmentTest(GarmentTest_ViewModel garmentTest_ViewModel,List<GarmentTest_Detail> detail)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);
                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (garmentTest_ViewModel.ID == 0) { emptyMsg += "Master ID cannot be 0." + Environment.NewLine; }
                if (string.IsNullOrEmpty(garmentTest_ViewModel.StyleID)) { emptyMsg += "Master StyleID cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(garmentTest_ViewModel.SeasonID)) { emptyMsg += "Master SeasonID cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(garmentTest_ViewModel.BrandID)) { emptyMsg += "Master BrandID cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(garmentTest_ViewModel.Article)) { emptyMsg += "Master Article cannot be empty." + Environment.NewLine; }

                foreach (var item in detail)
                {
                    if (item.No == 0) { emptyMsg += "detail No cannot be 0." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(item.MtlTypeID)) { emptyMsg += "detail MtlTypeID cannot be empty." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(item.AddName) && string.IsNullOrEmpty(item.EditName)) { emptyMsg += "detail AddName and EditName cannot be empty." + Environment.NewLine; }
                }

                #endregion
                result.SaveResult = _IGarmentTestProvider.Save_GarmentTest(garmentTest_ViewModel, detail);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.SaveResult = false;
                result.ErrMsg = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public GarmentTest_Detail_Result Generate_FGWT(GarmentTest_Detail_Result viewModel)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
                #region 判斷空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(viewModel.Detail.MtlTypeID)) { emptyMsg += "MtlTypeID cannot be empty" + Environment.NewLine; }
                //if (string.IsNullOrEmpty(garmentTest_ViewModel.StyleID)) { emptyMsg += "Master StyleID cannot be empty." + Environment.NewLine; }
                #endregion
            }
            catch (Exception ex)
            {

                throw;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public bool CheckOrderID(string OrderID, string BrandID, string SeasonID, string StyleID)
        {
            IOrdersProvider orders = new OrdersProvider(Common.ProductionDataAccessLayer);
            bool result = true;
            try
            {
                result = orders.Check_Style_ExistsOrder(OrderID, BrandID, SeasonID, StyleID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
