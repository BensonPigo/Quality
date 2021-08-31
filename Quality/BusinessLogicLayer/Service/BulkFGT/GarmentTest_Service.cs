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

        public enum DetailStatus
        {
            Encode,
            Amend,
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
                result.req = garmentTest_ViewModel;
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


        public GarmentTest_ViewModel SendMail(string ID, string No, string UserID)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel()
            {
                SaveResult = false,
                ErrMsg = "Err",
                Sender = UserID,
                SendDate = DateTime.Now.ToString("yyyy/MM/dd"),
            };

            return result;
        }

        public GarmentTest_ViewModel ReceiveMail(string ID, string No, string UserID)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel()
            {
                SaveResult = false,
                ErrMsg = "Err",
                Sender = UserID,
                SendDate = DateTime.Now.ToString("yyyy/MM/dd"),
            };

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
            // 僅傳入 List<GarmentTest_Detail> detail

            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
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

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    _ISQLDataTransaction.CloseConnection();
                    return result;
                }

                #endregion

                _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);
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

        public GarmentTest_Result Generate_FGWT(GarmentTest_Detail_Result viewModel)
        {
            GarmentTest_Result result = new GarmentTest_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
                #region 判斷空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(viewModel.Detail.MtlTypeID)) { emptyMsg += "MtlTypeID cannot be empty" + Environment.NewLine; }
                if (_IGarmentTestDetailFGWTProvider.Chk_FGWTExists(viewModel.Detail) == true) { emptyMsg += "Data already exists!!"; }
                #endregion

                result.Result = _IGarmentTestDetailFGWTProvider.Create_FGWT(viewModel.Main, viewModel.Detail);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrMsg = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public string Get_LastResult(string ID)
        {
            string result = string.Empty;
            try
            {
                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailProvider.Get_LastResult(ID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public GarmentTest_Detail_Result Encode_Detail(GarmentTest_Detail_Result viewModel, string GroupName, DetailStatus status)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                switch (status)
                {
                    case DetailStatus.Encode:
                        _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);

                        // 代表所有result 有任一個是Fail 就寄信
                        if (_IGarmentTestDetailProvider.Chk_AllResult(viewModel.Detail.ID.ToString(), viewModel.Detail.No.ToString()) == false)
                        {

                        }
                        result.Result = _IGarmentTestDetailProvider.Encode_GarmentTestDetail(ID, "Confirmed");
                        break;
                    case DetailStatus.Amend:
                        _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                        result.Result = _IGarmentTestDetailProvider.Encode_GarmentTestDetail(ID, "New");
                        break;
                    default:
                        break;
                }
              
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrMsg = ex.Message;
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

        public GarmentTest_ViewModel DeleteDetail(string ID, string No)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(ID)) { emptyMsg += "Master ID cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(No)) { emptyMsg += "No cannot be 0 or null" + Environment.NewLine; }

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    return result;
                }
                #endregion

                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                int deleteCnt = _IGarmentTestDetailProvider.Delete_GarmentTestDetail(ID, No);
                result.SaveResult = true;
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

        public GarmentTest_Detail_Result Save_GarmentTestDetail(GarmentTest_Detail_Result source)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                result.Result = true;
                string errMsg = string.Empty;

                // Detail Save
                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailProvider.Update_GarmentTestDetail(source.Detail) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update detail is empty.";
                    return result;
                }

                // Shrinkage Save
                _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(_ISQLDataTransaction);
                #region 檢查空值
                foreach (var item in source.Shrinkages)
                {
                    bool isAllEmpty = (item.AfterWash1 == 0) && (item.AfterWash2 == 0) && (item.AfterWash3 == 0);

                    if (item.BeforeWash == 0 && isAllEmpty == false)
                    {
                        result.Result = false;
                        result.ErrMsg = @"<BeforeWash> can not be empty or 0 !!";
                        return result;
                    }
                }
                #endregion

                if (_IGarmentTestDetailShrinkageProvider.Update_GarmentTestShrinkage(source.Shrinkages) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update Shrinkage is empty.";
                    return result;
                }

                // Spirality Save
                _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(_ISQLDataTransaction);
                if (_IGarmentDetailSpiralityProvider.Update_Spirality(source.Spiralities) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update Spirality is empty.";
                    return result;
                }

                // Apperance Save 
                _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailApperanceProvider.Update_Apperance(source.Apperance) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update Apperance is empty.";
                    return result;
                }

                // FGPT Save
                _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailFGPTProvider.Update_FGPT(source.FGPT) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update FGPT is empty.";
                    return result;
                }

                // FGWT Save
                _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailFGWTProvider.Update_FGWT(source.FGWT) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update FGWT is empty.";
                    return result;
                }
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrMsg = ex.Message;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }
    }
}
