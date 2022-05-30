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
using DatabaseObject.ManufacturingExecutionDB;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Excel = Microsoft.Office.Interop.Excel;
using Sci;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Library;
using System.Windows.Forms;

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

        public enum ReportType
        {
            Wash_Test_2018,
            Wash_Test_2020,
            Physical_Test,
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

        public List<string> Get_SizeCode(string StyleID, string SeasonID, string BrandID)
        {
            List<string> result = new List<string>();
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                result = _IGarmentTestDetailProvider.GetSizeCode(StyleID, SeasonID, BrandID).Select(x => x.SizeCode).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public List<string> Get_Scales()
        {
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IGarmentTestDetailProvider.GetScales();
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
       

        #region Save & Update & Encode & Delete
        public GarmentTest_Result Generate_FGWT(GarmentTest_ViewModel Main, GarmentTest_Detail_ViewModel Detail)
        {
            GarmentTest_Result result = new GarmentTest_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
                #region 判斷空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(Detail.MtlTypeID)) { emptyMsg += "MtlTypeID cannot be empty" + Environment.NewLine; }
                if (Detail.Above50NaturalFibres == false && Detail.Above50SyntheticFibres == false) { emptyMsg += " < 50% natural fibres>, <50% synthetic fibres(ex. polyester)> have to select one!" + Environment.NewLine; }
                if (_IGarmentTestDetailFGWTProvider.Chk_FGWTExists(Detail) == true) { emptyMsg += "Data already exists!!"; }
                #endregion

                result.Result = _IGarmentTestDetailFGWTProvider.Create_FGWT(Main, Detail);
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

        public GarmentTest_ViewModel Save_GarmentTest(GarmentTest_ViewModel garmentTest_ViewModel, List<GarmentTest_Detail_ViewModel> detail, string UserID)
        {
            // 僅傳入 List<GarmentTest_Detail> detail

            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            string mainServerName = string.Empty;
            string extendServerName = string.Empty;

            _IGarmentTestProvider = new GarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
            extendServerName = _IGarmentTestProvider.CheckInstance();
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            mainServerName = _IGarmentTestProvider.CheckInstance();

            bool sameInstance = mainServerName == extendServerName ? true : false;

            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                result.SaveResult = true;
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
                    // if (string.IsNullOrEmpty(item.AddName) && string.IsNullOrEmpty(item.EditName)) { emptyMsg += "detail AddName and EditName cannot be empty." + Environment.NewLine; }
                }

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    return result;
                }

                #endregion

                // 先將表身資料存檔
                _IGarmentTestProvider.Save_GarmentTest(garmentTest_ViewModel, detail, UserID, sameInstance);

                // 再判斷Detail的Result
                foreach (var item in detail)
                {
                    if (_IGarmentTestDetailProvider.Update_GarmentTestDetail_Result(item.ID.ToString(), item.No.ToString()) == false)
                    {
                        result.SaveResult = false;
                        result.ErrMsg = "update GarmentTest detail result is empry";
                        return result;
                    }
                }
                if (detail.Any())
                {
                    _IGarmentTestProvider.Update_GarmentTest_Result(detail.Select(s => s.ID.ToString()).First());
                }
            }
            catch (Exception ex)
            {
                result.SaveResult = false;
                result.ErrMsg = ex.Message;
            }

            return result;
        }

        public GarmentTest_ViewModel Import_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem)
        {
            // 僅傳入 List<GarmentTest_Detail> detail

            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                newItem.TestDetail = newItem.TestUnit;
                _IGarmentTestProvider.Save_New_FGPT_Item(newItem);
                result.SaveResult = true;
            }
            catch (Exception ex)
            {
                result.SaveResult = false;
                result.ErrMsg = ex.Message;
            }

            return result;
        }

        public GarmentTest_ViewModel Delete_Original_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem)
        {
            // 僅傳入 List<GarmentTest_Detail> detail

            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                newItem.TestDetail = newItem.TestUnit;
                _IGarmentTestProvider.Delete_Original_FGPT_Item(newItem);
                result.SaveResult = true;
            }
            catch (Exception ex)
            {
                result.SaveResult = false;
                result.ErrMsg = ex.Message;
            }

            return result;
        }

        public GarmentTest_Detail_Result Save_GarmentTestDetail(GarmentTest_Detail_Result source)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {

                string mainServerName = string.Empty;
                string extendServerName = string.Empty;

                _IGarmentTestProvider = new GarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                extendServerName = _IGarmentTestProvider.CheckInstance();

                _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
                mainServerName = _IGarmentTestProvider.CheckInstance();

                bool sameInstance = mainServerName == extendServerName ? true : false;

                result.Result = true;
                string errMsg = string.Empty;

                #region Shrinkage save
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
                #endregion

                #region Spirality Save
                _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(_ISQLDataTransaction);
                if (source.Spiralities != null && source.Spiralities.Count > 0)
                {
                    if (_IGarmentDetailSpiralityProvider.Update_Spirality(source.Spiralities) == false)
                    {
                        _ISQLDataTransaction.RollBack();
                        result.Result = false;
                        result.ErrMsg = "Update Spirality is empty.";
                        return result;
                    }
                }
                #endregion

                #region Apperance Save 
                _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailApperanceProvider.Update_Apperance(source.Apperance) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update Apperance is empty.";
                    return result;
                }
                #endregion

                #region FGPT Save
                _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(_ISQLDataTransaction);
                if (_IGarmentTestDetailFGPTProvider.Update_FGPT(source.FGPT) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update FGPT is empty.";
                    return result;
                }
                #endregion

                #region FGWT Save

                if (source.FGWT != null)
                {
                    _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(_ISQLDataTransaction);
                    if (_IGarmentTestDetailFGWTProvider.Update_FGWT(source.FGWT) == false)
                    {
                        _ISQLDataTransaction.RollBack();
                        result.Result = false;
                        result.ErrMsg = "Update FGWT is empty.";
                        return result;
                    }
                }
                #endregion

                #region Detail Save
                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                // 檢查必輸欄位
                if (source.Detail.LineDry == false && source.Detail.TumbleDry == false && source.Detail.HandWash == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "<Line Dry>, <Tumble Dry>, <Hand Wash> have to select one!";
                    return result;
                }

                if (_IGarmentTestDetailProvider.Update_GarmentTestDetail(source.Detail, sameInstance) == false)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrMsg = "Update detail is empty.";
                    return result;
                }
                #endregion

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

        public GarmentTest_ViewModel DeleteDetail(string ID, string No)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);

            string mainServerName = string.Empty;
            string extendServerName = string.Empty;

            _IGarmentTestProvider = new GarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
            extendServerName = _IGarmentTestProvider.CheckInstance();

            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            mainServerName = _IGarmentTestProvider.CheckInstance();

            bool sameInstance = mainServerName == extendServerName ? true : false;
            try
            {
                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(ID)) { emptyMsg += "Master ID cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(No)) { emptyMsg += "No cannot be 0 or null" + Environment.NewLine; }
                GarmentTest_Detail_ViewModel detail = _IGarmentTestDetailProvider.Get(ID, No, sameInstance).First();
                if (detail.Status.ToUpper() == "Confirmed")
                {
                    emptyMsg += "Encode data cannot delete.";
                } 

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    return result;
                }
                #endregion

                int deleteCnt = _IGarmentTestDetailProvider.Delete_GarmentTestDetail(ID, No, sameInstance);
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

        public GarmentTest_Detail_Result Encode_Detail(string ID, string No, DetailStatus status)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            result.sentMail = false;
            result.Result = true;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                switch (status)
                {
                    case DetailStatus.Encode:
                        _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                        _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);

                    
                        // 重新判斷Result
                        if (_IGarmentTestDetailProvider.Encode_GarmentTestDetail(ID, No, "Confirmed") == false ||
                            _IGarmentTestDetailProvider.Update_GarmentTestDetail_Result(ID, No) == false ||
                            _IGarmentTestProvider.Update_GarmentTest_Result(ID) == false)
                        {
                            result.Result = false;
                        }

                        // all result 有任一個是Fail 就寄信
                        result.sentMail = _IGarmentTestDetailProvider.Chk_AllResult(ID, No);

                        break;
                    case DetailStatus.Amend:
                        _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                        _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);

                        // 重新判斷Result
                        if (_IGarmentTestDetailProvider.Encode_GarmentTestDetail(ID, No, "New") == false ||
                            _IGarmentTestDetailProvider.Update_GarmentTestDetail_Result_Amend(ID, No) == false ||
                            _IGarmentTestProvider.Update_GarmentTest_Result(ID) == false)
                        {
                            result.Result = false;
                        }
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
       
        #endregion

        // Sent Mail 
        #region Sent Mail
        public GarmentTest_ViewModel SendMail(string ID, string No, string UserID)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(ID)) { emptyMsg += "Master ID cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(No)) { emptyMsg += "No cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(UserID)) { emptyMsg += "UserID cannot be 0 or null" + Environment.NewLine; }

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    return result;
                }
                #endregion

                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                result.SaveResult = _IGarmentTestDetailProvider.Update_Sender(ID, No, UserID);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.SaveResult = false;
                result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public GarmentTest_Result SentMail(string ID, string No, List<Quality_MailGroup> mailGroups)
        {
            GarmentTest_Result result = new GarmentTest_Result();
            string ToAddress = string.Empty;
            string CCAddress = string.Empty;
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                foreach (var item in mailGroups)
                {
                    ToAddress += item.ToAddress + ";";
                    CCAddress += item.CcAddress + ";";
                }

                if (string.IsNullOrEmpty(ToAddress) == true)
                {
                    result.Result = false;
                    result.ErrMsg = "To email address is empty!";
                    return result;
                }

                DataTable dtContent = _IGarmentTestDetailProvider.Get_Mail_Content(ID, No);
                string strHtml = MailTools.DataTableChangeHtml(dtContent, out System.Net.Mail.AlternateView plainView);

                SendMail_Request request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CCAddress,
                    Subject = "Garment Test – Test Fail",
                    Body = strHtml,
                    alternateView = plainView,
                };

                MailTools.SendMail(request);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }

        public GarmentTest_ViewModel ReceiveMail(string ID, string No, string UserID)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                #region 判斷是否空值
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(ID)) { emptyMsg += "Master ID cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(No)) { emptyMsg += "No cannot be 0 or null" + Environment.NewLine; }
                if (string.IsNullOrEmpty(UserID)) { emptyMsg += "UserID cannot be 0 or null" + Environment.NewLine; }

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    result.SaveResult = false;
                    result.ErrMsg = emptyMsg;
                    return result;
                }
                #endregion

                _IGarmentTestDetailProvider = new GarmentTestDetailProvider(_ISQLDataTransaction);
                result.SaveResult = _IGarmentTestDetailProvider.Update_Receive(ID, No, UserID);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.SaveResult = false;
                result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }
        #endregion

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


                bool chk = CheckOrderID(result.garmentTest.OrderID, result.garmentTest.BrandID, result.garmentTest.SeasonID, result.garmentTest.StyleID);
                if (chk)
                {
                    result.SizeCodes = Get_SizeCode(result.garmentTest.OrderID, result.garmentTest.Article);
                }
                else
                {
                    result.SizeCodes = Get_SizeCode(result.garmentTest.StyleID, result.garmentTest.SeasonID, result.garmentTest.BrandID);
                }

                result.req = garmentTest_ViewModel;
                result.Result = true;

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }

        public GarmentTest_ViewModel Get_Main(string ID)
        {
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                GarmentTest_Request garment = new GarmentTest_Request
                {
                    ID = ID,
                };
                return _IGarmentTestProvider.Get_GarmentTest(garment).First();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        // Get Detail Data
        #region Get Detail Data
        public GarmentTest_Detail_Result Get_All_Detail(string ID, string No)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();

            result.Main = Get_Main(ID);
            result.Detail = Get_Detail(ID, No);
            result.Shrinkages = Get_Shrinkage(ID, No).ToList();
            result.Spiralities = Get_Spirality(ID, No).ToList();
            result.Apperance = Get_Apperance(ID, No).ToList();
            result.FGPT = Get_FGPT(ID, No).ToList();
            result.FGWT = Get_FGWT(ID, No).ToList();
            result.Scales = Get_Scales();

            return result;
        }

        public IList<GarmentTest_Detail_Shrinkage> Get_Shrinkage(string ID, string No)
        {
            _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IGarmentTestDetailShrinkageProvider.Get_GarmentTest_Detail_Shrinkage(ID, No);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public GarmentTest_Detail_ViewModel Get_Detail(string ID, string No)
        {
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                string mainServerName = string.Empty;
                string extendServerName = string.Empty;

                _IGarmentTestProvider = new GarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                extendServerName = _IGarmentTestProvider.CheckInstance();
                _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
                mainServerName = _IGarmentTestProvider.CheckInstance();

                bool sameInstance = mainServerName == extendServerName ? true : false;

                return _IGarmentTestDetailProvider.Get(ID, No, sameInstance).First();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<Garment_Detail_Spirality> Get_Spirality(string ID, string No)
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

        public IList<GarmentTest_Detail_Apperance_ViewModel> Get_Apperance(string ID, string No)
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

        public IList<GarmentTest_Detail_FGWT_ViewModel> Get_FGWT(string ID, string No)
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

        public IList<GarmentTest_Detail_FGPT_ViewModel> Get_FGPT(string ID, string No)
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
        #endregion


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

        private void RowHeight(Excel.Worksheet worksheet, int row, string strComment)
        {
            if (strComment.Length > 15)
            {
                decimal n = Math.Ceiling(strComment.Length / (decimal)15.0) * (decimal)12.25;
                worksheet.Range[$"A{row}", $"A{row}"].RowHeight = n;
            }
        }

        private string AddShrinkageUnit_18(DataTable dt, int row, int columns)
        {
            string strValie = string.Empty;
            if (dt.Rows.Count > 0)
            {
                strValie = dt.Rows[row][columns].ToString();
                if (((string.Compare(dt.Columns[columns].ColumnName, "Shrinkage1", true) == 0) ||
                    (string.Compare(dt.Columns[columns].ColumnName, "Shrinkage2", true) == 0) ||
                    (string.Compare(dt.Columns[columns].ColumnName, "Shrinkage3", true) == 0)) &&
                    !MyUtility.Check.Empty(strValie))
                {
                    strValie = strValie + "%";
                }
            }

            return strValie;
        }

        public GarmentTest_Detail_Result ToReport(string ID, string No, ReportType type, bool IsToPDF, bool test = false)
        {
            GarmentTest_Detail_Result all_Data = Get_All_Detail(ID, No);
            all_Data.reportPath = string.Empty;
            P04Data data = new P04Data
            {
                DateSubmit = all_Data.Detail.SubmitDate,
                NumArriveQty = all_Data.Detail.ArrivedQty,
                TxtSize = all_Data.Detail.SizeCode,
                RdbtnLine = all_Data.Detail.LineDry,
                RdbtnHand = all_Data.Detail.HandWash,
                RdbtnTumble = all_Data.Detail.TumbleDry,
                ComboTemperature = all_Data.Detail.Temperature.ToString(),
                ComboMachineModel = all_Data.Detail.Machine,
                TxtFibreComposition = all_Data.Detail.Composition,
                ComboNeck = all_Data.Detail.Neck == true ? "Yes" : "No",
                TxtLotoFactory = all_Data.Detail.LOtoFactory,
                Remark = all_Data.Detail.Remark,
            };


            bool IsNewData = all_Data.Apperance.Count == 9 ? false : true;

            IStyleProvider styleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            string StyleName = styleProvider.GetStyleName(all_Data.Main.StyleID, all_Data.Main.SeasonID, all_Data.Main.BrandID);

            IOrdersProvider ordersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            Orders orders = new Orders();
            if (!string.IsNullOrEmpty(all_Data.Detail.OrderID))
            {
                var query = ordersProvider.Get(new Orders() { ID = all_Data.Detail.OrderID });
                if (query.Any())
                {
                    orders = query.FirstOrDefault();
                }
            }

            _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);
            DataTable dtShrinkages = _IGarmentTestDetailShrinkageProvider.Get_dt_Shrinkage(ID, No);

            switch (type)
            {
                case ReportType.Wash_Test_2018:
                    #region Print Wash 2018

                    if (all_Data.Apperance.Count == 0 || all_Data.Shrinkages.Count == 0)
                    {
                        all_Data.ErrMsg = " Datas not found!";
                        all_Data.Result = false;
                        return all_Data;
                    }

                    try
                    {
                        if (!test)
                        {
                            if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                            {
                                System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                            }

                            if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                            {
                                System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                            }
                        }

                        string openfilepath;
                        string basefileName_2018 = "WashTest_2018";
                        if (test)
                        {
                            openfilepath = "C:\\Willy_Repository\\Quality_KPI\\Quality\\Quality\\bin\\XLT\\WashTest_2018.xltx";
                        }
                        else
                        {
                            openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName_2018}.xltx";
                        }

                        Excel.Application objApp = MyUtility.Excel.ConnectExcel(openfilepath);
                        objApp.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                        Excel.Worksheet worksheet = objApp.ActiveWorkbook.Worksheets[1]; // 取得工作表

                        // WashName 調整次數 710(1,10,15), 701(1,3)
                        string WashName = all_Data.Main.WashName;
                        if (WashName == "710")
                        {
                            // Actual Shrinkage % 
                            // Top
                            worksheet.Cells[18, 8] = "10.wash";
                            worksheet.Cells[18, 10] = "15.wash";
                            // INNER
                            worksheet.Cells[26, 8] = "10.wash";
                            worksheet.Cells[26, 10] = "15.wash";
                            // OUTER
                            worksheet.Cells[34, 8] = "10.wash";
                            worksheet.Cells[34, 10] = "15.wash";
                            // BOTTOM
                            worksheet.Cells[44, 8] = "10.wash";
                            worksheet.Cells[44, 10] = "15.wash";

                            //After Wash Appearance Check list
                            worksheet.Cells[63, 6] = "10.wash";
                            worksheet.Cells[63, 8] = "15.wash";

                            // Team Wear
                            worksheet.Cells[15, 4] = "V";
                            worksheet.Cells[15, 8] = "15";
                        }
                        else if (WashName == "701")
                        {
                            // Actual Shrinkage % 
                            // Top
                            worksheet.Cells[18, 10] = "";
                            // INNER
                            worksheet.Cells[26, 10] = "";
                            // OUTER
                            worksheet.Cells[34, 10] = "";
                            // BOTTOM
                            worksheet.Cells[44, 10] = "";

                            //After Wash Appearance Check list
                            worksheet.Cells[63, 8] = "";

                            // Team Wear
                            worksheet.Cells[15, 4] = "";
                            worksheet.Cells[15, 8] = "3";
                        }

                        if (data.DateSubmit.HasValue)
                        {
                            worksheet.Cells[4, 4] = MyUtility.Convert.GetDate(data.DateSubmit.Value).Value.Year + "/" + MyUtility.Convert.GetDate(data.DateSubmit.Value).Value.Month + "/" + MyUtility.Convert.GetDate(data.DateSubmit.Value).Value.Day;
                        }

                        if (!MyUtility.Check.Empty(all_Data.Detail.inspdate))
                        {
                            worksheet.Cells[4, 7] = MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Year + "/" + MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Month + "/" + MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Day;
                        }

                        worksheet.Cells[4, 9] = MyUtility.Convert.GetString(all_Data.Main.OrderID);
                        worksheet.Cells[4, 11] = MyUtility.Convert.GetString(all_Data.Main.BrandID);
                        worksheet.Cells[6, 4] = MyUtility.Convert.GetString(all_Data.Main.StyleID);
                        worksheet.Cells[7, 8] = orders.CustPONO;
                        worksheet.Cells[7, 4] = MyUtility.Convert.GetString(all_Data.Main.Article);
                        worksheet.Cells[6, 8] = StyleName;
                        worksheet.Cells[8, 8] = MyUtility.Convert.GetDecimal(data.NumArriveQty.Value);

                        string sendDate = Convert.ToDateTime(orders.BuyerDelivery).ToShortDateString();
                        worksheet.Cells[8, 4] = sendDate;
                        worksheet.Cells[8, 10] = MyUtility.Convert.GetString(data.TxtSize);

                        worksheet.Cells[11, 4] = data.RdbtnLine == true ? "V" : string.Empty;
                        worksheet.Cells[12, 4] = data.RdbtnTumble == true ? "V" : string.Empty;
                        worksheet.Cells[13, 4] = data.RdbtnHand == true ? "V" : string.Empty;
                        worksheet.Cells[11, 8] = data.ComboTemperature + "˚C ";
                        worksheet.Cells[12, 8] = data.ComboMachineModel;
                        worksheet.Cells[13, 8] = data.TxtFibreComposition;
                        // Remark
                        worksheet.Cells[14, 4] = data.Remark;

                        string remark = data.Remark == null ? string.Empty : data.Remark;
                        int rowCtn = (remark.Length / 85) + 1;

                        worksheet.get_Range("14:14", Type.Missing).RowHeight = 15.75 * rowCtn;

                        #region 舊資料
                        if (!IsNewData)
                        {
                            #region 照片
                            if (all_Data.Detail.TestBeforePicture != null)
                            {
                                Excel.Range cell = worksheet.Cells[82, 2];
                                string imageName = $"{Guid.NewGuid()}.jpg";
                                string imgPath;
                                if (test)
                                {
                                    imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                }
                                else
                                {
                                    imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                }

                                using (var imageFile = new FileStream(imgPath, FileMode.Create))
                                {
                                    imageFile.Write(all_Data.Detail.TestBeforePicture, 0, all_Data.Detail.TestBeforePicture.Length);
                                    imageFile.Flush();
                                }
                                worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 328, 247);
                            }

                            if (all_Data.Detail.TestAfterPicture != null)
                            {
                                Excel.Range cell = worksheet.Cells[82, 7];
                                string imageName = $"{Guid.NewGuid()}.jpg";
                                string imgPath;
                                if (test)
                                {
                                    imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                }
                                else
                                {
                                    imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                }

                                using (var imageFile = new FileStream(imgPath, FileMode.Create))
                                {
                                    imageFile.Write(all_Data.Detail.TestAfterPicture, 0, all_Data.Detail.TestAfterPicture.Length);
                                    imageFile.Flush();
                                }
                                worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 300, 248);
                            }
                            #endregion

                            #region 最下面 Signature

                            bool ApperanceRejected = all_Data.Apperance.Where(x => x.Wash1.EqualString("Rejected")
                                                        || x.Wash2.EqualString("Rejected")
                                                        || x.Wash3.EqualString("Rejected")
                                                        || x.Wash4.EqualString("Rejected")
                                                        || x.Wash5.EqualString("Rejected")).ToList().Count() > 0;

                            bool ApperanceAccepted = all_Data.Apperance.Where(x => x.Wash1.EqualString("Accepted")
                                                      || x.Wash2.EqualString("Accepted")
                                                      || x.Wash3.EqualString("Accepted")
                                                      || x.Wash4.EqualString("Accepted")
                                                      || x.Wash5.EqualString("Accepted")).ToList().Count() > 0;

                            if ((MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("P") || ApperanceAccepted) && !ApperanceRejected)
                            {
                                worksheet.Cells[77, 4] = "V";
                            }
                            else if (MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("F") || ApperanceRejected)
                            {
                                worksheet.Cells[77, 6] = "V";
                            }

                            #endregion

                            #region 插入圖片與Technician名字

                            // ToPDF
                            if (IsToPDF)
                            {
                                string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.inspector}'
and t.GarmentTest=1
";
                                string technicianName = string.Empty;
                                DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0 && dtTechnicianInfo.Rows[0]["SignaturePic"] != null)
                                {
                                    technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                    byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];
                                  
                                    string imageName = $"{Guid.NewGuid()}.jpg";
                                    string imgPath;

                                    if (test)
                                    {
                                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                    }
                                    else
                                    {
                                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                    }

                                    // Name
                                    worksheet.Cells[78, 9] = technicianName;
                                    Excel.Range cell = worksheet.Cells[76, 9];

                                    using (MemoryStream ms = new MemoryStream(imgData))
                                    {
                                        Image img = Image.FromStream(ms);
                                        img.Save(imgPath);
                                    }
                                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);
                                }
                                else
                                {
                                    worksheet.Cells[78, 9] = MyUtility.Convert.GetString(all_Data.Detail.GarmentTest_Detail_Inspector);
                                }
                            }

                            // ToExcel
                            if (!IsToPDF)
                            {
                                worksheet.Cells[74, 8] = string.Empty;
                            }
                            #endregion

                            #region After Wash Appearance Check list
                            string tmpAR;

                            worksheet.Cells[65, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First());

                            worksheet.get_Range("65:65", Type.Missing).Rows.AutoFit();

                            // 大約21個字換行
                            int widhthBase = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First()).Length / 20;

                            worksheet.get_Range("65:65", Type.Missing).RowHeight = widhthBase == 0 ? 28 : 28 * widhthBase;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 8] = tmpAR;
                            }

                            string strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 65, strComment);
                            worksheet.Cells[65, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[66, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[66, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[66, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[66, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 66, strComment);
                            worksheet.Cells[66, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[67, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[67, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 67, strComment);
                            worksheet.Cells[67, 10] = strComment;

                            worksheet.Cells[68, 3] = all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First().ToString(); // type;

                            // 大約21個字換行
                            int widhthBase2 = all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First().ToString().Length / 20;

                            worksheet.get_Range("68:68", Type.Missing).RowHeight = widhthBase2 == 0 ? 28 : 28 * widhthBase2;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 68, strComment);
                            worksheet.Cells[68, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 69, strComment);
                            worksheet.Cells[69, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 70, strComment);
                            worksheet.Cells[70, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 71, strComment);
                            worksheet.Cells[71, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 72, strComment);
                            worksheet.Cells[72, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[73, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[73, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[73, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[73, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[73, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[73, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[73, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[73, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[73, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 73, strComment);
                            worksheet.Cells[73, 10] = strComment;
                            #endregion

                            #region Spirality

                            if (all_Data.Spiralities.Where(x => x.Location.EqualString("B")).Any())
                            {
                                var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("B")).First();
                                worksheet.Cells[58, 6] = dr.MethodA;
                                worksheet.Cells[59, 6] = dr.MethodB;
                                worksheet.Cells[60, 6] = dr.CM;
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A56:A60", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (all_Data.Spiralities.Where(x => x.Location.EqualString("T")).Any())
                            {
                                var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("T")).First();
                                worksheet.Cells[53, 6] = dr.MethodA;
                                worksheet.Cells[54, 6] = dr.MethodB;
                                worksheet.Cells[55, 6] = dr.CM;
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A51:A55", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }
                            #endregion

                            // Streched Neck Opening is OK according to size spec?
                            if (data.ComboNeck.EqualString("Yes"))
                            {
                                worksheet.Cells[42, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[42, 11] = "V";
                            }

                            #region Shrinkage

                            if (all_Data.Shrinkages.Where(x => x.Location.EqualString("B")).Any())
                            {
                                DataTable dt_B = dtShrinkages.Select("Location = 'B'").CopyToDataTable();



                                // 超過5個測量點則新增行數
                                if (dt_B.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt_B.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A50:A50", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt_B.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt_B.Columns.Count; c++)
                                    {
                                        worksheet.Cells[46 + r, c] = this.AddShrinkageUnit_18(dt_B, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A44:A51", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'O'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'O'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A40:A40", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[35+ r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A34:A41", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'I'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'I'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A32:A32", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[28 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A26:A33", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'T'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'T'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A24:A24", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[20 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A18:A25", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }
                            #endregion

                        }
                        #endregion

                        #region 新資料
                        if (IsNewData)
                        {
                            #region 照片
                            if (all_Data.Detail.TestBeforePicture != null)
                            {
                                Excel.Range cell = worksheet.Cells[82, 2];
                                string imageName = $"{Guid.NewGuid()}.jpg";
                                string imgPath;
                                if (test)
                                {
                                    imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                }
                                else
                                {
                                    imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                }

                                using (var imageFile = new FileStream(imgPath, FileMode.Create))
                                {
                                    imageFile.Write(all_Data.Detail.TestBeforePicture, 0, all_Data.Detail.TestBeforePicture.Length);
                                    imageFile.Flush();
                                }
                                worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 328, 247);
                            }

                            if (all_Data.Detail.TestAfterPicture != null)
                            {
                                Excel.Range cell = worksheet.Cells[82, 7];
                                string imageName = $"{Guid.NewGuid()}.jpg";
                                string imgPath;
                                if (test)
                                {
                                    imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                }
                                else
                                {
                                    imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                }

                                using (var imageFile = new FileStream(imgPath, FileMode.Create))
                                {
                                    imageFile.Write(all_Data.Detail.TestAfterPicture, 0, all_Data.Detail.TestAfterPicture.Length);
                                    imageFile.Flush();
                                }
                                worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 300, 248);
                            }
                            #endregion

                            worksheet.get_Range("66:66", Type.Missing).Delete();

                            #region 最下面 Signature

                            bool ApperanceRejected = all_Data.Apperance.Where(x => x.Wash1.EqualString("Rejected")
                                                        || x.Wash2.EqualString("Rejected")
                                                        || x.Wash3.EqualString("Rejected")
                                                        || x.Wash4.EqualString("Rejected")
                                                        || x.Wash5.EqualString("Rejected")).ToList().Count() > 0;

                            bool ApperanceAccepted = all_Data.Apperance.Where(x => x.Wash1.EqualString("Accepted")
                                                      || x.Wash2.EqualString("Accepted")
                                                      || x.Wash3.EqualString("Accepted")
                                                      || x.Wash4.EqualString("Accepted")
                                                      || x.Wash5.EqualString("Accepted")).ToList().Count() > 0;


                            if ((MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("P") || ApperanceAccepted) && !ApperanceRejected)
                            {
                                worksheet.Cells[76, 4] = "V";
                            }
                            else if (MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("F") || ApperanceRejected)
                            {
                                worksheet.Cells[76, 6] = "V";
                            } 
                            #endregion

                            #region 插入圖片與Technician名字

                            if (IsToPDF)
                            {
                                string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.inspector}'
and t.GarmentTest=1
";
                                string technicianName = string.Empty;
                               
                                Excel.Range cell = worksheet.Cells[12, 2];

                                DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection());

                                if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0 && dtTechnicianInfo.Rows[0]["SignaturePic"] != null)
                                {
                                    technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                    byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];
                                    string imageName = $"{Guid.NewGuid()}.jpg";
                                    string imgPath;

                                    if (test)
                                    {
                                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                                    }
                                    else
                                    {
                                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                                    }

                                    // Name
                                    worksheet.Cells[80, 9] = technicianName;
                                    Excel.Range cellNew = worksheet.Cells[78, 9];
                                    using (MemoryStream ms = new MemoryStream(imgData))
                                    {
                                        Image img = Image.FromStream(ms);
                                        img.Save(imgPath);
                                    }
                                    
                                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellNew.Left, cellNew.Top, 100, 24);
                                }
                                else
                                {
                                    worksheet.Cells[78, 9] = MyUtility.Convert.GetString(all_Data.Detail.GarmentTest_Detail_Inspector);
                                }
                            }

                            if (!IsToPDF)
                            {
                                worksheet.Cells[74, 8] = string.Empty;
                            }
                            #endregion

                            #region After Wash Appearance Check list
                            string tmpAR;

                            worksheet.Cells[65, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First());
                            worksheet.get_Range("65:65", Type.Missing).Rows.AutoFit();

                            // 大約21個字換行
                            int widhthBase = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First()).ToString().Length / 20;

                            worksheet.get_Range("65:65", Type.Missing).RowHeight = widhthBase == 0 ? 28 : 28 * widhthBase;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[65, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[65, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[65, 8] = tmpAR;
                            }

                            string strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 65, strComment);
                            worksheet.Cells[65, 10] = strComment;

                            worksheet.Cells[66, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Type).First());

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[66, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[66, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[66, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[66, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[66, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[66, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[66, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 66, strComment);
                            worksheet.Cells[66, 10] = strComment;

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash1).First());

                            worksheet.Cells[67, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Type).First()).ToString(); // type;

                            // 大約21個字換行
                            int widhthBase2 = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Type).First()).ToString().Length / 20;

                            worksheet.get_Range("67:67", Type.Missing).RowHeight = widhthBase2 == 0 ? 28 : 28 * widhthBase2;

                            if ((
                                    worksheet.get_Range("65:65", Type.Missing).RowHeight
                                    + worksheet.get_Range("66:66", Type.Missing).RowHeight
                                    + worksheet.get_Range("67:67", Type.Missing).RowHeight) < 81)
                            {
                                worksheet.get_Range("65:65", Type.Missing).RowHeight = worksheet.get_Range("65:65", Type.Missing).RowHeight > 28 ? worksheet.get_Range("65:65", Type.Missing).RowHeight : 28;
                                worksheet.get_Range("66:66", Type.Missing).RowHeight = 28;
                                worksheet.get_Range("67:67", Type.Missing).RowHeight = worksheet.get_Range("67:67", Type.Missing).RowHeight > 28 ? worksheet.get_Range("67:67", Type.Missing).RowHeight : 28;
                            }

                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[67, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[67, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[67, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[67, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[67, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 67, strComment);
                            worksheet.Cells[67, 10] = strComment;

                            worksheet.Cells[68, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First());
                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[68, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[68, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[68, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 68, strComment);
                            worksheet.Cells[68, 10] = strComment;

                            worksheet.Cells[69, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Type).First());

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[69, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[69, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[69, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 69, strComment);
                            worksheet.Cells[69, 10] = strComment;

                            worksheet.Cells[70, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Type).First());

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[70, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[70, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[70, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 70, strComment);
                            worksheet.Cells[70, 10] = strComment;

                            worksheet.Cells[71, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Type).First());

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[71, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[71, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[71, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 71, strComment);
                            worksheet.Cells[71, 10] = strComment;

                            worksheet.Cells[72, 3] = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Type).First());

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash1).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 4] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 5] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 4] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash2).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 6] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 7] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 6] = tmpAR;
                            }

                            tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash3).First());
                            if (tmpAR.EqualString("Accepted"))
                            {
                                worksheet.Cells[72, 8] = "V";
                            }
                            else if (tmpAR.EqualString("Rejected"))
                            {
                                worksheet.Cells[72, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[72, 8] = tmpAR;
                            }

                            strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Comment).First());
                            this.RowHeight(worksheet, 72, strComment);
                            worksheet.Cells[72, 10] = strComment;

                            #endregion

                            #region Spirality
                            if (all_Data.Spiralities.Where(x => x.Location.EqualString("B")).Any())
                            {
                                var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("B")).First();
                                worksheet.Cells[58, 6] = dr.MethodA;
                                worksheet.Cells[59, 6] = dr.MethodB;
                                worksheet.Cells[60, 6] = dr.CM;
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A56:A60", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (all_Data.Spiralities.Where(x => x.Location.EqualString("T")).Any())
                            {
                                var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("T")).First();
                                worksheet.Cells[53, 6] = dr.MethodA;
                                worksheet.Cells[54, 6] = dr.MethodB;
                                worksheet.Cells[55, 6] = dr.CM;
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A51:A55", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }
                            #endregion

                            // Streched Neck Opening is OK according to size spec?
                            if (data.ComboNeck.EqualString("Yes"))
                            {
                                worksheet.Cells[42, 9] = "V";
                            }
                            else
                            {
                                worksheet.Cells[42, 11] = "V";
                            }

                            #region Shrinkage
                            if (dtShrinkages.Select("Location = 'B'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'B'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A49:A49", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[46 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A44:A51", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'O'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'O'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A39:A39", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[36 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A34:A41", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'I'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'I'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A32:A32", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[28 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A26:A33", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }

                            if (dtShrinkages.Select("Location = 'T'").Length > 0)
                            {
                                DataTable dt = dtShrinkages.Select("Location = 'T'").CopyToDataTable();

                                // 超過5個測量點則新增行數
                                if (dt.Rows.Count > 5)
                                {
                                    for (int i = 0; i < dt.Rows.Count - 5; i++)
                                    {
                                        Excel.Range rng = worksheet.get_Range("A24:A24", Type.Missing).EntireRow;
                                        rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                    }
                                }

                                // 依不同品牌/套裝塞入資料
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    for (int c = 3; c < dt.Columns.Count; c++)
                                    {
                                        worksheet.Cells[20 + r, c] = this.AddShrinkageUnit_18(dt, r, c);
                                    }
                                }
                            }
                            else
                            {
                                Excel.Range rng = worksheet.get_Range("A18:A25", Type.Missing).EntireRow;
                                rng.Select();
                                rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                Marshal.ReleaseComObject(rng);
                            }
                            #endregion
                        }
                        #endregion

                        #region Save & Show Excel

                        string fileName_2018 = $"{basefileName_2018}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                        string filexlsx_2018 = fileName_2018 + ".xlsx";
                        string fileNamePDF_2018 = fileName_2018 + ".pdf";

                        string filepath_2018;
                        string filepathpdf_2018;
                        if (test)
                        {
                            filepath_2018 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx_2018);
                            filepathpdf_2018 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF_2018);
                        }
                        else
                        {
                            filepath_2018 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx_2018);
                            filepathpdf_2018 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF_2018);
                        }

                        string fileProcessName_2018 = "FGWT" + "_"
                            + all_Data.Main.SeasonID.ToString() + "_" + all_Data.Main.StyleID.ToString() + "_" + all_Data.Main.Article.ToString();
                        Excel.Workbook workbook_2018 = objApp.ActiveWorkbook;
                        workbook_2018.SaveAs(filepath_2018);
                        workbook_2018.Close();
                        objApp.Quit();
                        Marshal.ReleaseComObject(worksheet);
                        Marshal.ReleaseComObject(workbook_2018);
                        Marshal.ReleaseComObject(objApp);

                        if (IsToPDF)
                        {
                            if (ConvertToPDF.ExcelToPDF(filepath_2018, filepathpdf_2018))
                            {
                                all_Data.reportPath = fileNamePDF_2018;
                                all_Data.Result = true;
                            }
                            else
                            {
                                all_Data.ErrMsg = "Convert To PDF Fail";
                                all_Data.Result = false;
                            }
                        }

                        if (!IsToPDF)
                        {
                            all_Data.reportPath = filexlsx_2018;
                            all_Data.Result = true;
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        all_Data.ErrMsg = ex.Message.Replace("'", string.Empty);
                        all_Data.Result = false;
                    }

                    #endregion
                    break;
                case ReportType.Wash_Test_2020:
                    #region Print Wash 2020
                    if (all_Data.FGWT.Count == 0)
                    {
                        all_Data.Result = false;
                        all_Data.ErrMsg = "FGWT data not found!";
                        return all_Data;
                    }

                    string basefileName_2020 = "WashTest_2020_FGWT";
                    string openfilepath_2020;
                    if (test)
                    {
                        openfilepath_2020 = "C:\\Willy_Repository\\Quality_KPI\\Quality\\Quality\\bin\\XLT\\WashTest_2020_FGWT.xltx";
                    }
                    else
                    {
                        openfilepath_2020 = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName_2020}.xltx";
                    }

                    Excel.Application objApp_2020 = MyUtility.Excel.ConnectExcel(openfilepath_2020);
                    objApp_2020.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                    Excel.Worksheet worksheet_2020 = objApp_2020.ActiveWorkbook.Worksheets[1]; // 取得工作表

                    // objApp.Visible = true;
                    #region 插入圖片與Technician名字

                    if (IsToPDF)
                    {
                        string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.inspector}'
and t.GarmentTest=1
";
                        string technicianName = string.Empty;
                        DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection());

                        if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0 && dtTechnicianInfo.Rows[0]["SignaturePic"] != null)
                        {
                            technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                            byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];
                            string imageName = $"{Guid.NewGuid()}.jpg";
                            string imgPath;

                            if (test)
                            {
                                imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                            }
                            else
                            {
                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                            }

                            // Name
                            worksheet_2020.Cells[31, 7] = technicianName;
                            Excel.Range cellNew = worksheet_2020.Cells[29, 7];

                            using (MemoryStream ms = new MemoryStream(imgData))
                            {
                                Image img = Image.FromStream(ms);
                                img.Save(imgPath);
                            }

                            worksheet_2020.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellNew.Left, cellNew.Top, 100, 24);
                        }
                        else
                        {
                            worksheet_2020.Cells[31, 7] = MyUtility.Convert.GetString(all_Data.Detail.GarmentTest_Detail_Inspector);
                        }
                    }

                    if (!IsToPDF)
                    {
                        worksheet_2020.Cells[27, 6] = string.Empty;
                    }

                    #endregion

                    // 若為QA 10產生則顯示New Development Testing ( V )，若為QA P04產生則顯示1st Bulk Testing ( V )
                    worksheet_2020.Cells[4, 3] = "1st Bulk Testing ( V )";

                    worksheet_2020.Cells[5, 1] = "adidas Article No.: " + MyUtility.Convert.GetString(all_Data.Main.Article);
                    worksheet_2020.Cells[5, 3] = "adidas Working No.: " + MyUtility.Convert.GetString(all_Data.Main.StyleID);
                    worksheet_2020.Cells[5, 4] = "adidas Model No.: " + StyleName;

                    worksheet_2020.Cells[6, 1] = "T1 Supplier Ref.: " + orders.FactoryID;
                    worksheet_2020.Cells[6, 3] = "T1 Factory Name: " +  orders.BrandAreaCode;
                    worksheet_2020.Cells[6, 4] = "LO to Factory: " + data.TxtLotoFactory;

                    if (data.DateSubmit.HasValue)
                    {
                        worksheet_2020.Cells[8, 1] = "Date: " + data.DateSubmit.Value.ToString("yyyy/MM/dd");
                    }

                    int copyCount = all_Data.FGWT.Count - 2;

                    for (int i = 0; i <= copyCount - 1; i++)
                    {
                        // 複製儲存格
                        Excel.Range rgCopy = worksheet_2020.get_Range("A13:A13").EntireRow;

                        // 選擇要被貼上的位置
                        Excel.Range rgPaste = worksheet_2020.get_Range("A13:A13", Type.Missing);

                        // 貼上
                        rgPaste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, rgCopy.Copy(Type.Missing));
                    }

                    worksheet_2020.get_Range($"B12", $"B{all_Data.FGWT.Count + 11}").Merge(false);

                    int startRowIndex = 12;

                    // 開始填入表身
                    foreach (var dr in all_Data.FGWT)
                    {
                        // Requirement
                        worksheet_2020.Cells[startRowIndex, 3] = MyUtility.Convert.GetString(dr.Type);

                        // Test Results
                        // 若[GarmentTest_Detail_FGWT.Scale]非null則帶入Scale，若為null則帶入 [GarmentTest_Detail_FGWT.AfterWash - GarmentTest_Detail_FGWT.BeforeWash.]
                        if (!string.IsNullOrEmpty(dr.Scale))
                        {
                            worksheet_2020.Cells[startRowIndex, 4] = MyUtility.Convert.GetString(dr.Scale);
                        }
                        else
                        {
                            if ((dr.BeforeWash != 0 && dr.AfterWash != 0 && dr.Shrinkage != 0)
                                || MyUtility.Convert.GetBool(dr.IsInPercentage))
                            {
                                // TestDetail  % 或Range% 視作相同
                                if (MyUtility.Convert.GetString(dr.TestDetail).Contains("%"))
                                {
                                    worksheet_2020.Cells[startRowIndex, 4] = MyUtility.Convert.GetDouble(dr.Shrinkage);
                                }
                                else
                                {
                                    worksheet_2020.Cells[startRowIndex, 4] = MyUtility.Convert.GetDouble(dr.AfterWash) - MyUtility.Convert.GetDouble(dr.BeforeWash);
                                }
                            }
                        }

                        // Test Details
                        worksheet_2020.Cells[startRowIndex, 5] = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);

                        // adidas pass
                        worksheet_2020.Cells[startRowIndex, 6] = MyUtility.Convert.GetString(dr.Result);

                        startRowIndex++;
                    }

                    #region Save & Show Excel


                 
                    string fileName_2020 = $"{basefileName_2020}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                    string filexlsx_2020 = fileName_2020 + ".xlsx";
                    string fileNamePDF_2020 = fileName_2020 + ".pdf";

                    string filepath_2020;
                    string filepathpdf_2020;
                    if (test)
                    {
                        filepath_2020 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx_2020);
                        filepathpdf_2020 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF_2020);
                    }
                    else
                    {
                        filepath_2020 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx_2020);
                        filepathpdf_2020 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF_2020);
                    }

                    string fileProcessName_2020 = "FGWT" + "_"
                       + all_Data.Main.SeasonID.ToString() + "_" + all_Data.Main.StyleID.ToString() + "_" + all_Data.Main.Article.ToString();
                    Excel.Workbook workbook_2020 = objApp_2020.ActiveWorkbook;
                    workbook_2020.SaveAs(filepath_2020);
                    workbook_2020.Close();
                    objApp_2020.Quit();
                    Marshal.ReleaseComObject(worksheet_2020);
                    Marshal.ReleaseComObject(workbook_2020);
                    Marshal.ReleaseComObject(objApp_2020);

                    if (IsToPDF)
                    {
                        if (ConvertToPDF.ExcelToPDF(filepath_2020, filepathpdf_2020))
                        {
                            all_Data.reportPath = fileNamePDF_2020;
                            all_Data.Result = true;
                        }
                        else
                        {
                            all_Data.ErrMsg = "Convert To PDF Fail";
                            all_Data.Result = false;
                        }
                    }

                    if (!IsToPDF)
                    {
                        all_Data.reportPath = filexlsx_2020;
                        all_Data.Result = true;
                    }
                    #endregion

                    #endregion
                    break;
                case ReportType.Physical_Test:
                    #region Print Physical
                    if (all_Data.FGPT.Count == 0)
                    {
                        all_Data.ErrMsg = "FGPT data not found.";
                        all_Data.Result = false;
                        return all_Data;
                    }

                    string basefileName_Physical = "WashTest_Physical";
                    string openfilepath_Physical;
                    if (test)
                    {
                        openfilepath_Physical = "C:\\Willy_Repository\\Quality_KPI\\Quality\\Quality\\bin\\XLT\\WashTest_Physical.xltx";
                    }
                    else
                    {
                        openfilepath_Physical = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName_Physical}.xltx";
                    }

                    Excel.Application objApp_Physical = MyUtility.Excel.ConnectExcel(openfilepath_Physical);

                    objApp_Physical.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                    Excel.Worksheet worksheet_Physical = objApp_Physical.ActiveWorkbook.Worksheets[1]; // 取得工作表

                    // objApp.Visible = true;
                    #region 插入圖片與Technician名字
                    if (IsToPDF)
                    {
                        string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.inspector}'
and t.GarmentTest=1
";
                        string technicianName = string.Empty;
                        DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection());


                        if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0 && dtTechnicianInfo.Rows[0]["SignaturePic"] != null)
                        {
                            technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                            byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];


                            string imageName = $"{Guid.NewGuid()}.jpg";
                            string imgPath;

                            if (test)
                            {
                                imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                            }
                            else
                            {
                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                            }

                            // Name
                            worksheet_Physical.Cells[159, 7] = technicianName;
                            Excel.Range cellNew = worksheet_Physical.Cells[157, 7];

                            using (MemoryStream ms = new MemoryStream(imgData))
                            {
                                Image img = Image.FromStream(ms);
                                img.Save(imgPath);
                            }
                            worksheet_Physical.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellNew.Left, cellNew.Top, 80, 24);
                        }
                        else
                        {
                            worksheet_Physical.Cells[159, 7] = MyUtility.Convert.GetString(all_Data.Detail.GarmentTest_Detail_Inspector);
                        }
                    }

                    if (!IsToPDF)
                    {
                        worksheet_Physical.Cells[156, 6] = string.Empty;
                    }
                    #endregion

                    // 若為QA 10產生則顯示New Development Testing ( V )，若為QA P04產生則顯示1st Bulk Testing ( V )
                    worksheet_Physical.Cells[4, 3] = "1st Bulk Testing ( V )";

                    worksheet_Physical.Cells[5, 1] = "adidas Article No.: " + MyUtility.Convert.GetString(all_Data.Main.Article);
                    worksheet_Physical.Cells[5, 3] = "adidas Working No.: " + MyUtility.Convert.GetString(all_Data.Main.StyleID);
                    worksheet_Physical.Cells[5, 4] = "adidas Model No.: " + StyleName;

                    worksheet_Physical.Cells[6, 1] = "T1 Supplier Ref.: " + orders.FactoryID;
                    worksheet_Physical.Cells[6, 3] = "T1 Factory Name: " + orders.BrandAreaCode;
                    worksheet_Physical.Cells[6, 4] = "LO to Factory: " + data.TxtLotoFactory;

                    if (data.DateSubmit.HasValue)
                    {
                        worksheet_Physical.Cells[8, 1] = "Date: " + data.DateSubmit.Value.ToString("yyyy/MM/dd");
                    }

                    var testName_1 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName) == "PHX-AP0413");
                    var testName_2 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName) == "PHX-AP0450 Seam Breakage");
                    var testName_3 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName) == "PHX-AP0451");

                    #region 儲存格處理

                    // 因為PHX-AP0451在最下面，且只會有一筆，因此先複製這個，不然要重算Row index

                    // PHX-AP0451

                    // Requirement
                    worksheet_Physical.Cells[150, 3] = MyUtility.Convert.GetString(testName_3.FirstOrDefault().Type);

                    // Test Results
                    worksheet_Physical.Cells[150, 4] = MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestResult);

                    // Test Details
                    worksheet_Physical.Cells[150, 5] = MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestDetail);

                    // adidas pass
                    worksheet_Physical.Cells[150, 6] = MyUtility.Convert.GetString(testName_3.FirstOrDefault().Result);

                    // PHX-AP0450
                    int copyCount_2 = testName_2.Count() - 2;

                    for (int i = 0; i <= copyCount_2 - 1; i++)
                    {
                        // 複製儲存格
                        Excel.Range rgCopy = worksheet_Physical.get_Range("A149:A149").EntireRow;

                        // 選擇要被貼上的位置
                        Excel.Range rgPaste = worksheet_Physical.get_Range("A149:A149", Type.Missing);

                        // 貼上
                        rgPaste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, rgCopy.Copy(Type.Missing));
                    }

                    worksheet_Physical.get_Range($"B148", $"B{copyCount_2 + 149}").Merge(false);

                    // PHX - AP0413
                    int copyCount_1 = testName_1.Count() - 2;

                    for (int i = 0; i <= copyCount_1 - 1; i++)
                    {
                        // 複製儲存格
                        Excel.Range rgCopy = worksheet_Physical.get_Range("A135:A135").EntireRow;

                        // 選擇要被貼上的位置
                        Excel.Range rgPaste = worksheet_Physical.get_Range("A135:A135", Type.Missing);

                        // 貼上
                        rgPaste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, rgCopy.Copy(Type.Missing));
                    }

                    worksheet_Physical.get_Range($"B134", $"B{copyCount_1 + 135}").Merge(false);

                    #endregion

                    // 開始填入表身，先填PHX - AP0413
                    int startRowIndex_Pyhsical = 134;
                    foreach (var dr in testName_1)
                    {
                        // Requirement
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 3] = MyUtility.Convert.GetString(dr.Type);

                        // Test Results
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 4] = MyUtility.Convert.GetString(dr.TestResult);

                        // Test Details
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 5] = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);

                        // adidas pass
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 6] = MyUtility.Convert.GetString(dr.Result);

                        startRowIndex_Pyhsical++;
                    }

                    // 開始填入表身，填PHX - AP0450
                    startRowIndex_Pyhsical = testName_1.Count() + 133 + 12 + 1;
                    /*說明PHX - AP0413 這個Test Name最後的Index 為copyCount_1 + 133,與PHX-AP0450起點Index中間差了12 Row*/

                    foreach (var dr in testName_2)
                    {
                        // Requirement
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 3] = MyUtility.Convert.GetString(dr.Type);

                        // Test Results
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 4] = MyUtility.Convert.GetString(dr.TestResult);

                        // Test Details
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 5] = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);

                        // adidas pass
                        worksheet_Physical.Cells[startRowIndex_Pyhsical, 6] = MyUtility.Convert.GetString(dr.Result);

                        startRowIndex_Pyhsical++;
                    }

                    #region Save & Show Excel

                    string fileProcessName_Physical = "FGWT" + "_"
                      + all_Data.Main.SeasonID.ToString() + "_" + all_Data.Main.StyleID.ToString() + "_" + all_Data.Main.Article.ToString();

                    string fileName_Physical = $"WashTest_Physical_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                    string filexlsx_Physical = fileName_Physical + ".xlsx";
                    string fileNamePDF_Physical = fileName_Physical + ".pdf";

                    string filepath_Physical;
                    string filepathpdf_Physical;
                    if (test)
                    {
                        filepath_Physical = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx_Physical);
                        filepathpdf_Physical = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF_Physical);
                    }
                    else
                    {
                        filepath_Physical = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx_Physical);
                        filepathpdf_Physical = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF_Physical);
                    }


                    Excel.Workbook workbook_Physical = objApp_Physical.ActiveWorkbook;
                    workbook_Physical.SaveAs(filepath_Physical);
                    workbook_Physical.Close();
                    objApp_Physical.Quit();
                    Marshal.ReleaseComObject(worksheet_Physical);
                    Marshal.ReleaseComObject(workbook_Physical);
                    Marshal.ReleaseComObject(objApp_Physical);

                    if (IsToPDF)
                    {
                        if (ConvertToPDF.ExcelToPDF(filepath_Physical, filepathpdf_Physical))
                        {
                            all_Data.reportPath = fileNamePDF_Physical;
                            all_Data.Result = true;
                        }
                        else
                        {
                            all_Data.ErrMsg = "Convert To PDF Fail";
                            all_Data.Result = false;
                        }
                    }

                    if (!IsToPDF)
                    {
                        all_Data.reportPath = filexlsx_Physical;
                        all_Data.Result = true;
                    }
                    #endregion
                    #endregion
                    break;
                default:
                    break;
            }

            return all_Data;
        }
    }
}
