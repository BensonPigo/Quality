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
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System.IO;
using Library;
using Ict;
using System.Net.Mail;
using System.Web;
using ClosedXML.Excel;
using ADOHelper.Template.MSSQL;
using DocumentFormat.OpenXml.Wordprocessing;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Spreadsheet;

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
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;

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

                        if (!_IGarmentTestDetailProvider.Encode_GarmentTestDetail_OrderIDCheck(ID, No))
                        {
                            result.Result = false;
                            result.ErrMsg = "SP# cant not be empty.";
                        }
                        else
                        {
                            // 重新判斷Result
                            if (_IGarmentTestDetailProvider.Encode_GarmentTestDetail(ID, No, "Confirmed") == false ||
                                _IGarmentTestDetailProvider.Update_GarmentTestDetail_Result(ID, No) == false ||
                                _IGarmentTestProvider.Update_GarmentTest_Result(ID) == false)
                            {
                                result.Result = false;
                            }

                            // all result 有任一個是Fail 就寄信
                            result.sentMail = _IGarmentTestDetailProvider.Chk_AllResult(ID, No);
                        }

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
        public GarmentTest_ViewModel UpdateMailSender(string ID, string No, string UserID)
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

        public GarmentTest_Result SentMail(string ID, string No, List<Quality_MailGroup> mailGroups, string Subject, string Body, List<HttpPostedFileBase> Files)
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
                DataTable allResult = _IGarmentTestDetailProvider.Get_AllResult(ID, No);
                string aICommentType = string.Empty;
                if (allResult.Rows[0]["WashResult"].ToString() == "F")
                {
                    aICommentType += "Garment Wash Test";
                }
                if (allResult.Rows[0]["SeamBreakageResult"].ToString() == "F")
                {
                    aICommentType += ",Seam Breakage";
                }
                if (allResult.Rows[0]["OdourResult"].ToString() == "F")
                {
                    aICommentType += ",Odour Test";
                }

                string name = $"Garment Test_{dtContent.Rows[0]["OrderID"]}_" +
                    $"{dtContent.Rows[0]["StyleID"]}_" +
                    $"{dtContent.Rows[0]["Article"]}_" +
                    $"{dtContent.Rows[0]["Result"]}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                GarmentTest_Detail_Result baseResult = ToReport(ID, No, ReportType.Wash_Test_2018, true, AssignedFineName: name);
                string FileName = baseResult.Result.Value ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.reportPath) : string.Empty;
                SendMail_Request request = new SendMail_Request()
                {
                    To = ToAddress,
                    CC = CCAddress,
                    //Subject = "Garment Test – Test Fail",
                    Subject = $"Garment Test/{dtContent.Rows[0]["OrderID"]}/" +
                    $"{dtContent.Rows[0]["StyleID"]}/" +
                    $"{dtContent.Rows[0]["Article"]}/" +
                    $"{dtContent.Rows[0]["Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = strHtml,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = aICommentType,
                    StyleID = dtContent.Rows[0]["StyleID"].ToString(),
                    SeasonID = dtContent.Rows[0]["SeasonID"].ToString(),
                    BrandID = dtContent.Rows[0]["BrandID"].ToString(),
                };


                if (!string.IsNullOrEmpty(Subject))
                {
                    request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(request);
                string buyReadyDate = _MailService.GetBuyReadyDate(request);
                string mailBody = MailTools.DataTableChangeHtml(dtContent, comment, buyReadyDate, Body, out AlternateView plainView);

                request.Body = mailBody;
                request.alternateView = plainView;

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

                foreach (var detail in result.garmentTest_Details)
                {
                    DataTable dtContent = _IGarmentTestDetailProvider.Get_Mail_Content(result.garmentTest.ID.ToString(), detail.No.ToString());

                    string Subject = $"Garment Test/{dtContent.Rows[0]["OrderID"]}/" +
                    $"{dtContent.Rows[0]["StyleID"]}/" +
                    $"{dtContent.Rows[0]["Article"]}/" +
                    $"{dtContent.Rows[0]["Result"]}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    detail.MailSubject = Subject;
                }

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

        //private void RowHeight2(Excel.Worksheet worksheet, int row, string strComment)
        //{
        //    if (strComment.Length > 15)
        //    {
        //        decimal n = Math.Ceiling(strComment.Length / (decimal)15.0) * (decimal)12.25;
        //        worksheet.Range[$"A{row}", $"A{row}"].RowHeight = n;
        //    }
        //}
        private void RowHeight(IXLWorksheet worksheet, int rowNumber, string strComment)
        {
            if (strComment.Length > 15)
            {
                var row = worksheet.Row(rowNumber);

                double n = System.Math.Ceiling(strComment.Length / (double)15.0) * (double)12.25;
                row.Height = n;
            }
        }

        private string AddShrinkageUnit_18(System.Data.DataTable dt, int row, int columns)
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


        public GarmentTest_Detail_Result ToReport(string ID, string No, ReportType type, bool IsToPDF, bool test = false, string AssignedFineName = "")
        {
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            IStyleProvider styleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            IOrdersProvider ordersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);

            // 取得資料
            GarmentTest_Detail_Result all_Data = this.Get_All_Detail(ID, No);

            all_Data.reportPath = string.Empty;
            P04Data p04Data = new P04Data
            {
                DateSubmit = all_Data.Detail.SubmitDate,
                ReceiveDate = all_Data.Detail.ReceiveDate,
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

            var testCodeFgwt = _QualityBrandTestCodeProvider.Get(all_Data.Main.BrandID, "Garment Test-FGWT").ToList();

            bool isNewData = all_Data.Apperance.Count != 9;

            string styleName = styleProvider.GetStyleName(all_Data.Main.StyleID, all_Data.Main.SeasonID, all_Data.Main.BrandID);
            string criticalName = styleProvider.GetStyleCritical(all_Data.Main.StyleID, all_Data.Main.SeasonID, all_Data.Main.BrandID);

            var orderDetails = this.GetOrderDetails(all_Data);

            var dtShrinkages = _IGarmentTestDetailShrinkageProvider.Get_dt_Shrinkage(ID, No);

            string typeName = this.GenerateTypeName(all_Data);

            char[] invalidChars = this.GetInvalidFileNameChars();

            Dictionary<string, string> result = GenerateReport(type, all_Data, p04Data, isNewData, styleName, criticalName, orderDetails, dtShrinkages, testCodeFgwt, invalidChars, IsToPDF, AssignedFineName);

            all_Data.Result = Convert.ToBoolean(result["Result"]);
            all_Data.reportPath = result["reportPath"];
            all_Data.reportFileFullPath = result["reportFileFullPath"];
            all_Data.ErrMsg = result["ErrMsg"];

            return all_Data;
        }

        public GarmentTest_Detail_Result DownloadAllReport(string ID, string No, bool IsToPDF)
        {
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            GarmentTest_Detail_Result report1 = this.ToReport(ID, No, ReportType.Wash_Test_2018, IsToPDF);
            GarmentTest_Detail_Result report2 = this.ToReport(ID, No, ReportType.Wash_Test_2020, IsToPDF);
            GarmentTest_Detail_Result report3 = this.ToReport(ID, No, ReportType.Physical_Test, IsToPDF);

            Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile();
            string zipName = $"Garment Test Wash All Report_{DateTime.Now.ToString("yyyyMMdd")}.zip";
            string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);
            if (report1.Result.Value)
            {
                zipFile.AddFile(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report1.reportPath), string.Empty);
            }
            if (report2.Result.Value)
            {
                zipFile.AddFile(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report2.reportPath), string.Empty);
            }
            if (report3.Result.Value)
            {
                zipFile.AddFile(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report3.reportPath), string.Empty);
            }
            zipFile.Save(zipPath);
            result.reportPath = zipName;
            result.Result = true;
            return result;
        }

        private Orders GetOrderDetails(GarmentTest_Detail_Result all_Data)
        {
            IOrdersProvider ordersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            if (!string.IsNullOrEmpty(all_Data.Detail.OrderID))
            {
                var query = ordersProvider.Get(new Orders() { ID = all_Data.Detail.OrderID });
                if (query.Any())
                {
                    return query.FirstOrDefault();
                }
            }
            return new Orders();
        }
        private string GenerateTypeName(GarmentTest_Detail_Result all_Data)
        {
            return $"Garment Test_{all_Data.Detail.OrderID}_{all_Data.Main.StyleID}_{all_Data.Main.Article}_{(all_Data.Detail.Result == "P" ? "Pass" : "Fail")}_{DateTime.Now:yyyyMMddHHmmss}";
        }
        private char[] GetInvalidFileNameChars()
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            char[] additionalChars = { '-', '+' };
            return invalidChars.Concat(additionalChars).ToArray();
        }

        private Dictionary<string, string> GenerateReport(ReportType type, GarmentTest_Detail_Result all_Data, P04Data p04data, bool isNewData, string styleName, string criticalName, Orders orders, DataTable dtShrinkages, List<QualityBrandTestCode> testCodeFgwt, char[] invalidChars, bool isToPDF, string AssignedFineName = "")
        {
            Dictionary<string, string> dicRt = new Dictionary<string, string>
            {
                ["reportPath"] = string.Empty,
                ["reportFileFullPath"] = string.Empty,
                ["Result"] = string.Empty,
                ["ErrMsg"] = string.Empty,
            };

            try
            {
                string factory = orders.FactoryID.Empty() ? System.Web.HttpContext.Current.Session["FactoryID"].ToString() :orders.FactoryID;

                string FactoryNameEN = _IGarmentTestProvider.GetFactoryNameEN(factory);
                string typeName = $"Garment Test_{all_Data.Detail.OrderID}_" +
                    $"{all_Data.Main.StyleID}_" +
                    $"{all_Data.Main.Article}_" +
                    $"{(all_Data.Detail.Result == "P" ? "Pass" : "Fail")}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                switch (type)
                {
                    case ReportType.Wash_Test_2018:
                        #region Print Wash 2018

                        if (all_Data.Apperance.Count == 0 || all_Data.Shrinkages.Count == 0)
                        {
                            dicRt["ErrMsg"] = "Data not found!";
                            dicRt["Result"] = "false";
                            return dicRt;
                        }

                        typeName = $"Garment Test Wash 2018_{all_Data.Detail.OrderID}_" +
                           $"{all_Data.Main.StyleID}_" +
                           $"{all_Data.Main.Article}_" +
                           $"{(all_Data.Detail.Result == "P" ? "Pass" : "Fail")}_" +
                           $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                        try
                        {
                            if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                            {
                                System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                            }

                            if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                            {
                                System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                            }

                            string openfilepath;
                            string basefileName_2018 = "WashTest_2018";
                            openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName_2018}.xltx";

                            using (var workbook = new XLWorkbook(openfilepath))
                            {
                                var worksheet = workbook.Worksheet(1);

                                #region WashName 條件處理
                                if (all_Data.Main.WashName == "710")
                                {
                                    worksheet.Cell(18, 8).Value = "10.wash";
                                    worksheet.Cell(18, 10).Value = "15.wash";
                                    worksheet.Cell(26, 8).Value = "10.wash";
                                    worksheet.Cell(26, 10).Value = "15.wash";
                                    worksheet.Cell(34, 8).Value = "10.wash";
                                    worksheet.Cell(34, 10).Value = "15.wash";
                                    worksheet.Cell(44, 8).Value = "10.wash";
                                    worksheet.Cell(44, 10).Value = "15.wash";

                                    worksheet.Cell(63, 6).Value = "10.wash";
                                    worksheet.Cell(63, 8).Value = "15.wash";

                                    worksheet.Cell(15, 4).Value = "V";
                                    worksheet.Cell(15, 8).Value = "15";
                                }
                                else if (all_Data.Main.WashName == "701")
                                {
                                    worksheet.Cell(18, 10).Value = "";
                                    worksheet.Cell(26, 10).Value = "";
                                    worksheet.Cell(34, 10).Value = "";
                                    worksheet.Cell(44, 10).Value = "";

                                    worksheet.Cell(63, 8).Value = "";

                                    worksheet.Cell(15, 4).Value = "";
                                    worksheet.Cell(15, 8).Value = "3";
                                }
                                #endregion

                                if (testCodeFgwt.Any())
                                {
                                    worksheet.Cell(1, 2).Value = $@"Finished Garment Wash Test Report({testCodeFgwt.FirstOrDefault().TestCode})";
                                }
                                worksheet.Cell(3, 4).Value = MyUtility.Convert.GetString(all_Data.Detail.ReportNo);

                                if (all_Data.Detail.SubmitDate.HasValue)
                                    worksheet.Cell(3, 7).Value = all_Data.Detail.SubmitDate.Value.ToString("yyyy/MM/dd");
                                if (all_Data.Detail.ReceiveDate.HasValue)
                                    worksheet.Cell(3, 10).Value = all_Data.Detail.ReceiveDate.Value.ToString("yyyy/MM/dd");
                                if (all_Data.Detail.inspdate.HasValue)
                                    worksheet.Cell(4, 4).Value = all_Data.Detail.inspdate.Value.ToString("yyyy/MM/dd");


                                // Report Date
                                if (!MyUtility.Check.Empty(all_Data.Detail.SubmitDate))
                                {
                                    worksheet.Cell(3, 7).Value = MyUtility.Convert.GetDate(all_Data.Detail.SubmitDate).Value.Year + "/" + MyUtility.Convert.GetDate(all_Data.Detail.SubmitDate).Value.Month + "/" + MyUtility.Convert.GetDate(all_Data.Detail.SubmitDate).Value.Day;
                                }

                                // Receive Date
                                if (!MyUtility.Check.Empty(all_Data.Detail.ReceiveDate))
                                {
                                    worksheet.Cell(3, 10).Value = MyUtility.Convert.GetDate(all_Data.Detail.ReceiveDate).Value.Year + "/" + MyUtility.Convert.GetDate(all_Data.Detail.ReceiveDate).Value.Month + "/" + MyUtility.Convert.GetDate(all_Data.Detail.ReceiveDate).Value.Day;
                                }

                                // Submit Date
                                if (!MyUtility.Check.Empty(all_Data.Detail.inspdate))
                                {
                                    worksheet.Cell(4, 4).Value = MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Year + "/" + MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Month + "/" + MyUtility.Convert.GetDate(all_Data.Detail.inspdate).Value.Day;
                                }


                                worksheet.Cell(4, 7).Value = MyUtility.Convert.GetString(orders.ID);
                                worksheet.Cell(4, 9).Value = MyUtility.Convert.GetString(all_Data.Main.BrandID);
                                worksheet.Cell(4, 11).Value = MyUtility.Convert.GetString(all_Data.Main.SeasonID);

                                worksheet.Cell(6, 4).Value = MyUtility.Convert.GetString(all_Data.Main.StyleID);
                                worksheet.Cell(7, 8).Value = orders.CustPONO;
                                worksheet.Cell(7, 4).Value = MyUtility.Convert.GetString(all_Data.Main.Article);
                                worksheet.Cell(6, 8).Value = styleName;
                                worksheet.Cell(6, 11).Value = string.IsNullOrEmpty(criticalName) ? string.Empty : "Y";
                                worksheet.Cell(8, 8).Value = MyUtility.Convert.GetDecimal(p04data.NumArriveQty);

                                string sendDate = Convert.ToDateTime(orders.BuyerDelivery).ToShortDateString();
                                worksheet.Cell(8, 4).Value = sendDate;
                                worksheet.Cell(8, 10).Value = MyUtility.Convert.GetString(p04data.TxtSize);

                                worksheet.Cell(11, 4).Value = p04data.RdbtnLine == true ? "V" : string.Empty;
                                worksheet.Cell(12, 4).Value = p04data.RdbtnTumble == true ? "V" : string.Empty;
                                worksheet.Cell(13, 4).Value = p04data.RdbtnHand == true ? "V" : string.Empty;
                                worksheet.Cell(11, 8).Value = p04data.ComboTemperature + "˚C ";
                                worksheet.Cell(12, 8).Value = p04data.ComboMachineModel;
                                worksheet.Cell(13, 8).Value = p04data.TxtFibreComposition;
                                // Remark
                                worksheet.Cell(14, 4).Value = p04data.Remark;

                                string remark = p04data.Remark == null ? string.Empty : p04data.Remark;
                                int rowCtn = (remark.Length / 85) + 1;

                                worksheet.Cell(14, 14).WorksheetRow().Height = 15.75 * rowCtn;

                                if (all_Data.Main.IsSkirt)
                                {
                                    worksheet.Cell(44, 3).Value = "SKIRT";
                                }

                                #region 舊資料
                                if (!isNewData)
                                {
                                    #region 照片
                                    if (all_Data.Detail.TestBeforePicture != null)
                                    {
                                        this.AddImageToWorksheet(worksheet, all_Data.Detail.TestBeforePicture, 94, 2, 328, 247);
                                    }

                                    if (all_Data.Detail.TestAfterPicture != null)
                                    {
                                        this.AddImageToWorksheet(worksheet, all_Data.Detail.TestAfterPicture, 94, 7, 328, 247);
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
                                        worksheet.Cell(87, 4).Value = "V";
                                    }
                                    else if (MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("F") || ApperanceRejected)
                                    {
                                        worksheet.Cell(87, 6).Value = "V";
                                    }

                                    #endregion

                                    #region 插入圖片與Technician名字

                                    // ToPDF
                                    if (!string.IsNullOrEmpty(all_Data.Detail.inspector))
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

                                        if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0)
                                        {
                                            technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                            if (dtTechnicianInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                            {

                                                byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];

                                                string imageName = $"{Guid.NewGuid()}.jpg";
                                                string imgPath;

                                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                                                // Name
                                                worksheet.Cell(86, 8).Value = technicianName;
                                                this.AddImageToWorksheet(worksheet, imgData, 88, 8, 100, 24);
                                            }
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(all_Data.Detail.Approver))
                                    {
                                        string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.Approver}'
and t.GarmentTest=1
";

                                        string approverName = string.Empty;
                                        DataTable dtApproverInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                        if (dtApproverInfo != null && dtApproverInfo.Rows.Count > 0)
                                        {
                                            approverName = dtApproverInfo.Rows[0]["name"].ToString();
                                            if (dtApproverInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                            {

                                                byte[] imgData = (byte[])dtApproverInfo.Rows[0]["SignaturePic"];

                                                string imageName = $"{Guid.NewGuid()}.jpg";
                                                string imgPath;

                                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                                                // Name
                                                worksheet.Cell(86, 10).Value = approverName;
                                                this.AddImageToWorksheet(worksheet, imgData, 88, 10, 100, 24);
                                            }
                                        }
                                    }

                                    #endregion

                                    #region After Wash Appearance Check list
                                    string tmpAR;

                                    worksheet.Cell(75, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First());

                                    //worksheet.get_Range("75:75", Type.Missing).Rows.AutoFit();

                                    // 大約21個字換行
                                    int widhthBase = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First()).Length / 20;

                                    //worksheet.get_Range("75:75", Type.Missing).RowHeight = widhthBase == 0 ? 28 : 28 * widhthBase;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 8).Value = tmpAR;
                                    }

                                    string strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 75, strComment);
                                    worksheet.Cell(75, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(76, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(76, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(76, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(76, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 76, strComment);
                                    worksheet.Cell(76, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(77, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(77, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 77, strComment);
                                    worksheet.Cell(77, 10).Value = strComment;

                                    worksheet.Cell(78, 3).Value = all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First().ToString(); // type;

                                    // 大約21個字換行
                                    int widhthBase2 = all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First().ToString().Length / 20;

                                    worksheet.Row(78).Height = widhthBase2 == 0 ? 28 : 28 * widhthBase2;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 78, strComment);
                                    worksheet.Cell(78, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 79, strComment);
                                    worksheet.Cell(79, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 80, strComment);
                                    worksheet.Cell(80, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 81, strComment);
                                    worksheet.Cell(81, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 82, strComment);
                                    worksheet.Cell(82, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(83, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(83, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(83, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(83, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(83, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(83, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(83, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(83, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(83, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(9)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 83, strComment);
                                    worksheet.Cell(83, 10).Value = strComment;
                                    #endregion

                                    #region Spirality

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("O")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("O")).First();
                                        worksheet.Cell(68, 6).Value = dr.MethodA;
                                        worksheet.Cell(69, 6).Value = dr.MethodB;
                                        worksheet.Cell(70, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(66, 70).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A66:A70", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("I")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("I")).First();
                                        worksheet.Cell(63, 6).Value = dr.MethodA;
                                        worksheet.Cell(64, 6).Value = dr.MethodB;
                                        worksheet.Cell(65, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(61, 65).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A61:A65", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("B")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("B")).First();
                                        worksheet.Cell(58, 6).Value = dr.MethodA;
                                        worksheet.Cell(59, 6).Value = dr.MethodB;
                                        worksheet.Cell(60, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(56, 60).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A56:A60", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("T")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("T")).First();
                                        worksheet.Cell(53, 6).Value = dr.MethodA;
                                        worksheet.Cell(54, 6).Value = dr.MethodB;
                                        worksheet.Cell(55, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(51, 55).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A51:A55", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }
                                    #endregion

                                    // Streched Neck Opening is OK according to size spec?
                                    if (p04data.ComboNeck.EqualString("Yes"))
                                    {
                                        worksheet.Cell(42, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(42, 11).Value = "V";
                                    }

                                    #region Shrinkage

                                    if (dtShrinkages.Select("Location = 'B'").Length > 0)
                                    {
                                        System.Data.DataTable dt_B = dtShrinkages.Select("Location = 'B'").CopyToDataTable();



                                        // 超過5個測量點則新增行數
                                        if (dt_B.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt_B.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(50);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(51).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(51);
                                                //Excel.Range rng = worksheet.get_Range("A50:A50", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt_B.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt_B.Columns.Count; c++)
                                            {
                                                worksheet.Cell(46 + r, c).Value = this.AddShrinkageUnit_18(dt_B, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(44, 51).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A44:A51", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'O'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'O'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(40);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(41).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(41);
                                                //Excel.Range rng = worksheet.get_Range("A40:A40", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(35 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(34, 41).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A34:A41", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'I'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'I'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製
                                                var rowToCopy = worksheet.Row(32);

                                                // 2. 插入一列
                                                worksheet.Row(33).InsertRowsAbove(1);

                                                // 3. 複製格式
                                                var newRow = worksheet.Row(32);
                                                //Excel.Range rng = worksheet.get_Range("A32:A32", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(28 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(26, 33).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A26:A33", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'T'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'T'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製
                                                var rowToCopy = worksheet.Row(24);

                                                // 2. 插入一列
                                                worksheet.Row(25).InsertRowsAbove(1);

                                                // 3. 複製格式
                                                var newRow = worksheet.Row(25);
                                                //Excel.Range rng = worksheet.get_Range("A24:A24", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(20 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(18, 25).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A18:A25", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    #endregion

                                }
                                #endregion

                                #region 新資料
                                if (isNewData)
                                {
                                    #region 照片
                                    if (all_Data.Detail.TestBeforePicture != null)
                                    {
                                        this.AddImageToWorksheet(worksheet, all_Data.Detail.TestBeforePicture, 94, 2, 328, 247);
                                    }

                                    if (all_Data.Detail.TestAfterPicture != null)
                                    {
                                        this.AddImageToWorksheet(worksheet, all_Data.Detail.TestAfterPicture, 94, 7, 328, 247);
                                    }
                                    #endregion

                                    //worksheet.get_Range("76:76", Type.Missing).Delete();
                                    worksheet.Row(76).Delete();

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
                                        worksheet.Cell(86, 4).Value = "V";
                                    }
                                    else if (MyUtility.Convert.GetString(all_Data.Detail.WashResult).EqualString("F") || ApperanceRejected)
                                    {
                                        worksheet.Cell(86, 6).Value = "V";
                                    }
                                    #endregion

                                    #region 插入圖片與Technician名字

                                    if (!string.IsNullOrEmpty(all_Data.Detail.inspector))
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

                                        System.Data.DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                        if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0)
                                        {
                                            technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                            // Name
                                            worksheet.Cell(85, 8).Value = technicianName;

                                            if (dtTechnicianInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                            {
                                                byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];
                                                string imageName = $"{Guid.NewGuid()}.jpg";
                                                string imgPath;

                                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                                                this.AddImageToWorksheet(worksheet, imgData, 87, 8, 100, 24);
                                            }
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(all_Data.Detail.Approver))
                                    {
                                        string sql_cmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
outer apply (select PicPath from Production.dbo.system) s 
where t.ID = '{all_Data.Detail.Approver}'
and t.GarmentTest=1
";
                                        string technicianName = string.Empty;

                                        System.Data.DataTable dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sql_cmd, new ADOHelper.Template.MSSQL.SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                        if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0)
                                        {
                                            technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                            // Name
                                            worksheet.Cell(85, 10).Value = technicianName;

                                            if (dtTechnicianInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                            {
                                                byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];
                                                string imageName = $"{Guid.NewGuid()}.jpg";
                                                string imgPath;

                                                imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                                                this.AddImageToWorksheet(worksheet, imgData, 87, 10, 100, 24);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region After Wash Appearance Check list
                                    string tmpAR;

                                    worksheet.Cell(75, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First());
                                    //worksheet.get_Range("75:75", Type.Missing).Rows.AutoFit();

                                    // 大約21個字換行
                                    int widhthBase = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Type).First()).ToString().Length / 20;

                                    worksheet.Row(75).Height = widhthBase == 0 ? 28 : 28 * widhthBase;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(75, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(75, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(75, 8).Value = tmpAR;
                                    }

                                    string strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(1)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 75, strComment);
                                    worksheet.Cell(75, 10).Value = strComment;

                                    worksheet.Cell(76, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Type).First());

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(76, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(76, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(76, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(76, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(76, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(76, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(76, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(2)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 76, strComment);
                                    worksheet.Cell(76, 10).Value = strComment;

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash1).First());

                                    worksheet.Cell(77, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Type).First()).ToString(); // type;

                                    // 大約21個字換行
                                    int widhthBase2 = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Type).First()).ToString().Length / 20;

                                    //worksheet.get_Range("77:77", Type.Missing).RowHeight = widhthBase2 == 0 ? 28 : 28 * widhthBase2;
                                    worksheet.Row(77).Height = widhthBase2 == 0 ? 28 : 28 * widhthBase2;

                                    if ((
                                            worksheet.Row(77).Height
                                            + worksheet.Row(77).Height
                                            + worksheet.Row(77).Height) < 81)
                                    {
                                        worksheet.Row(77).Height = worksheet.Row(77).Height > 28 ? worksheet.Row(77).Height : 28;
                                        worksheet.Row(77).Height = 28;
                                        worksheet.Row(77).Height = worksheet.Row(77).Height > 28 ? worksheet.Row(77).Height : 28;
                                    }

                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(77, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(77, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(77, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(77, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(77, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(3)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 77, strComment);
                                    worksheet.Cell(77, 10).Value = strComment;

                                    worksheet.Cell(78, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Type).First());
                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(78, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(78, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(78, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(4)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 78, strComment);
                                    worksheet.Cell(78, 10).Value = strComment;

                                    worksheet.Cell(79, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Type).First());

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(79, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(79, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(79, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(5)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 79, strComment);
                                    worksheet.Cell(79, 10).Value = strComment;

                                    worksheet.Cell(80, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Type).First());

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(80, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(80, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(80, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(6)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 80, strComment);
                                    worksheet.Cell(80, 10).Value = strComment;

                                    worksheet.Cell(81, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Type).First());

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(81, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(81, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(81, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(7)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 81, strComment);
                                    worksheet.Cell(81, 10).Value = strComment;

                                    worksheet.Cell(82, 3).Value = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Type).First());

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash1).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 4).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 5).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 4).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash2).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 6).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 7).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 6).Value = tmpAR;
                                    }

                                    tmpAR = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Wash3).First());
                                    if (tmpAR.EqualString("Accepted"))
                                    {
                                        worksheet.Cell(82, 8).Value = "V";
                                    }
                                    else if (tmpAR.EqualString("Rejected"))
                                    {
                                        worksheet.Cell(82, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(82, 8).Value = tmpAR;
                                    }

                                    strComment = MyUtility.Convert.GetString(all_Data.Apperance.Where(x => x.Seq.Equals(8)).Select(x => x.Comment).First());
                                    this.RowHeight(worksheet, 82, strComment);
                                    worksheet.Cell(82, 10).Value = strComment;

                                    #endregion

                                    #region Spirality
                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("O")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("O")).First();
                                        worksheet.Cell(68, 6).Value = dr.MethodA;
                                        worksheet.Cell(69, 6).Value = dr.MethodB;
                                        worksheet.Cell(70, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(66, 70).Delete();

                                        //Excel.Range rng = worksheet.get_Range("A66:A70", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("I")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("I")).First();
                                        worksheet.Cell(63, 6).Value = dr.MethodA;
                                        worksheet.Cell(64, 6).Value = dr.MethodB;
                                        worksheet.Cell(65, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(61, 65).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A61:A65", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("B")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("B")).First();
                                        worksheet.Cell(58, 6).Value = dr.MethodA;
                                        worksheet.Cell(59, 6).Value = dr.MethodB;
                                        worksheet.Cell(60, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(56, 60).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A56:A60", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (all_Data.Spiralities.Where(x => x.Location.EqualString("T")).Any())
                                    {
                                        var dr = all_Data.Spiralities.Where(x => x.Location.EqualString("T")).First();
                                        worksheet.Cell(53, 6).Value = dr.MethodA;
                                        worksheet.Cell(54, 6).Value = dr.MethodB;
                                        worksheet.Cell(55, 6).Value = dr.CM;
                                    }
                                    else
                                    {
                                        worksheet.Rows(51, 55).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A51:A55", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }
                                    #endregion

                                    // Streched Neck Opening is OK according to size spec?
                                    if (p04data.ComboNeck.EqualString("Yes"))
                                    {
                                        worksheet.Cell(42, 9).Value = "V";
                                    }
                                    else
                                    {
                                        worksheet.Cell(42, 11).Value = "V";
                                    }

                                    #region Shrinkage
                                    if (dtShrinkages.Select("Location = 'B'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'B'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(49);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(50).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(50);
                                                //Excel.Range rng = worksheet.get_Range("A49:A49", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(46 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(44, 51).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A44:A51", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'O'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'O'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(39);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(40).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(40);

                                                rowToCopy.CopyTo(newRow);
                                                //Excel.Range rng = worksheet.get_Range("A39:A39", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(36 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(34, 41).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A34:A41", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'I'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'I'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(32);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(33).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(33);

                                                rowToCopy.CopyTo(newRow);
                                                //Excel.Range rng = worksheet.get_Range("A32:A32", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(28 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(26, 33).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A26:A33", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    if (dtShrinkages.Select("Location = 'T'").Length > 0)
                                    {
                                        System.Data.DataTable dt = dtShrinkages.Select("Location = 'T'").CopyToDataTable();

                                        // 超過5個測量點則新增行數
                                        if (dt.Rows.Count > 5)
                                        {
                                            for (int i = 0; i < dt.Rows.Count - 5; i++)
                                            {
                                                // 1. 複製第 10 列
                                                var rowToCopy = worksheet.Row(24);

                                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                                worksheet.Row(25).InsertRowsAbove(1);

                                                // 3. 複製格式到新插入的列
                                                var newRow = worksheet.Row(25);

                                                rowToCopy.CopyTo(newRow);
                                                //Excel.Range rng = worksheet.get_Range("A24:A24", Type.Missing).EntireRow;
                                                //rng.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                            }
                                        }

                                        // 依不同品牌/套裝塞入資料
                                        for (int r = 0; r < dt.Rows.Count; r++)
                                        {
                                            for (int c = 3; c < dt.Columns.Count; c++)
                                            {
                                                worksheet.Cell(20 + r, c).Value = this.AddShrinkageUnit_18(dt, r, c);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Rows(18, 25).Delete();
                                        //Excel.Range rng = worksheet.get_Range("A18:A25", Type.Missing).EntireRow;
                                        //rng.Select();
                                        //rng.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                                        //Marshal.ReleaseComObject(rng);
                                    }

                                    #endregion
                                }
                                #endregion

                                #region Title
                                // 1. 複製第 1列
                                //var rowToCopy1 = worksheet.Row(2);

                                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                worksheet.Row(1).InsertRowsAbove(1);

                                // 3. 合併欄位
                                worksheet.Range("B1:K1").Merge();
                                // 設置字體樣式
                                var mergedCell = worksheet.Cell("B1");
                                mergedCell.Value = FactoryNameEN;
                                mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                                mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                                mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                mergedCell.Style.Font.Bold = true;
                                #endregion
                                #region Save & Show Excel

                                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                                {
                                    typeName = AssignedFineName;
                                }
                                // 去除非法字元
                                typeName = FileNameHelper.SanitizeFileName(typeName);
                                string filexlsx_2018 = typeName + ".xlsx";
                                string fileNamePDF_2018 = typeName + ".pdf";

                                string filepath_2018;
                                string filepathpdf_2018;
                                filepath_2018 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx_2018);
                                filepathpdf_2018 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF_2018);

                                string fileProcessName_2018 = "FGWT" + "_"
                                    + all_Data.Main.SeasonID.ToString() + "_" + all_Data.Main.StyleID.ToString() + "_" + all_Data.Main.Article.ToString();
                                //Excel.Workbook workbook_2018 = objApp.ActiveWorkbook;
                                workbook.SaveAs(filepath_2018);
                                dicRt["reportPath"] = filexlsx_2018;
                                dicRt["reportFileFullPath"] = filepath_2018;

                                //if (isToPDF)
                                //{
                                //    //LibreOfficeService s = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                                //    //s.ConvertExcelToPdf(filepath_2018, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                                //    ConvertToPDF.ExcelToPDF(filepath_2018, filepathpdf_2018);
                                //    dicRt["reportPath"] = fileNamePDF_2018;
                                //    dicRt["reportFileFullPath"] = filepathpdf_2018;
                                //}

                                dicRt["Result"] = "true";

                                #endregion
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }

                        #endregion
                        break;
                    case ReportType.Wash_Test_2020:
                        #region Print Wash 2020
                        if (all_Data.FGWT.Count == 0)
                        {
                            dicRt["ErrMsg"] = "FGWT data not found!";
                            dicRt["Result"] = "false";
                            return dicRt;
                        }

                        typeName = $"Garment Test Wash 2020_{all_Data.Detail.OrderID}_" +
                                   $"{all_Data.Main.StyleID}_" +
                                   $"{all_Data.Main.Article}_" +
                                   $"{(all_Data.Detail.Result == "P" ? "Pass" : "Fail")}_" +
                                   $"{DateTime.Now:yyyyMMddHHmmss}";

                        string basefileName_2020 = "WashTest_2020_FGWT";
                        string templatePath_2020 = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "XLT", $"{basefileName_2020}.xltx");

                        if (!System.IO.File.Exists(templatePath_2020))
                        {
                            dicRt["ErrMsg"] = "Template not found!";
                            dicRt["Result"] = "false";
                            return dicRt;
                        }

                        try
                        {
                            using (var workbook = new XLWorkbook(templatePath_2020))
                            {
                                var worksheet_2020 = workbook.Worksheet(1);

                                #region 插入圖片與Technician名字

                                if (!string.IsNullOrEmpty(all_Data.Detail.inspector))
                                {
                                    string sqlCmd = $@"
select p.name, [SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
where t.ID = '{all_Data.Detail.inspector}' and t.GarmentTest = 1";

                                    var dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sqlCmd, new SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                    if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0)
                                    {
                                        string technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                        // 填寫名字
                                        worksheet_2020.Cell(28, 5).Value = technicianName;

                                        if (dtTechnicianInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                        {
                                            byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];

                                            // 插入圖片
                                            AddImageToWorksheet(worksheet_2020, imgData, 30, 5, 100, 24);
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(all_Data.Detail.Approver))
                                {
                                    string sqlCmd = $@"
select p.name, [SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
where t.ID = '{all_Data.Detail.Approver}' and t.GarmentTest = 1";

                                    var dtApproverInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sqlCmd, new SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                    if (dtApproverInfo != null && dtApproverInfo.Rows.Count > 0)
                                    {
                                        string approverName = dtApproverInfo.Rows[0]["name"].ToString();
                                        // 填寫名字
                                        worksheet_2020.Cell(28, 7).Value = approverName;

                                        if (dtApproverInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                        {
                                            byte[] imgData = (byte[])dtApproverInfo.Rows[0]["SignaturePic"];

                                            // 插入圖片
                                            AddImageToWorksheet(worksheet_2020, imgData, 30, 7, 100, 24);
                                        }
                                    }
                                }

                                #endregion

                                // 填寫固定資料
                                worksheet_2020.Cell(1, 1).Value = testCodeFgwt.Any()
                                    ? $"Product TEST REPORT({testCodeFgwt.FirstOrDefault().TestCode})"
                                    : "Product TEST REPORT";

                                worksheet_2020.Cell(4, 3).Value = "1st Bulk Testing ( V )";
                                worksheet_2020.Cell(5, 1).Value = "adidas Article No.: " + MyUtility.Convert.GetString(all_Data.Main.Article);
                                worksheet_2020.Cell(5, 3).Value = "adidas Working No.: " + MyUtility.Convert.GetString(all_Data.Main.StyleID);
                                worksheet_2020.Cell(5, 4).Value = "adidas Model No.: " + styleName;

                                worksheet_2020.Cell(6, 1).Value = "T1 Supplier Ref.: " + orders.FactoryID;
                                worksheet_2020.Cell(6, 3).Value = "T1 Factory Name: " + orders.BrandAreaCode;
                                worksheet_2020.Cell(6, 4).Value = "LO to Factory: " + p04data.TxtLotoFactory;

                                if (p04data.DateSubmit.HasValue)
                                {
                                    worksheet_2020.Cell(8, 1).Value = "Report Date: " + p04data.DateSubmit.Value.ToString("yyyy/MM/dd");
                                }
                                if (p04data.ReceiveDate.HasValue)
                                {
                                    worksheet_2020.Cell(9, 1).Value = "Receive Date: " + p04data.ReceiveDate.Value.ToString("yyyy/MM/dd");
                                }

                                // 複製範圍與填寫資料
                                int copyCount = all_Data.FGWT.Count - 2;
                                for (int i = 0; i < copyCount; i++)
                                {
                                    //worksheet_2020.Row(13).InsertRowsAbove(1);

                                    // 1. 複製第 10 列
                                    var rowToCopy = worksheet_2020.Row(13);

                                    // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                                    worksheet_2020.Row(14).InsertRowsAbove(1);

                                    // 3. 複製格式到新插入的列
                                    var newRow = worksheet_2020.Row(14);

                                    rowToCopy.CopyTo(newRow);
                                }

                                worksheet_2020.Range($"B12:B{all_Data.FGWT.Count + 11}").Merge();

                                int startRowIndex = 12;
                                foreach (var dr in all_Data.FGWT)
                                {
                                    worksheet_2020.Cell(startRowIndex, 3).Value = MyUtility.Convert.GetString(dr.Type);

                                    if (!string.IsNullOrEmpty(dr.Scale))
                                    {
                                        worksheet_2020.Cell(startRowIndex, 4).Value = MyUtility.Convert.GetString(dr.Scale);
                                    }
                                    else if ((dr.BeforeWash != 0 && dr.AfterWash != 0 && dr.Shrinkage != 0) || MyUtility.Convert.GetBool(dr.IsInPercentage))
                                    {
                                        if (MyUtility.Convert.GetString(dr.TestDetail).Contains("%"))
                                        {
                                            worksheet_2020.Cell(startRowIndex, 4).Value = MyUtility.Convert.GetDouble(dr.Shrinkage);
                                        }
                                        else
                                        {
                                            worksheet_2020.Cell(startRowIndex, 4).Value = MyUtility.Convert.GetDouble(dr.AfterWash) - MyUtility.Convert.GetDouble(dr.BeforeWash);
                                        }
                                    }

                                    worksheet_2020.Cell(startRowIndex, 5).Value = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);
                                    worksheet_2020.Cell(startRowIndex, 6).Value = MyUtility.Convert.GetString(dr.StandardRemark);
                                    worksheet_2020.Cell(startRowIndex, 7).Value = MyUtility.Convert.GetString(dr.Result);

                                    worksheet_2020.Row(startRowIndex).Height = CalculateRowHeight(worksheet_2020.Row(startRowIndex).Cells());
                                    startRowIndex++;
                                }

                                // Excel 合併 + 塞資料
                                #region Title
                                // 1. 插入一列
                                worksheet_2020.Row(1).InsertRowsAbove(1);

                                // 2. 合併欄位
                                worksheet_2020.Range("A1:I1").Merge();
                                // 設置字體樣式
                                var mergedCell = worksheet_2020.Cell("A1");
                                mergedCell.Value = FactoryNameEN;
                                mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                                mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                                mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                mergedCell.Style.Font.Bold = true;

                                //// 自動檢測使用範圍
                                var usedRange = worksheet_2020.RangeUsed();
                                var lastRow = worksheet_2020.CellsUsed().Max(cell => cell.Address.RowNumber);
                                //// 確認範圍不為空
                                if (usedRange != null)
                                {
                                    // 清除所有已有的列印範圍
                                    worksheet_2020.PageSetup.PrintAreas.Clear();

                                    // 設定列印範圍為使用範圍
                                    worksheet_2020.PageSetup.PrintAreas.Add($"A1:I{lastRow + 3}");
                                }
                                #endregion

                                // 儲存檔案
                                // 去除非法字元
                                typeName = FileNameHelper.SanitizeFileName(typeName);
                                string excelPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{typeName}.xlsx");
                                string pdfPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{typeName}.pdf");

                                workbook.SaveAs(excelPath);
                                dicRt["reportPath"] = $"{typeName}.xlsx";
                                dicRt["reportFileFullPath"] = excelPath;

                                if (isToPDF)
                                {
                                    //LibreOfficeService s = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                                    //s.ConvertExcelToPdf(excelPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                                    //ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                                    //dicRt["reportPath"] = $"{sanitizedTypeName}.pdf";
                                    //dicRt["reportFileFullPath"] = pdfPath;
                                }

                                dicRt["Result"] = "true";
                            }
                        }
                        catch (Exception ex)
                        {
                            dicRt["ErrMsg"] = ex.Message;
                            dicRt["Result"] = "false";
                        }

                        #endregion
                        break;
                    case ReportType.Physical_Test:
                        #region Print Physical
                        if (all_Data.FGPT.Count == 0)
                        {
                            dicRt["ErrMsg"] = "FGPT data not found.";
                            dicRt["Result"] = "false";
                            return dicRt;
                        }

                        typeName = $"Garment Test Physical_{all_Data.Detail.OrderID}_" +
                                   $"{all_Data.Main.StyleID}_" +
                                   $"{all_Data.Main.Article}_" +
                                   $"{(all_Data.Detail.Result == "P" ? "Pass" : "Fail")}_" +
                                   $"{DateTime.Now:yyyyMMddHHmmss}";

                        string basefileName_Physical = "WashTest_Physical";
                        string openfilepath_Physical = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "XLT", $"{basefileName_Physical}.xltx");

                        if (!File.Exists(openfilepath_Physical))
                        {
                            dicRt["ErrMsg"] = "Template not found.";
                            dicRt["Result"] = "false";
                            return dicRt;
                        }

                        try
                        {
                            using (var workbook = new XLWorkbook(openfilepath_Physical))
                            {
                                var worksheet_Physical = workbook.Worksheet(1);

                                #region 插入圖片與Technician名字
                                if (!string.IsNullOrEmpty(all_Data.Detail.inspector))
                                {
                                    string sqlCmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
where t.ID = '{all_Data.Detail.inspector}' and t.GarmentTest=1";
                                    var dtTechnicianInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sqlCmd, new SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                    if (dtTechnicianInfo != null && dtTechnicianInfo.Rows.Count > 0)
                                    {
                                        string technicianName = dtTechnicianInfo.Rows[0]["name"].ToString();
                                        if (dtTechnicianInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                        {
                                            byte[] imgData = (byte[])dtTechnicianInfo.Rows[0]["SignaturePic"];

                                            // 填寫名字
                                            worksheet_Physical.Cell(117, 5).Value = technicianName;

                                            // 使用共用函數插入圖片
                                            AddImageToWorksheet(worksheet_Physical, imgData, 119, 5, 100, 24);
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(all_Data.Detail.Approver))
                                {
                                    string sqlCmd = $@"
select p.name,[SignaturePic] = t.Signature
from Production.dbo.Technician t WITH (NOLOCK)
inner join Production.dbo.pass1 p WITH (NOLOCK) on t.ID = p.ID  
where t.ID = '{all_Data.Detail.Approver}' and t.GarmentTest=1";
                                    var dtApproverInfo = ADOHelper.Template.MSSQL.SQLDAL.ExecuteDataTable(CommandType.Text, sqlCmd, new SQLParameterCollection(), Common.ProductionDataAccessLayer);

                                    if (dtApproverInfo != null && dtApproverInfo.Rows.Count > 0)
                                    {
                                        string approverName = dtApproverInfo.Rows[0]["name"].ToString();
                                        if (dtApproverInfo.Rows[0]["SignaturePic"] != DBNull.Value)
                                        {
                                            byte[] imgData = (byte[])dtApproverInfo.Rows[0]["SignaturePic"];

                                            // 填寫名字
                                            worksheet_Physical.Cell(117, 7).Value = approverName;

                                            // 使用共用函數插入圖片
                                            AddImageToWorksheet(worksheet_Physical, imgData, 119, 7, 100, 24);
                                        }
                                    }
                                }
                                #endregion

                                // 若為QA 10產生則顯示New Development Testing ( V )，若為QA P04產生則顯示1st Bulk Testing ( V )
                                worksheet_Physical.Cell(4, 3).Value = "1st Bulk Testing ( V )";
                                worksheet_Physical.Cell(5, 1).Value = "adidas Article No.: " + MyUtility.Convert.GetString(all_Data.Main.Article);
                                worksheet_Physical.Cell(5, 3).Value = "adidas Working No.: " + MyUtility.Convert.GetString(all_Data.Main.StyleID);
                                worksheet_Physical.Cell(5, 4).Value = "adidas Model No.: " + styleName;

                                worksheet_Physical.Cell(6, 1).Value = "T1 Supplier Ref.: " + orders.FactoryID;
                                worksheet_Physical.Cell(6, 3).Value = "T1 Factory Name: " + orders.BrandAreaCode;
                                worksheet_Physical.Cell(6, 4).Value = "LO to Factory: " + p04data.TxtLotoFactory;

                                if (p04data.DateSubmit.HasValue)
                                {
                                    worksheet_Physical.Cell(8, 1).Value = "Report Date: " + p04data.DateSubmit.Value.ToString("yyyy/MM/dd");
                                }
                                if (p04data.ReceiveDate.HasValue)
                                {
                                    worksheet_Physical.Cell(9, 1).Value = "Receive Date: " + p04data.ReceiveDate.Value.ToString("yyyy/MM/dd");
                                }

                                var testName_1 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName).Contains("PHX-AP0413"));
                                var testName_2 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName).Contains("PHX-AP0450"));
                                var testName_3 = all_Data.FGPT.AsEnumerable().Where(o => MyUtility.Convert.GetString(o.TestName).Contains("PHX-AP0451"));

                                // PHX-AP0451

                                // Requirement
                                worksheet_Physical.Cell(110, 3).Value = MyUtility.Convert.GetString(testName_3.FirstOrDefault().Type);

                                // Test Results
                                worksheet_Physical.Cell(110, 4).Value = MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestResult);

                                // Test Details
                                worksheet_Physical.Cell(110, 5).Value = MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(testName_3.FirstOrDefault().TestDetail);

                                // adidas pass
                                worksheet_Physical.Cell(110, 8).Value = MyUtility.Convert.GetString(testName_3.FirstOrDefault().Result);

                                // PHX-AP0450 複製空格
                                int copyCount_2 = testName_2.Count() - 2;

                                // 第三筆開始才需要插入新的Row
                                if (copyCount_2 > 0)
                                {
                                    for (int i = 0; i <= copyCount_2 - 1; i++)
                                    {
                                        // 1. 複製
                                        var rowToCopy = worksheet_Physical.Row(109);

                                        // 2. 插入
                                        worksheet_Physical.Row(110).InsertRowsAbove(1);

                                        // 3. 複製格式
                                        var newRow = worksheet_Physical.Row(110);
                                        rowToCopy.CopyTo(newRow);
                                    }
                                    worksheet_Physical.Range($"B108", $"B{copyCount_2 + 108}").Merge(false);
                                }

                                // 開始填入表身，填PHX - AP0450
                                int startRowIndex_Pyhsical = 108;

                                foreach (var dr in testName_2)
                                {
                                    // Requirement
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 3).Value = MyUtility.Convert.GetString(dr.Type);

                                    // Test Results
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 4).Value = MyUtility.Convert.GetString(dr.TestResult);

                                    // Test Details
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 5).Value = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);

                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 6).Value = MyUtility.Convert.GetString(dr.StandardRemark);

                                    // adidas pass
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 7).Value = MyUtility.Convert.GetString(dr.Result);

                                    worksheet_Physical.Row(startRowIndex_Pyhsical).Height = CalculateRowHeight(worksheet_Physical.Row(startRowIndex_Pyhsical).Cells());

                                    startRowIndex_Pyhsical++;
                                }

                                // PHX - AP0413 複製空格
                                int copyCount_1 = testName_1.Count() - 2;

                                // 第三筆開始才需要插入新的Row
                                if (copyCount_1 > 0)
                                {
                                    for (int i = 0; i <= copyCount_1 - 1; i++)
                                    {
                                        // 1. 複製
                                        var rowToCopy = worksheet_Physical.Row(97);

                                        // 2. 插入
                                        worksheet_Physical.Row(98).InsertRowsAbove(1);

                                        // 3. 複製格式
                                        var newRow = worksheet_Physical.Row(98);
                                        rowToCopy.CopyTo(newRow);
                                    }

                                    worksheet_Physical.Range($"B96", $"B{copyCount_1 + 96}").Merge(false);
                                }

                                // 開始填入表身，先填PHX - AP0413
                                startRowIndex_Pyhsical = 96;
                                foreach (var dr in testName_1)
                                {
                                    // Requirement
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 3).Value = MyUtility.Convert.GetString(dr.Type);

                                    // Test Results
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 4).Value = MyUtility.Convert.GetString(dr.TestResult);

                                    // Test Details
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 5).Value = MyUtility.Convert.GetString(dr.TestDetail) == "Range%" ? "%" : MyUtility.Convert.GetString(dr.TestDetail);

                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 6).Value = MyUtility.Convert.GetString(dr.StandardRemark);

                                    // adidas pass
                                    worksheet_Physical.Cell(startRowIndex_Pyhsical, 7).Value = MyUtility.Convert.GetString(dr.Result);

                                    worksheet_Physical.Row(startRowIndex_Pyhsical).Height = CalculateRowHeight(worksheet_Physical.Row(startRowIndex_Pyhsical).Cells());
                                    startRowIndex_Pyhsical++;
                                }

                                // Excel 合併 + 塞資料
                                #region Title
                                // 1. 複製第 1列
                                var rowToCopy1 = worksheet_Physical.Row(2);

                                // 2. 插入一列
                                worksheet_Physical.Row(1).InsertRowsAbove(1);

                                // 3. 合併欄位
                                worksheet_Physical.Range("A1:I1").Merge();
                                // 設置字體樣式
                                var mergedCell = worksheet_Physical.Cell("A1");
                                mergedCell.Value = FactoryNameEN;
                                mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                                mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                                mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                mergedCell.Style.Font.Bold = true;

                                //// 自動檢測使用範圍
                                var usedRange = worksheet_Physical.RangeUsed();
                                var lastRow = worksheet_Physical.CellsUsed().Max(cell => cell.Address.RowNumber);
                                //// 確認範圍不為空
                                if (usedRange != null)
                                {
                                    // 清除所有已有的列印範圍
                                    worksheet_Physical.PageSetup.PrintAreas.Clear();

                                    // 設定列印範圍為使用範圍
                                    worksheet_Physical.PageSetup.PrintAreas.Add($"A1:I{lastRow + 3}");
                                }
                                #endregion

                                // 儲存檔案

                                // 去除非法字元
                                typeName = FileNameHelper.SanitizeFileName(typeName);
                                string excelPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{typeName}.xlsx");
                                string pdfPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{typeName}.pdf");

                                workbook.SaveAs(excelPath);
                                dicRt["reportPath"] = $"{typeName}.xlsx";
                                dicRt["reportFileFullPath"] = excelPath;

                                //if (isToPDF)
                                //{
                                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                                //    //officeService.ConvertExcelToPdf(excelPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                                //    ConvertToPDF.ExcelToPDF(excelPath, pdfPath);

                                //    dicRt["reportPath"] = $"{sanitizedTypeName}.pdf";
                                //    dicRt["reportFileFullPath"] = pdfPath;
                                //}

                                dicRt["Result"] = "true";
                            }
                        }
                        catch (Exception ex)
                        {
                            dicRt["ErrMsg"] = ex.Message;
                            dicRt["Result"] = "false";
                        }

                        #endregion
                        break;
                    default:
                        break;
                }

                dicRt["Result"] = "true";
            }
            catch (Exception ex)
            {
                dicRt["ErrMsg"] = ex.Message;
                dicRt["Result"] = "false";
            }
            return dicRt;
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
        public static double CalculateRowHeight(IXLCells cells)
        {
            double baseHeight = 15; // 預設列高
            double maxHeight = baseHeight;

            foreach (var cell in cells)
            {
                if (cell.IsMerged()) continue; // 跳過合併儲存格

                string cellText = cell.GetValue<string>();
                if (string.IsNullOrEmpty(cellText)) continue;

                // 1. 取得儲存格內文字的行數
                int lineCount = cellText.Split(new[] { '\n' }, StringSplitOptions.None).Length;

                // 2. 計算文字總長度
                int textLength = cellText.Length;

                // 3. 根據列寬計算適合的行數
                double columnWidth = cell.WorksheetColumn().Width;
                double estimatedLines = Math.Ceiling(textLength / columnWidth);

                // 4. 取最大行數，考慮多行文字
                double requiredLines = Math.Max(lineCount, estimatedLines);

                // 5. 根據行數計算高度 (假設每行高度為 15 單位)
                double cellHeight = requiredLines * baseHeight;

                // 更新該行的最大高度
                maxHeight = Math.Max(maxHeight, cellHeight);
            }

            return maxHeight; // 返回最終列高
        }

        public GarmentTest_Detail_Result StyleResult_BulkFGTReport(string BrandID, string StyleID, string SeasonID, string Article, string Type)
        {
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            GarmentTest_Detail_Result result = new GarmentTest_Detail_Result();
            try
            {
                GarmentTest_Request garment = new GarmentTest_Request
                {
                    Brand = BrandID,
                    Style = StyleID,
                    Season = SeasonID,
                    Article = Article,
                };

                GarmentTest_Detail_ViewModel detail_ViewModel = _IGarmentTestProvider.GetDetail_LastTestNo(garment, Type).FirstOrDefault();
                switch (Type)
                {
                    case "450":
                    case "451":
                        result = ToReport(detail_ViewModel.ID.ToString(), detail_ViewModel.No.ToString(), ReportType.Physical_Test, true);
                        break;
                    default:
                        Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile();
                        string zipName = $"WashTest_Physical_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.zip";
                        string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);
                        result = ToReport(detail_ViewModel.ID.ToString(), detail_ViewModel.No.ToString(), ReportType.Wash_Test_2018, true);
                        if (result.Result.Value)
                        {
                            zipFile.AddFile(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.reportPath), string.Empty);
                        }

                        result = ToReport(detail_ViewModel.ID.ToString(), detail_ViewModel.No.ToString(), ReportType.Wash_Test_2020, true);
                        if (result.Result.Value)
                        {
                            zipFile.AddFile(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.reportPath), string.Empty);
                        }

                        zipFile.Save(zipPath);
                        result.reportPath = zipName;
                        break;
                }

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrMsg = ex.Message;
            }

            return result;
        }
    }
}
