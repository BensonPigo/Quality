using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{
    public class InspectionService : IInspectionService
    {
        private IInspectionProvider _InspectionProvider;
        private IRFTInspectionProvider _RFTInspectionProvider;
        private IRFTInspectionDetailProvider _RFTInspectionDetailProvider;
        private IGarmentDefectTypeProvider _GarmentDefectTypeProvider;
        private IGarmentDefectCodeProvider _GarmentDefectCodeProvider;
        private IMailToProvider _IMailToProvider;
        private IAreaProvider _AreaProvider;
        private IDropDownListProvider _DropDownListProvider;
        private IReworkCardProvider _IReworkCardProvider;
        private IReworkListProvider _IReworkListProvider;
        private IRFTOrderCommentsProvider _IRFTOrderCommentsProvider;
        private IRFTPicDuringDummyFittingProvider _IRFTPicDuringDummyFittingProvider;
        private IRFTInspectionMeasurementProvider _IRFTInspectionMeasurementProvider;
        private IDQSReasonProvider _IDQSReasonProvider;

        // Production
        private IOrdersProvider _IOrdersProvider;
        private IStyleProvider _IStyleProvider;

        public enum SelectType
        {
            OrderID,
            StyleID,
            Article,
            Size,
            ProductType,
        }

        public enum DefectType
        {
            Responsibility,
            BAAuditCriteria,
        }

        public enum ReworkListType
        {
            Pass,
            Wash, 
            Repl,
            Print,
            Shade,
            Dispose,
        }

        public IList<Inspection_ViewModel> GetSelectItemData(Inspection_ViewModel inspection_ViewModel)
        {
            _InspectionProvider = new InspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Inspection_ViewModel> result = _InspectionProvider.GetSelectItemData(inspection_ViewModel).ToList();
            return result;
        }

        public IList<Inspection_ViewModel> CheckSelectItemData(Inspection_ViewModel inspection_ViewModel, SelectType type)
        {
            _InspectionProvider = new InspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Inspection_ViewModel> result = _InspectionProvider.CheckSelectItemData(inspection_ViewModel).ToList();

            if (result.Count() > 1)
            {
                switch (type)
                {
                    case SelectType.StyleID:
                        result = result.GroupBy(x => new { x.StyleID, x.OrderID, x.ProductTypePMS, x.Brand, x.Season, x.SampleStage, x.OriginalLine, x.OrderQty, x.SizeBalanceQty })
                                    .Select(x => new Inspection_ViewModel()
                                    {
                                        StyleID = x.Key.StyleID,
                                        OrderID = x.Key.OrderID,
                                        ProductTypePMS = x.Key.ProductTypePMS,
                                        Brand = x.Key.Brand,
                                        Season = x.Key.Season,
                                        SampleStage = x.Key.SampleStage,
                                        OriginalLine = x.Key.OriginalLine,
                                        OrderQty = x.Key.OrderQty,
                                        SizeBalanceQty = x.Key.SizeBalanceQty
                                    }).Where(r => Convert.ToInt32(r.SizeBalanceQty) > 0).ToList();
                        break;
                    case SelectType.OrderID:
                        result = result.GroupBy(x => new { x.StyleID, x.OrderID, x.ProductTypePMS, x.Brand, x.Season, x.SampleStage, x.OriginalLine, x.OrderQty, x.SizeBalanceQty })
                                    .Select(x => new Inspection_ViewModel()
                                    {
                                        StyleID = x.Key.StyleID,
                                        OrderID = x.Key.OrderID,
                                        ProductTypePMS = x.Key.ProductTypePMS,
                                        Brand = x.Key.Brand,
                                        Season = x.Key.Season,
                                        SampleStage = x.Key.SampleStage,
                                        OriginalLine = x.Key.OriginalLine,
                                        OrderQty = x.Key.OrderQty,
                                        SizeBalanceQty = x.Key.SizeBalanceQty
                                    }).Where(r => Convert.ToInt32(r.SizeBalanceQty) > 0).ToList();
                        break;
                    case SelectType.Article:
                        result = result.GroupBy(x => new { x.Article })
                                    .Select(x => new Inspection_ViewModel()
                                    {
                                        Article = x.Key.Article,
                                    }).ToList();
                        break;
                    case SelectType.Size:
                        result = result.GroupBy(x => new { x.Size, x.SizeQty, x.OrderQty })
                                    .Select(x => new Inspection_ViewModel()
                                    {
                                        Size = x.Key.Size,
                                        SizeQty = x.Key.SizeQty,
                                        OrderQty = x.Key.OrderQty
                                    }).ToList();
                        break;
                    case SelectType.ProductType:
                        result = result.GroupBy(x => new { x.ProductType, x.SizeBalanceQty, x.OrderBalanceQty })
                                    .Select(x => new Inspection_ViewModel()
                                    {
                                        ProductType = x.Key.ProductType,
                                        SizeBalanceQty = x.Key.SizeBalanceQty,
                                        OrderBalanceQty = x.Key.OrderBalanceQty
                                    }).ToList();
                        break;
                }
            }

            return result;
        }

        public IList<Inspection_ChkOrderID_ViewModel> CheckSelectItemData_SP(Inspection_ViewModel inspection_ViewModel)
        {
            _InspectionProvider = new InspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Inspection_ChkOrderID_ViewModel> result = _InspectionProvider.CheckSelectItemData_SP(inspection_ViewModel).ToList();
            return result;
        }

        public Inspection_ViewModel GetTop3(Inspection_ViewModel inspection_ViewModel)
        {
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);

            Inspection_ViewModel result = new Inspection_ViewModel();
            List<RFT_Inspection> inspections = _RFTInspectionProvider.Get(new RFT_Inspection()
                    {
                        FactoryID = inspection_ViewModel.FactoryID,
                        Line = inspection_ViewModel.Line,
                        InspectionDate = inspection_ViewModel.InspectionDate,
                    })
                    .ToList();

            List<RFT_Inspection> inspections_six = _RFTInspectionProvider.Get(new RFT_Inspection()
                    {
                        FactoryID = inspection_ViewModel.FactoryID,
                        Line = inspection_ViewModel.Line,
                    })
                    .ToList();

            List<RFT_Inspection_Detail> inspection_details = _RFTInspectionDetailProvider.Top3Defects(new RFT_Inspection()
                    {
                        FactoryID = inspection_ViewModel.FactoryID,
                        Line = inspection_ViewModel.Line,
                        InspectionDate = inspection_ViewModel.InspectionDate,
                    })
                    .ToList();

            result.Pass = inspections.Where(x => x.Status.Equals("Pass") || x.Status.Equals("Fixed")).Count().ToString();
            result.Reject = inspections.Where(x => x.Status.Equals("Reject") || x.Status.Equals("Fixed")).Count().ToString();
            result.Hard = inspections_six.Where(x => x.FixType.Equals("Hard") && x.Status.Equals("Reject")).Count().ToString();
            result.Quick = inspections_six.Where(x => x.FixType.Equals("Quick") && x.Status.Equals("Reject")).Count().ToString();
            result.Wash = inspections_six.Where(x => x.FixType.Equals("Wash") && x.Status.Equals("Reject")).Count().ToString();
            result.Repl = inspections_six.Where(x => x.FixType.Equals("Repl.") && x.Status.Equals("Reject")).Count().ToString();
            result.Print = inspections_six.Where(x => x.FixType.Equals("Print") && x.Status.Equals("Reject")).Count().ToString();
            result.Shade = inspections_six.Where(x => x.FixType.Equals("Shade") && x.Status.Equals("Reject")).Count().ToString();
            result.Dispose = inspections.Where(x => x.Status.Equals("Dispose")).Count().ToString();

            List<string> top3DefectCode = inspection_details
                .Select(x => x.DefectCode)
                .ToList();

            List<string> top3Area = inspection_details
                .Select(x => x.AreaCode)
                .ToList();

            List<Top3> top3 = new List<Top3>();

            for (int i = 1; i <= 3; i++)
            {
                Top3 top = new Top3()
                {
                    Defect = top3DefectCode.Count >= i ? top3DefectCode[i - 1] : string.Empty,
                    Area = top3Area.Count >= i ? top3Area[i - 1] : string.Empty,
                };

                top3.Add(top);
            }

            result.top3 = top3;

            return result;
        }

        public IList<GarmentDefectType> GetGarmentDefectType()
        {
            _GarmentDefectTypeProvider = new GarmentDefectTypeProvider(Common.ProductionDataAccessLayer);
            List<GarmentDefectType> defectTypes = _GarmentDefectTypeProvider.Get().ToList();
            return defectTypes;
        }

        public IList<GarmentDefectCode> GetGarmentDefectCode(GarmentDefectCode defectCode)
        {
            _GarmentDefectCodeProvider = new GarmentDefectCodeProvider(Common.ProductionDataAccessLayer);
            List<GarmentDefectCode> defectCodes = _GarmentDefectCodeProvider.Get(defectCode).ToList();
            return defectCodes;
        }

        public IList<Area> GetArea(Area area)
        {
            _AreaProvider = new AreaProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Area> areas = _AreaProvider.Get(area).ToList();
            return areas;
        }

        public IList<DropDownList> GetDropDownList(DropDownList downList)
        {
            _DropDownListProvider = new DropDownListProvider(Common.ProductionDataAccessLayer);
            List<DropDownList> dropDowns = _DropDownListProvider.Get(downList).ToList();
            return dropDowns;
        }

        public IList<MailTo> GetMailTo(MailTo item)
        {
            _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<MailTo> mailTos = _IMailToProvider.Get(
               new MailTo()
               {
                   ID = item.ID,
               }).ToList();

            return mailTos;
        }

        public InspectionSave_ViewModel SaveRFTInspection(InspectionSave_ViewModel inspections)
        {
            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                if (inspections.rft_Inspection == null)
                {
                    inspections.Result = false;
                    inspections.ErrMsg = "InspectionSave_ViewModel is null";
                    return inspections;
                }

                string OrderID = inspections.rft_Inspection.OrderID;

                _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
                IList<Orders> orders = _IOrdersProvider.Get(new Orders() { ID = OrderID });
                if (orders.Count > 0)
                {
                    string StyleUkey = orders[0].StyleUkey.ToString();
                    inspections.rft_Inspection.StyleUkey = long.Parse(StyleUkey);
                }

                #region update/insert [RFT_Inspection] and [RFT_Inspection_Detail]

                _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(_ISQLDataTransaction);

                #region 判斷資料是否是空值

                // Master
                string emptyMsg = string.Empty;
                if (string.IsNullOrEmpty(inspections.rft_Inspection.OrderID)){emptyMsg += "Master OrderID cannot be empty." + Environment.NewLine;}
                if (string.IsNullOrEmpty(inspections.rft_Inspection.Article)) { emptyMsg += "Master Article cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(inspections.rft_Inspection.Location)) { emptyMsg += "Master Location cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(inspections.rft_Inspection.Size)) { emptyMsg += "Master Size cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(inspections.rft_Inspection.Line)) { emptyMsg += "Master Line cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(inspections.rft_Inspection.FactoryID)) { emptyMsg += "Master FactoryID cannot be empty." + Environment.NewLine; }
                if (inspections.rft_Inspection.StyleUkey == 0 || inspections.rft_Inspection.StyleUkey == 0) { emptyMsg += "Master StyleUkey cannot be empty." + Environment.NewLine; }

                if (inspections.rft_Inspection.Status.ToLower() != "pass")
                {
                    if (string.IsNullOrEmpty(inspections.rft_Inspection.FixType)) { emptyMsg += "Master FixType cannot be empty." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(inspections.rft_Inspection.ReworkCardNo)) { emptyMsg += "Master ReworkCardNo cannot be empty." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(inspections.rft_Inspection.ReworkCardType)) { emptyMsg += "Master ReworkCardType cannot be empty." + Environment.NewLine; }
                }

                if (string.IsNullOrEmpty(inspections.rft_Inspection.Status)) { emptyMsg += "Master Status cannot be empty." + Environment.NewLine; }
                if (string.IsNullOrEmpty(inspections.rft_Inspection.AddName)) { emptyMsg += "Master AddName cannot be empty." + Environment.NewLine; }
           
                if (inspections.rft_Inspection.InspectionDate == null) { emptyMsg += "Master InspectionDate cannot be empty." + Environment.NewLine; }

                // Detail
                foreach (var item in inspections.fT_Inspection_Details)
                {
                    if (string.IsNullOrEmpty(item.DefectCode)) { emptyMsg += "Detail DefectCode cannot be empty." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(item.AreaCode)) { emptyMsg += "Detail AreaCode cannot be empty." + Environment.NewLine; }

                    if (string.IsNullOrEmpty(item.PMS_RFTRespID)) { emptyMsg += "Detail PMS_RFTRespID cannot be empty." + Environment.NewLine; }

                    if (string.IsNullOrEmpty(item.GarmentDefectTypeID)) { emptyMsg += "Detail GarmentDefectTypeID cannot be empty." + Environment.NewLine; }
                    if (string.IsNullOrEmpty(item.GarmentDefectCodeID)) { emptyMsg += "Detail GarmentDefectCodeID cannot be empty." + Environment.NewLine; }
                }

                if (!string.IsNullOrEmpty(emptyMsg))
                {
                    inspections.Result = false;
                    inspections.ErrMsg = emptyMsg;
                    return inspections;
                }

                #endregion

                // 判斷檢驗數量不可超過Order_Qty
                DataTable dtchk = _RFTInspectionDetailProvider.ChkInspQty(inspections.rft_Inspection);
                if (dtchk == null || dtchk.Rows.Count == 0)
                {
                    inspections.Result = false;
                    inspections.ErrMsg = "Inspection Qty cannot over than Order Qty.";
                    return inspections;
                }

                int createCnt = _RFTInspectionDetailProvider.Create_Master_Detail(inspections.rft_Inspection, inspections.fT_Inspection_Details);
                _ISQLDataTransaction.Commit();
                // Reject才寄信
                if (createCnt > 0)
                {
                    if (inspections.fT_Inspection_Details != null & inspections.fT_Inspection_Details.Count > 0)
                    {
                        _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);
                        List<MailTo> mailToSubject = _IMailToProvider.Get(
                           new MailTo()
                           {
                               ID = "200",
                           }).ToList();

                        // 取得 MR,SMR mail address
                        List<MailTo> mailToAddress = _IMailToProvider.GetMR_SMR_MailAddress(
                          new RFT_OrderComments()
                          {
                              OrderID = OrderID,
                          }, "200").ToList();

                        SendMail_Request request = new SendMail_Request()
                        {
                            To = mailToAddress[0].ToAddress,
                            CC = string.Empty,
                            Subject = mailToSubject[0].Subject.ToString().Replace("{0}", OrderID),
                            Body = mailToSubject[0].Content,
                        };

                        SendMail_Result result = MailTools.SendMail(request);

                        if (!string.IsNullOrEmpty(result.resultMsg))
                        {
                            inspections.ErrMsg = result.resultMsg;
                            inspections.Result = false;
                        }
                        else
                        {
                            inspections.Result = true;
                            inspections.ErrMsg = string.Empty;
                        }

                    }
                    else
                    {
                        inspections.Result = true;
                        inspections.ErrMsg = string.Empty;
                    }
                }
                else
                {
                    inspections.Result = false;
                    inspections.ErrMsg = "Inspection update row count = 0";
                }

                #endregion
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                inspections.Result = false;
                inspections.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return inspections;
        }

        public InspectionSave_ViewModel ChkInspQty(InspectionSave_ViewModel inspections)
        {
            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);
            // 判斷檢驗數量不可超過Order_Qty
            DataTable dtchk = _RFTInspectionDetailProvider.ChkInspQty(inspections.rft_Inspection);
            if (dtchk == null || dtchk.Rows.Count == 0)
            {
                inspections.Result = false;
                inspections.ErrMsg = "Inspection Qty cannot over than Order Qty.";
            }
            else
            {
                inspections.Result = true;
            }

            return inspections;
        }

        public IList<ReworkCard> GetReworkCards(ReworkCard rework)
        {
            _IReworkCardProvider = new ReworkCardProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<ReworkCard> reworkCards = new List<ReworkCard>();
            try
            {
                reworkCards = _IReworkCardProvider.Get(
                  new ReworkCard()
                  {
                      FactoryID = rework.FactoryID,
                      Line = rework.Line,
                      Type = rework.Type,
                  }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return reworkCards;
        }

        public List<ReworkList_ViewModel> GetReworkList(ReworkList_ViewModel reworkList)
        {
            _IReworkListProvider = new ReworkListProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<ReworkList_ViewModel> reworkList_Views = new List<ReworkList_ViewModel>();
            try
            {
                reworkList_Views = _IReworkListProvider.Get(
               new ReworkList_ViewModel()
               {
                   FactoryID = reworkList.FactoryID,
                   Line = reworkList.Line,
               }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return reworkList_Views;
        }

        public InspectionSave_ViewModel SaveReworkListAction(List<RFT_Inspection> rFT_Inspections, ReworkListType reworkListType)
        {
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            InspectionSave_ViewModel reworkList_SaveView = new InspectionSave_ViewModel();

            try
            {
                _RFTInspectionProvider = new RFTInspectionProvider(_ISQLDataTransaction);
                int updateCnt = _RFTInspectionProvider.SaveReworkListAction(rFT_Inspections, reworkListType.ToString());
                _ISQLDataTransaction.Commit();
                if (updateCnt == 0)
                {
                    reworkList_SaveView.Result = false;
                    reworkList_SaveView.ErrMsg = "Rework List Action update row count = 0";
                }
                else
                {
                    reworkList_SaveView.Result = true;
                    reworkList_SaveView.ErrMsg = string.Empty;
                }
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                reworkList_SaveView.Result = false;
                reworkList_SaveView.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return reworkList_SaveView;
        }

        public InspectionSave_ViewModel SaveReworkListAddReject(RFT_Inspection_Detail detail)
        {
            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            InspectionSave_ViewModel reworkList_SaveView = new InspectionSave_ViewModel()
            {
                Result = true,
                ErrMsg = string.Empty,
            };

            try
            {
                // 確認ID是否存在於表頭               
                List<RFT_Inspection> inspections = _RFTInspectionProvider.Get(new RFT_Inspection()
                {
                    ID = detail.ID,
                }).ToList();

                if (inspections.Count == 0)
                {
                    reworkList_SaveView.Result = false;
                    reworkList_SaveView.ErrMsg = "ReworkList Reject save is failed";
                    return reworkList_SaveView;
                }

                // 新增一筆RFT_Inspection_Detail,  ID 會傳入。
                _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(_ISQLDataTransaction);
                int updateCnt = _RFTInspectionDetailProvider.Create_Detail(detail);
                _ISQLDataTransaction.Commit();
                if (updateCnt == 0 && detail.ID == 0)
                {
                    reworkList_SaveView.Result = false;
                    reworkList_SaveView.ErrMsg = "ReworkList Reject save is failed";
                }
                else
                {
                    reworkList_SaveView.Result = true;
                    reworkList_SaveView.ErrMsg = string.Empty;
                }
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                reworkList_SaveView.Result = false;
                reworkList_SaveView.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return reworkList_SaveView;
        }

        public InspectionSave_ViewModel SaveReworkListDelete(LogIn_Request logIn_Request, List<RFT_Inspection> rFT_Inspection)
        {
            IQualityPass1Provider QualityPass1Provider = QualityPass1Provider = new QualityPass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            ManufacturingExecutionDataAccessLayer.Interface.IPass1Provider MESPass1Provider = MESPass1Provider = new ManufacturingExecutionDataAccessLayer.Provider.MSSQL.Pass1Provider(Common.ManufacturingExecutionDataAccessLayer);
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            InspectionSave_ViewModel reworkList_SaveView = new InspectionSave_ViewModel()
            {
                Result = true,
            };
            LogIn_Result result = new LogIn_Result();

            // 驗證 LogIn_Request
            try
            {
                List<Quality_Pass1> quality_Pass1s = QualityPass1Provider.Get(new Quality_Pass1() { ID = logIn_Request.UserID }).ToList();
                if (quality_Pass1s.Count == 0)
                {
                    throw new Exception("User ID not exist.");
                }

                // 先判斷ID，在判斷密碼。
                // ID 不存在改抓MES PASS1
                ProductionDataAccessLayer.Interface.IPass1Provider PMSPass1Provider = PMSPass1Provider = new ProductionDataAccessLayer.Provider.MSSQL.Pass1Provider(Common.ProductionDataAccessLayer);
                List<DatabaseObject.ProductionDB.Pass1> pmsPass1 = PMSPass1Provider.Get(new DatabaseObject.ProductionDB.Pass1() { ID = logIn_Request.UserID }).ToList();
                if (pmsPass1.Count == 0)
                {
                    // 改抓MES PASS1
                    List<DatabaseObject.ManufacturingExecutionDB.Pass1> mesPass1 = MESPass1Provider.Get(new DatabaseObject.ManufacturingExecutionDB.Pass1() { ID = logIn_Request.UserID, Password = logIn_Request.Password.ToUpper() }).ToList();
                    if (mesPass1.Count == 0)
                    {
                        throw new Exception("Incorrect password.");
                    }
                }
                else if (!pmsPass1.Where(x => x.Password.ToUpper().Equals(logIn_Request.Password.ToUpper())).Any())
                {
                    throw new Exception("Incorrect Password.");
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                reworkList_SaveView.Result = false;
                reworkList_SaveView.ErrMsg = ex.Message.Replace("'", string.Empty);
                result.Result = false;
            }

            // 成功 依RFT_Inspection.ID EditName  更新ReworkCard、 刪除RFT_Inspection、RFT_Inspection_Detail
            if (result.Result == true)
            {
                try
                {
                    _RFTInspectionProvider = new RFTInspectionProvider(_ISQLDataTransaction);
                    if (rFT_Inspection == null)
                    {
                        reworkList_SaveView.Result = false;
                        reworkList_SaveView.ErrMsg = "Please select data before click.";
                        return reworkList_SaveView;
                    }
                    int updateCnt = _RFTInspectionProvider.SaveReworkListDelete(rFT_Inspection);
                    _ISQLDataTransaction.Commit();
                    if (updateCnt == 0)
                    {
                        reworkList_SaveView.Result = false;
                        reworkList_SaveView.ErrMsg = "Rework List Delete row count = 0";
                    }
                    else
                    {
                        reworkList_SaveView.Result = true;
                        reworkList_SaveView.ErrMsg = string.Empty;
                    }
                }
                catch (Exception ex)
                {
                    _ISQLDataTransaction.RollBack();
                    reworkList_SaveView.Result = false;
                    reworkList_SaveView.ErrMsg = ex.Message.Replace("'", string.Empty);
                }
                finally { _ISQLDataTransaction.CloseConnection(); }
            }

            return reworkList_SaveView;
        }

        public List<DQSReason> GetDQSReason(DQSReason dQSReason)
        {
            _IDQSReasonProvider = new DQSReasonProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<DQSReason> reasons = new List<DQSReason>();
            try
            {
                reasons = _IDQSReasonProvider.Get(
                  new DQSReason()
                  {
                      Type = dQSReason.Type.ToString(),
                  }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return reasons;
        }

        public List<RFT_Inspection_Measurement_ViewModel> GetMeasurement(string OrderID, string SizeCode, string UserID)
        {
            _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _IStyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            List<RFT_Inspection_Measurement_ViewModel> _Inspection_Measurement_ViewModels = new List<RFT_Inspection_Measurement_ViewModel>();
            try
            {
                IList<Orders> ordersList = _IOrdersProvider.Get(new Orders() { ID = OrderID });
                string strSizeUnit = string.Empty;
                string longStyleUkey = string.Empty;
                if (ordersList.Count > 0)
                {
                    longStyleUkey = ordersList[0].StyleUkey.ToString();
                    IList<Style> StyleList = _IStyleProvider.GetSizeUnit(Convert.ToInt64(longStyleUkey));

                    strSizeUnit = StyleList[0].SizeUnit;
                }

                _IRFTInspectionMeasurementProvider = new RFTInspectionMeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
                _Inspection_Measurement_ViewModels = _IRFTInspectionMeasurementProvider.Get(Convert.ToInt64(longStyleUkey), SizeCode, UserID).ToList();
                    
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
            return _Inspection_Measurement_ViewModels;
        }

        public RFT_Inspection_Measurement_ViewModel SaveMeasurement(List<RFT_Inspection_Measurement> Measurement)
        {
            _IRFTInspectionMeasurementProvider = new RFTInspectionMeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
            RFT_Inspection_Measurement_ViewModel _Measurement_ViewModel = new RFT_Inspection_Measurement_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            try
            {
                _IRFTInspectionMeasurementProvider = new RFTInspectionMeasurementProvider(_ISQLDataTransaction);
                int updateCnt = _IRFTInspectionMeasurementProvider.Save(Measurement);
                _ISQLDataTransaction.Commit();
                if (updateCnt == 0)
                {
                    throw new Exception("Save Fail");
                }

                _Measurement_ViewModel.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                _Measurement_ViewModel.Result = false;
                _Measurement_ViewModel.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return _Measurement_ViewModel;
        }

        public List<RFT_OrderComments_ViewModel> GetRFT_OrderComments(RFT_OrderComments rFT_OrderComments)
        {
            _IRFTOrderCommentsProvider = new RFTOrderCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<RFT_OrderComments_ViewModel> rFT_OrderComments_ViewModel = new List<RFT_OrderComments_ViewModel>();
            try
            {
                rFT_OrderComments_ViewModel = _IRFTOrderCommentsProvider.Get(
                new RFT_OrderComments()
                {
                    OrderID = rFT_OrderComments.OrderID,
                }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return rFT_OrderComments_ViewModel;
        }

        public RFT_OrderComments_ViewModel SaveRFT_OrderComments(List<RFT_OrderComments> rFT_OrderComments)
        {
            _IRFTOrderCommentsProvider = new RFTOrderCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            RFT_OrderComments_ViewModel rFT_OrderComments_ViewModel = new RFT_OrderComments_ViewModel();
            try
            {
                _IRFTOrderCommentsProvider = new RFTOrderCommentsProvider(_ISQLDataTransaction);
                int updateCnt = _IRFTOrderCommentsProvider.Save_upd_ins(rFT_OrderComments);
                _ISQLDataTransaction.Commit();
                if (updateCnt == 0)
                {
                    rFT_OrderComments_ViewModel.Result = false;
                    rFT_OrderComments_ViewModel.ErrMsg = "Save RFT_OrderComments row count = 0";
                }
                else
                {
                    rFT_OrderComments_ViewModel.Result = true;
                    rFT_OrderComments_ViewModel.ErrMsg = string.Empty;
                }
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                rFT_OrderComments_ViewModel.Result = false;
                rFT_OrderComments_ViewModel.ErrMsg = ex.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return rFT_OrderComments_ViewModel;
        }

        public RFT_OrderComments_ViewModel SendMailRFT_OrderComments(RFT_OrderComments rFT_OrderComments, string UserID)
        {
            RFT_OrderComments_ViewModel rFT_OrderComments_ViewModel = new RFT_OrderComments_ViewModel();
            try
            {
                // 撈資料
                List<RFT_OrderComments_ViewModel> queryData = GetRFT_OrderComments(rFT_OrderComments);
                _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
                Orders orders = _IOrdersProvider.Get(new Orders { ID = rFT_OrderComments.OrderID }).FirstOrDefault();
                #region 寄信
                // 取得 mail to address
                _IMailToProvider = new MailToProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<MailTo> mailToAddress = _IMailToProvider.GetMR_SMR_MailAddress(
                  new RFT_OrderComments()
                  {
                      OrderID = rFT_OrderComments.OrderID,
                  }, "201").ToList();


                List<MailTo> mailToSubject = _IMailToProvider.Get(
                   new MailTo()
                   {
                       ID = "201",
                   }).ToList();

                if (mailToAddress.Count == 0)
                {
                    rFT_OrderComments_ViewModel.ErrMsg = $"Result MSG: mail address is empty! SP#: {rFT_OrderComments.OrderID}";
                    rFT_OrderComments_ViewModel.Result = false;

                    return rFT_OrderComments_ViewModel;
                }

                if (mailToSubject.Count == 0)
                {
                    rFT_OrderComments_ViewModel.ErrMsg = $"Result message: subject is empty! SP#: {rFT_OrderComments.OrderID}";
                    rFT_OrderComments_ViewModel.Result = false;

                    return rFT_OrderComments_ViewModel;
                }

                if (queryData.Count == 0)
                {
                    rFT_OrderComments_ViewModel.ErrMsg = $"Result message: Comments data is empty! SP#: {rFT_OrderComments.OrderID}";
                    rFT_OrderComments_ViewModel.Result = false;

                    return rFT_OrderComments_ViewModel;
                }

                string html = @"
<style>
.CFTCommentsTable {
width: 69vw;
}

.CFTCommentsTable > thead > tr > th {
text-align: center;
vertical-align: middle;
}

.CFTCommentsTable > tbody > tr > td {
text-align: center;
vertical-align: middle;
}

.CFTCommentsTable > tbody > tr > td > textarea {
    min-height: 14vh;
}

 .DefectTable {
        font-size: 1rem;
        font-weight: bold;
        border: solid 1px black;
    }

        .DefectTable > thead > tr {
            background-color: gray;
        }

            .DefectTable > thead > tr > th {
                padding: 1em 2em 1em 2em;
            }

        .DefectTable > tbody > tr > td {
            border: solid 1px gray;
            cursor: pointer;
        }

        .DefectTable .tdEmpty {
            padding: 2em 1em 2em 1em;
        }

        .DefectTable .tdValue {
            padding: 1em;
        }
</style>";
                html += $@"
<table class='CFTCommentsTable DefectTable'>
<tbody>
    <tr style='width:17vw;'>
        <td><p>Style</p></td>
        <td><p>{orders.StyleID}</p></td>
    </tr>
    <tr>
        <td><p>Sample Stage</p></td>
        <td><p>{orders.OrderTypeID}</p></td>
    </tr>
    <tr>
        <td><p>Season</p></td>
        <td><p>{orders.SeasonID}</p></td>
    </tr>
    <tr>
        <td><p>Data Released</p></td>
        <td><p>{DateTime.Now.ToString("yyyy/MM/dd")}</p></td>
    </tr>
    <tr>
        <td><p>Released By</p></td>
        <td>{UserID}</td>
    </tr>
</tbody>
</table>
";

                html += @"
<table class='CFTCommentsTable DefectTable'>
<thead>
<tr>
<th style = 'width:17vw;'> 
    <p> Comments Category </p>    
</th>    
<th>    
    <p> Comments </p>    
</th>    
</tr>    
</thead>    
<tbody>
    ";
                foreach (RFT_OrderComments_ViewModel item in queryData)
                {
                    html += "<tr>";
                    html += "<td><p>" + item.PMS_RFTCommentsDescription + "</p></td>";
                    html += "<td><textarea id='" + item.PMS_RFTCommentsDescription + item.PMS_RFTCommentsID + "' idx='" + item.PMS_RFTCommentsID + "' cols='85'> " + item.Comnments + " </textarea></td>";
                    html += "</tr>";
                }

                html += "   </tbody> </table>";

                SendMail_Request request = new SendMail_Request()
                {
                    To = mailToAddress[0].ToAddress,
                    CC = string.Empty,
                    Subject = mailToSubject[0].Subject.ToString().Replace("{0}", rFT_OrderComments.OrderID),
                    Body = html,
                };

                SendMail_Result result = MailTools.SendMail(request);

                if (!result.result)
                {
                    rFT_OrderComments_ViewModel.Result = false;
                    rFT_OrderComments_ViewModel.ErrMsg = result.resultMsg;
                    return rFT_OrderComments_ViewModel;
                }

                #endregion

                rFT_OrderComments_ViewModel.Result = true;
            }
            catch (Exception ex)
            {
                rFT_OrderComments_ViewModel.Result = false;
                rFT_OrderComments_ViewModel.ErrMsg = ex.Message.Replace("'", string.Empty);
            }
            return rFT_OrderComments_ViewModel;
        }

        public RFT_PicDuringDummyFitting GetRFT_PicDuringDummyFitting(RFT_PicDuringDummyFitting picDuringDummyFitting)
        {
            _IRFTPicDuringDummyFittingProvider = new RFTPicDuringDummyFittingProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<RFT_PicDuringDummyFitting> PicDuringDummyFitting = new List<RFT_PicDuringDummyFitting>();
            try
            {
                PicDuringDummyFitting = _IRFTPicDuringDummyFittingProvider.Get(
                new RFT_PicDuringDummyFitting()
                {
                    OrderID = picDuringDummyFitting.OrderID,
                    Article = picDuringDummyFitting.Article,
                    Size = picDuringDummyFitting.Size
                }).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return PicDuringDummyFitting.FirstOrDefault();
        }

        public RFT_PicDuringDummyFitting_ViewModel SaveRFT_PicDuringDummyFitting(RFT_PicDuringDummyFitting picDuringDummyFitting)
        {
            _IRFTPicDuringDummyFittingProvider = new RFTPicDuringDummyFittingProvider(Common.ManufacturingExecutionDataAccessLayer);
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            RFT_PicDuringDummyFitting_ViewModel rFT_OrderComments_ViewModel = new RFT_PicDuringDummyFitting_ViewModel();
            try
            {
                _IRFTPicDuringDummyFittingProvider = new RFTPicDuringDummyFittingProvider(_ISQLDataTransaction);
                int updateCnt = _IRFTPicDuringDummyFittingProvider.Save_Upd_Ins(picDuringDummyFitting);
                _ISQLDataTransaction.Commit();
                rFT_OrderComments_ViewModel.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                rFT_OrderComments_ViewModel.Result = false;
                rFT_OrderComments_ViewModel.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return rFT_OrderComments_ViewModel;
        }

    }
}
