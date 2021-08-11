using BusinessLogicLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
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

        // Production
        private IOrdersProvider _IOrdersProvider;

        public enum SelectType
        {
            OrderID ,
            StyleID ,
            Article,
            Size ,
            ProductType ,
        }

        public enum DefectType
        {
            Responsibility,
            BAAuditCriteria,
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
                switch(type)
                {
                    case SelectType.OrderID:
                        result = result.GroupBy(x => new { x.StyleID, x.OrderID, x.ProductTypePMS, x.Brand, x.Season, x.SampleStage, x.OriginalLine, x.OrderQty })
                                    .Select(x => new Inspection_ViewModel() 
                                    {
                                       StyleID = x.Key.StyleID,
                                       OrderID = x.Key.OrderID,
                                       ProductTypePMS = x.Key.ProductTypePMS,
                                       Brand = x.Key.Brand,
                                       Season=  x.Key.Season,
                                       SampleStage = x.Key.SampleStage,
                                       OriginalLine = x.Key.OriginalLine,
                                       OrderQty = x.Key.OrderQty
                                    }).ToList();
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

        public Inspection_ViewModel GetTop3(Inspection_ViewModel inspection_ViewModel)
        {
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);

            Inspection_ViewModel result = new Inspection_ViewModel();
            List<RFT_Inspection> inspections = _RFTInspectionProvider.Get(new RFT_Inspection() 
                                                { 
                                                    FactoryID = inspection_ViewModel.FactoryID ,
                                                    Line = inspection_ViewModel.Line,
                                                    InspectionDate = inspection_ViewModel.InspectionDate,
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
            result.Hard = inspections.Where(x => x.FixType.Equals("Hard")).Count().ToString();
            result.Quick = inspections.Where(x => x.FixType.Equals("Quick")).Count().ToString();
            result.Wash = inspections.Where(x => x.FixType.Equals("Wash")).Count().ToString();
            result.Repl = inspections.Where(x => x.FixType.Equals("Repl.")).Count().ToString();
            result.Print = inspections.Where(x => x.FixType.Equals("Print")).Count().ToString();
            result.Shade = inspections.Where(x => x.FixType.Equals("Shade")).Count().ToString();
            result.Dispose = inspections.Where(x => x.Status.Equals("Dispose")).Count().ToString();

            List<string> top3DefectCode = inspection_details
                .Where(x => !string.IsNullOrEmpty(x.DefectCode))
                .GroupBy(x => x.DefectCode)
                .Select(x => new { DefectCode = x.Key, cnt = x.Count() })
                .OrderByDescending(x => x.cnt)
                .Take(3)
                .Select(x => x.DefectCode).ToList();

            List<string> top3Area = inspection_details
                .Where(x => !string.IsNullOrEmpty(x.AreaCode) && top3DefectCode.Contains(x.DefectCode))
                .GroupBy(x => x.AreaCode)
                .Select(x => new { AreaCode = x.Key, cnt = x.Count() })
                .OrderByDescending(x => x.cnt)
                .Take(3)
                .Select(x => x.AreaCode).ToList();

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
            inspections.Result = true;
            inspections.ErrMsg = string.Empty;

            if (inspections.rft_Inspection == null)
            {
                inspections.Result = false;
                inspections.ErrMsg = "InspectionSave_ViewModel is null";
                return inspections;
            }

            _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);

            string OrderID = inspections.rft_Inspection.OrderID;
            IList<Orders> orders = _IOrdersProvider.Get(new Orders() { ID = OrderID });
            if (orders.Count > 0)
            {
                string StyleUkey = orders[0].StyleUkey.ToString();
                inspections.rft_Inspection.StyleUkey = long.Parse(StyleUkey);
            }

            #region update/insert [RFT_Inspection] and [RFT_Inspection_Detail]

            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            _RFTInspectionDetailProvider = new RFTInspectionDetailProvider(Common.ManufacturingExecutionDataAccessLayer);

            int createCnt = _RFTInspectionDetailProvider.Create_Master_Detail(inspections.rft_Inspection, inspections.fT_Inspection_Details);

            if (createCnt == 0)
            {
                inspections.Result = false;
                inspections.ErrMsg = "Save RFTInspection row count = 0";
                return inspections;
            }

            #endregion

            return inspections;
        }

        public IList<ReworkCard> GetReworkCards(ReworkCard rework)
        {
            _IReworkCardProvider = new ReworkCardProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<ReworkCard> reworkCards = _IReworkCardProvider.Get(
                new ReworkCard()
                {
                    FactoryID = rework.FactoryID,
                    Line = rework.Line,
                    Type = rework.Type,
                }).ToList();

            return reworkCards;
        }

        public ReworkList_ViewModel GetReworkListFilter(ReworkList_ViewModel reworkList, SelectType type)
        {
            _RFTInspectionProvider = new RFTInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            _IReworkListProvider = new ReworkListProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<ReworkList_ViewModel> reworkList_Views = _IReworkListProvider.GetReworkListFilter(
                new ReworkList_ViewModel()
                {
                    FactoryID = reworkList.FactoryID,
                    Line = reworkList.Line,
                    Status = reworkList.Status,
                }, type.ToString()).ToList();


            // reworkList.SP = _RFTInspections.;

            //List<string> spList = new List<string>();
            //List<string> StyleList = new List<string>();
            //List<string> ArticleList = new List<string>();
            //List<string> SizeList = new List<string>();

            //foreach (var item in _RFTInspections)
            //{
            //    spList.Add(item.SP);
            //    StyleList.Add(item.Style);
            //    ArticleList.Add(item.Article);
            //    SizeList.Add(item.Size);
            //}

            //reworkList.SPList = spList;
            //reworkList.StyleList = StyleList;
            //reworkList.ArticleList = ArticleList;
            //reworkList.SizeList = SizeList;

            //_IReworkListProvider = new ReworkListProvider(Common.ManufacturingExecutionDataAccessLayer);
            //List<ReworkList> reworkLists = _IReworkListProvider.Get(
            //    new ReworkList()
            //    {
            //        FactoryID = reworkList.rft_Inspection.FactoryID,
            //        Line = reworkList.rft_Inspection.Line,
            //        Status = reworkList.rft_Inspection.Status,
            //    }).ToList();
            //reworkList.ReworkList = reworkLists;


            return reworkList;
        }
    }
}
