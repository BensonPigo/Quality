using BusinessLogicLayer.Interface.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ADOHelper.Utility;

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

        public GarmentTest_Result GetGarmentTest(GarmentTest_ViewModel garmentTest_ViewModel)
        {
            _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
            _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
            GarmentTest_Result result = new GarmentTest_Result();
            try
            {
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

        public GarmentTest_ViewModel Save_GarmentTest(GarmentTest_ViewModel garmentTest_ViewModel)
        {
            GarmentTest_ViewModel result = new GarmentTest_ViewModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IGarmentTestProvider = new GarmentTestProvider(_ISQLDataTransaction);
                
            }
            catch (Exception)
            {

                throw;
            }

            return result;
        }
    }
}
