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

        public IList<GarmentTest_ViewModel> CheckSelectItemData(GarmentTest_ViewModel garmentTest_ViewModel, SelectType type)
        {
            List<GarmentTest_ViewModel> result = new List<GarmentTest_ViewModel>();
            return result;
        }
    }
}
