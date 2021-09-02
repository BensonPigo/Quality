using BusinessLogicLayer.Interface;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class StyleManagementService : IStyleManagementService
    {
        public IStyleManagementProvider _StyleManagementProvider { get; set; }

        public IList<SelectListItem> GetBrands()
        {
            try
            {
                _StyleManagementProvider = new StyleManagementProvider(Common.ProductionDataAccessLayer);
                return _StyleManagementProvider.GetBrands().ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetSeasons(string brandID)
        {
            try
            {
                _StyleManagementProvider = new StyleManagementProvider(Common.ProductionDataAccessLayer);
                return _StyleManagementProvider.GetSeasons(brandID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetStyles(StyleManagement_Request Req)
        {
            try
            {
                _StyleManagementProvider = new StyleManagementProvider(Common.ProductionDataAccessLayer);
                return _StyleManagementProvider.GetStyles(Req).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<StyleResult_ViewModel> Get_PrintBarcodeStyleInfo(StyleManagement_Request styleResult_Request)
        {
            try
            {
                _StyleManagementProvider = new StyleManagementProvider(Common.ProductionDataAccessLayer);

                styleResult_Request.CallType = StyleManagement_Request.EnumCallType.StyleResult;
                IList<StyleResult_ViewModel> styleResults = _StyleManagementProvider.Get_StyleInfo(styleResult_Request);

                return styleResults;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<StyleResult_ViewModel> Get_StyleResult_Browse(StyleManagement_Request styleResult_Request)
        {
            _StyleManagementProvider = new StyleManagementProvider(Common.ProductionDataAccessLayer);

            styleResult_Request.CallType = StyleManagement_Request.EnumCallType.PrintBarcode;
            IList<StyleResult_ViewModel> styleResults = _StyleManagementProvider.Get_StyleInfo(styleResult_Request);

            return styleResults;
        }
    }
}
