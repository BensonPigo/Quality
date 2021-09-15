using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.SampleRFT;
using ProductionDataAccessLayer.Provider.MSSQL.StyleManagement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class StyleResultService
    {
        public StyleResultProvider _Provider { get; set; }


        public StyleResult_ViewModel Get_StyleResult_Browse(StyleResult_Request styleResult_Request)
        {
            StyleResult_ViewModel result = new StyleResult_ViewModel()
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclamier = new List<StyleResult_FTYDisclamier>(),
                RRLR = new List<StyleResult_RRLR>(),
                BulkFGT = new List<StyleResult_BulkFGT>()
            };
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);

                styleResult_Request.CallType = StyleResult_Request.EnumCallType.StyleResult;

                var styleResults = _Provider.Get_StyleInfo(styleResult_Request).ToList();

                if (styleResults.Any())
                {
                    result = styleResults.FirstOrDefault();
                }
                result.SampleRFT = _Provider.Get_StyleResult_SampleRFT(styleResult_Request).ToList();
                result.FTYDisclamier = _Provider.Get_StyleResult_FTYDisclamier(styleResult_Request).ToList();
                result.RRLR = _Provider.Get_StyleResult_RRLR(styleResult_Request).ToList();
                result.BulkFGT = new List<StyleResult_BulkFGT>();


                result.Result = true;

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.MsgScript = $@"msg.WithInfo('{ex.Message}');";
            }

            return result;
        }

        public IList<SelectListItem> GetStyles(StyleResult_Request Req)
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetStyles(Req).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetBrands()
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetBrands().ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetSeasons(string brandID = "")
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetSeasons(brandID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
