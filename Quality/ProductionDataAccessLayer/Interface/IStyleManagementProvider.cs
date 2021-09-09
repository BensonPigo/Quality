using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleManagementProvider
    {
        IList<SelectListItem> GetStyles(StyleManagement_Request Req);
        IList<SelectListItem> GetBrands();
        IList<SelectListItem> GetSeasons(string brandID);

        IList<StyleResult_ViewModel> Get_StyleInfo(StyleManagement_Request styleResult_Request);

        IList<StyleResult_SampleRFT> Get_StyleResult_SampleRFT(long styleUkey);
    }
}
