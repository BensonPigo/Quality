using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface
{
    public interface IStyleManagementService
    {
        IList<SelectListItem> GetStyles(StyleManagement_Request Req);
        IList<SelectListItem> GetBrands();
        IList<SelectListItem> GetSeasons(string brandID);

        IList<StyleResult_ViewModel> Get_StyleResult_Browse(StyleManagement_Request styleResult_Request);
        IList<StyleResult_ViewModel> Get_PrintBarcodeStyleInfo(StyleManagement_Request styleResult_Request);
    }
}
