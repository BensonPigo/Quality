using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface ISearchListService
    {
        List<SelectListItem> GetTypeDatasource(string Pass1ID);
        SearchList_ViewModel Get_SearchList(SearchList_ViewModel Req);
        SearchList_ViewModel ToExcel(SearchList_ViewModel Req);
    }
}
