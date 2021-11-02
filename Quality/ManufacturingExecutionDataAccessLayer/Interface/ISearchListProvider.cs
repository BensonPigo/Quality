using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface ISearchListProvider
    {
        IList<SelectListItem> GetTypeDatasource(string Pass1ID);
        IList<SearchList_Result> Get_SearchList(SearchList_ViewModel Req);
    }
}
