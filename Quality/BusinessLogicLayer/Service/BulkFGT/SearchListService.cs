using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class SearchListService : ISearchListService
    {
        private ISearchListProvider SearchListProvider;

        public List<SelectListItem> GetTypeDatasource(string Pass1ID)
        {
            List<SelectListItem> result = new List<SelectListItem>();
            try
            {
                SearchListProvider = new SearchListProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = SearchListProvider.GetTypeDatasource(Pass1ID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

         public SearchList_ViewModel Get_SearchList(SearchList_ViewModel Req)
        {
            Req.DataList = new List<SearchList_Result>();
            SearchList_ViewModel model = Req;

            try
            {
                SearchListProvider = new SearchListProvider(Common.ProductionDataAccessLayer);
                var res = SearchListProvider.Get_SearchList(Req).ToList();

                model.DataList = res;
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

    }
}
