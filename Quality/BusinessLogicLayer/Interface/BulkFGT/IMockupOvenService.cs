using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupOvenService
    {
        MockupOven_ViewModel GetMockupOven(MockupOven_Request MockupOven);

        List<SelectListItem> GetAccessoryRefNo(AccessoryRefNo_Request Request);

        List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request);

        List<Orders> GetOrders(Orders Orders);

        List<Order_Qty> GetDistinctArticle(Order_Qty Orders);

        BaseResult Create(MockupOven_ViewModel MockupOven);

        BaseResult Update(MockupOven_ViewModel MockupOven);

        BaseResult Delete(MockupOven_ViewModel MockupOven);

        BaseResult DeleteDetail(List<MockupOven_Detail_ViewModel> MockupWashDetail);

        Report_Result GetPDF(MockupOven_ViewModel MockupOven, bool test = false);
    }
}
