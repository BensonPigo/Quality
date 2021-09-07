using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupWashService
    {
        MockupWash_ViewModel GetMockupWash(MockupWash_Request MockupWash);

        List<SelectListItem> GetAccessoryRefNo(AccessoryRefNo_Request Request);

        List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request);

        List<Orders> GetOrders(Orders Orders);

        List<Order_Qty> GetDistinctArticle(Order_Qty Orders);

        BaseResult Create(MockupWash_ViewModel MockupWash);

        BaseResult Update(MockupWash_ViewModel MockupWash);

        BaseResult Delete(MockupWash_ViewModel MockupWash);

        BaseResult DeleteDetail(List<MockupWash_Detail_ViewModel> MockupWashDetail);

        MockupWash_ViewModel GetPDF(MockupWash_ViewModel MockupWash, bool test = false);
    }
}
