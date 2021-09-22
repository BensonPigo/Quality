using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
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

        List<SelectListItem> GetTestingMethod();

        List<Orders> GetOrders(Orders Orders);

        List<Order_Qty> GetDistinctArticle(Order_Qty Orders);

        BaseResult Create(MockupWash_ViewModel MockupWash, string Mdivision, string userid, out string NewReportNo);

        BaseResult Update(MockupWash_ViewModel MockupWash, string userid);

        BaseResult Delete(MockupWash_ViewModel MockupWash);

        BaseResult DeleteDetail(List<MockupWash_Detail_ViewModel> MockupWashDetail);

        Report_Result GetPDF(MockupWash_ViewModel MockupWash, bool test = false);

        SendMail_Result FailSendMail(MockupFailMail_Request mail_Request);
    }
}
