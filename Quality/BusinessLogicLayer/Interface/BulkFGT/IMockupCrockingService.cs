using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking);

        List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request);

        List<Orders> GetOrders(Orders Orders);

        List<Order_Qty> GetDistinctArticle(Order_Qty Orders);

        BaseResult Create(MockupCrocking_ViewModel MockupCrocking, string Mdivision, string userid, out string NewReportNo);

        BaseResult Update(MockupCrocking_ViewModel MockupCrocking, string userid);

        BaseResult Delete(MockupCrocking_ViewModel MockupCrocking);

        BaseResult DeleteDetail(List<MockupCrocking_Detail_ViewModel> MockupWashDetail);

        Report_Result GetPDF(MockupCrocking_ViewModel MockupCrocking, bool test = false, string AssignedFineName = "");

        SendMail_Result SendMail(MockupFailMail_Request mail_Request);
    }
}
