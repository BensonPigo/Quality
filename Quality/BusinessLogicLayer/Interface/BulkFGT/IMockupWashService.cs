using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupWashService
    {
        MockupWashs_ViewModel GetMockupWash(MockupWash_Request MockupWash);

        List<SelectListItem> GetAccessoryRefNo(AccessoryRefNo_Request Request);

        MockupWashs_ViewModel Create(MockupWashs_ViewModel MockupWash);

        MockupWash_ViewModel GetExcel(string ReportNo);
    }
}
