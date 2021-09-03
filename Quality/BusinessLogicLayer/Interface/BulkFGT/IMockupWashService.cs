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

        MockupWash_ViewModel Create(MockupWash_ViewModel MockupWash);

        MockupWash_ViewModel Update(MockupWash_ViewModel MockupWash);

        MockupWash_ViewModel Delete(MockupWash_ViewModel MockupWash);

        MockupWash_ViewModel GetPDF(MockupWash_ViewModel MockupWash, bool test = false);
    }
}
