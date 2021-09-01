using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupWashService
    {
        MockupWashs_ViewModel GetMockupWash(MockupWash_Request MockupWash);

        AccessoryRefNos GetAccessoryRefNo(AccessoryRefNo_Request Request);

        MockupWashs_ViewModel Create(MockupWashs_ViewModel MockupWash);

        MockupWash_ViewModel GetExcel(string ReportNo);
    }
}
