using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupWashService
    {
        MockupWashs_ViewModel GetMockupWash(MockupWash MockupWash);

        MockupWashs_ViewModel Create(MockupWashs_ViewModel MockupWash);

        MockupWash_ViewModel GetExcel(string ReportNo);
    }
}
