using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrockings_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking);

        MockupCrockings_ViewModel Create(MockupCrockings_ViewModel MockupCrocking);

        MockupCrocking_ViewModel GetExcel(string ReportNo);
    }
}
