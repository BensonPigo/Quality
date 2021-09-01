using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrockings_ViewModel GetMockupCrocking(MockupCrocking MockupCrocking);

        MockupCrockings_ViewModel Create(MockupCrockings_ViewModel MockupCrocking);

        MockupCrocking_ViewModel GetExcel(string ReportNo);

        MockupCrockingScale GetScale();
    }
}
