using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking);

        MockupCrocking_ViewModel Create(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel Update(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel Delete(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel GetPDF(MockupCrocking_ViewModel MockupCrocking, bool test = false);
    }
}
