using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingProvider
    {
        IList<MockupCrocking_ViewModel> GetMockupCrockingReportNoList(MockupCrocking_Request Item);

        IList<MockupCrocking_ViewModel> GetMockupCrocking(MockupCrocking_Request Item);

        int Create(MockupCrocking Item);

        int Update(MockupCrocking Item);

        int Delete(MockupCrocking Item);
    }
}
