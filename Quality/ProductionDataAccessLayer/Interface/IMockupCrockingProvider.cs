using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingProvider
    {
        IList<MockupCrocking_ViewModel> GetMockupCrocking(MockupCrocking Item);

        int Create(MockupCrocking Item);

        int Update(MockupCrocking Item);

        int Delete(MockupCrocking Item);
    }
}
