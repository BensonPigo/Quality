using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingDetailProvider
    {
        IList<MockupCrocking_Detail_ViewModel> GetMockupCrocking_Detail(MockupCrocking_Detail Item);

        int Create(MockupCrocking_Detail Item);
        int Update(MockupCrocking_Detail Item);
        int Delete(MockupCrocking_Detail Item);
    }
}
