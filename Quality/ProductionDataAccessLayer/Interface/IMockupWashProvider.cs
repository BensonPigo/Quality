using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupWashProvider
    {
        IList<MockupWash_ViewModel> GetMockupWash(MockupWash Item);

        int Create(MockupWash Item);

        int Update(MockupWash Item);

        int Delete(MockupWash Item);
    }
}
