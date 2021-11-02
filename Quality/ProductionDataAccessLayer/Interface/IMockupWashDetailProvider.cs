using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupWashDetailProvider
    {
        IList<MockupWash_Detail_ViewModel> GetMockupWash_Detail(MockupWash_Detail Item);

        int Create(MockupWash_Detail Item);
        int Update(MockupWash_Detail Item);
        int Delete(MockupWash_Detail Item);
    }
}
