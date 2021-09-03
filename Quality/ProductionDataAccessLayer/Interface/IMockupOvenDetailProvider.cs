using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupOvenDetailProvider
    {
        IList<MockupOven_Detail_ViewModel> GetMockupOven_Detail(MockupOven_Detail Item);

        int Create(MockupOven_Detail Item);
        int Update(MockupOven_Detail Item);
        int Delete(MockupOven_Detail Item);
    }
}
