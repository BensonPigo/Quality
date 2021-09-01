using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupWashProvider
    {
        IList<MockupWash_ViewModel> GetMockupWash(MockupWash_Request Item);

        IList<AccessoryRefNo> GetAccessoryRefNo(AccessoryRefNo_Request Item);

        int Create(MockupWash Item);

        int Update(MockupWash Item);

        int Delete(MockupWash Item);
    }
}
