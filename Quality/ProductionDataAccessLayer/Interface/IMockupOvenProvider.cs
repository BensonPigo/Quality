using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupOvenProvider
    {
        IList<MockupOven_ViewModel> GetMockupOvenReportNoList(MockupOven_Request Item);

        IList<MockupOven_ViewModel> GetMockupOven(MockupOven_Request Item, bool istop1 = false);

        int Create(MockupOven Item);

        int Update(MockupOven Item);

        int Delete(MockupOven Item);
    }
}
