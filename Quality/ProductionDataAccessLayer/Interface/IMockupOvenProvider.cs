using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupOvenProvider
    {
        IList<MockupOven_ViewModel> GetMockupOvenReportNoList(MockupOven_Request Item);

        IList<MockupOven_ViewModel> GetMockupOven(MockupOven_Request Item, bool istop1 = false);

        DataTable GetMockupOvenFailMailContentData(string ReportNo);

        int Create(MockupOven Item, string Mdivision, out string NewReportNo);

        void Update(MockupOven_ViewModel Item);

        int Delete(MockupOven Item);
    }
}
