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

        int CreateMockupOven(MockupOven Item, string Mdivision, out string NewReportNo);

        void UpdateMockupOven(MockupOven_ViewModel Item);

        int DeleteMockupOven(MockupOven Item);
        int DeleteDetail(MockupOven_Detail Item);

        int CreateDetail(MockupOven_Detail Item);
        IList<MockupOven_Detail_ViewModel> GetMockupOven_Detail(MockupOven_Detail Item);

    }
}
