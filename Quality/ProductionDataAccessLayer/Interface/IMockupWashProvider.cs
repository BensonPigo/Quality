using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupWashProvider
    {
        string GetFactoryNameEN(string ReportNo, string FactoryID);
        IList<MockupWash_ViewModel> GetMockupWashReportNoList(MockupWash_Request Item);

        IList<MockupWash_ViewModel> GetMockupWash(MockupWash_Request Item, bool istop1 = false);

        DataTable GetMockupWashFailMailContentData(string ReportNo);

        int CreateMockupWash(MockupWash Item, string Mdivision, out string NewReportNo);

        void UpdateMockupWash(MockupWash_ViewModel Item);

        int DeleteMockupWash(MockupWash Item);

        int CreateDetail(MockupWash_Detail Item);
        int DeleteDetail(MockupWash_Detail Item);
        IList<MockupWash_Detail_ViewModel> GetMockupWash_Detail(MockupWash_Detail Item);
    }
}
