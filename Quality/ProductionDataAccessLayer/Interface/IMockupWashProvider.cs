using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupWashProvider
    {
        IList<MockupWash_ViewModel> GetMockupWashReportNoList(MockupWash_Request Item);

        IList<MockupWash_ViewModel> GetMockupWash(MockupWash_Request Item, bool istop1 = false);

        DataTable GetMockupWashFailMailContentData(string ReportNo);

        int Create(MockupWash Item);

        void Update(MockupWash_ViewModel Item);

        int Delete(MockupWash Item);
    }
}
