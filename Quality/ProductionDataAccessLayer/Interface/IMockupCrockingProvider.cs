using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Data;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingProvider
    {
        IList<MockupCrocking_ViewModel> GetMockupCrockingReportNoList(MockupCrocking_Request Item);

        IList<MockupCrocking_ViewModel> GetMockupCrocking(MockupCrocking_Request Item, bool istop1 = false);

        DataTable GetMockupCrockingFailMailContentData(string ReportNo);

        int Create(MockupCrocking Item, string Mdivision, out string NewReportNo);

        void Update(MockupCrocking_ViewModel Item);

        int Delete(MockupCrocking Item);
    }
}
