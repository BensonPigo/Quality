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

        int CreateMockupCrocking(MockupCrocking Item, string Mdivision, out string NewReportNo);

        void UpdateMockupCrocking(MockupCrocking_ViewModel Item);

        int DeleteMockupCrocking(MockupCrocking Item);
        IList<MockupCrocking_Detail_ViewModel> GetMockupCrocking_Detail(MockupCrocking_Detail Item);

        int CreateDetail(MockupCrocking_Detail Item);
        int DeleteDetail(MockupCrocking_Detail Item);
    }
}
