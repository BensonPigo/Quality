using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel
{
    public class MockupCrocking_ViewModel : ResultModelBase<MockupCrocking_Result>
    {
        public string ReportNo { get; set; }
        public string TempFileName { get; set; }
        public List<MockupCrocking> MockupCrocking { get; set; }
    }

    public class MockupCrocking_Result
    {

    }
}
