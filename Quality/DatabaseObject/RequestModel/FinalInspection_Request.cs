using System;

namespace DatabaseObject.RequestModel
{
    public class FinalInspection_Query
    {
        public string SP { get; set; }
        public string POID { get; set; }
        public string StyleID { get; set; }
        public DateTime? SciDeliveryStart { get; set; } = null;
        public DateTime? SciDeliveryEnd { get; set; } = null;
        public string InspectionResult { get; set; }
    }
}
