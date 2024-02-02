using System;

namespace DatabaseObject.RequestModel
{
    public class AccessoryRefNo_Request
    {
        public long StyleUkey { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string MtlTypeID { get; set; }
    }
}
