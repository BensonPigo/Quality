namespace DatabaseObject.ResultModel
{
    public class StyleManagement_Request
    {
        public enum EnumCallType
        {
            PrintBarcode,
            StyleResult
        }

        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleUkey { get; set; }

        public EnumCallType CallType { get; set; }
    }
}
