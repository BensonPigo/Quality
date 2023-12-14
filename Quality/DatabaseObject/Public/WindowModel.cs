using System;
using System.ComponentModel.DataAnnotations;


namespace DatabaseObject.Public
{
    public class Window_Brand
    {
        /// <summary>客戶代碼</summary>
        [Display(Name = "客戶代碼")]
        public string ID { get; set; }
    }

    public class Window_Season
    {
        public string ID { get; set; }
        public string BrandID { get; set; }
    }

    public class Window_Style
    {
        public Int64 StyleUkey { get; set; }
        public string ID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
    }

    public class Window_Article
    {
        public string Article { get; set; }
        public Int64 StyleUkey { get; set; }
        public string OrderID { get; set; }
    }

    public class Window_Size
    {
        public string SizeCode { get; set; }
        public string Article { get; set; }
        public Int64 StyleUkey { get; set; }
        public string OrderID { get; set; }
    }

    public class Window_Technician
    {
        public string CallFunction { get; set; }
        public string Region { get; set; }

        public string ID { get; set; }
        public string Name { get; set; }
        public string ExtNo { get; set; }
        public string Factory { get; set; }
    }

    public class Window_Pass1
    {
        public string Title { get; set; }
        public string Region { get; set; }

        public string ID { get; set; }
        public string Name { get; set; }
        public string ExtNo { get; set; }
        public string Factory { get; set; }
        public string EMail { get; set; }
    }

    public class Window_LocalSupp
    {
        public string Title { get; set; }
        public string Region { get; set; }

        public string ID { get; set; }
        public string Abb { get; set; }
        public string Name { get; set; }
    }

    public class Window_TPESupp
    {
        public string Title { get; set; }
        public string Region { get; set; }

        public string ID { get; set; }
        public string Abb { get; set; }
        public string Name { get; set; }
    }

    public class Window_Po_Supp_Detail
    {
        public string POID { get; set; }
        public string FabricType { get; set; }

        public string SEQ1 { get; set; }
        public string SEQ2 { get; set; }
        public string SCIRefno { get; set; }
        public string Refno { get; set; }
        public string ColorID { get; set; }
        public string SuppID { get; set; }
    }

    public class Window_FtyInventory
    {
        public string Title { get; set; }
        public string POID { get; set; }
        public string SEQ1 { get; set; }
        public string SEQ2 { get; set; }
        public string Region { get; set; }

        public string Roll { get; set; }
        public string Dyelot { get; set; }
    }

    public class Window_Appearance
    {
        public string Lab { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
    }

    public class Window_SewingLine
    {
        public string ID { get; set; }
    }

    public class Window_Color
    {
        public string ID { get; set; }

        public string Name { get; set; }
        public string BrandID { get; set; }
    }

    public class Window_FGPT
    {
        public string TypeSelection_VersionID { get; set; }

        public string Seq { get; set; }
        public string Code { get; set; }
    }

    public class Window_Picture
    {

        public string PKey_1 { get; set; }
        public string PKey_2 { get; set; }
        public string PKey_3 { get; set; }

        public string PKey_1_val { get; set; }
        public string PKey_2_val { get; set; }
        public string PKey_3_val { get; set; }

        public string Table { get; set; }
        public string BeforeColumn { get; set; }
        public string AfterColumn { get; set; }
        public string OneColumn { get; set; }
        public string TwoColumn { get; set; }
        public string ThreeColumn { get; set; }

        public byte[] BrforeImage { get; set; }
        public byte[] AfterImage { get; set; }
        public byte[] OneImage { get; set; }
        public byte[] TwoImage { get; set; }
        public byte[] ThreeImage { get; set; }
    }
    public class Window_MartindalePillingTest
    {
        public string ReportNo { get; set; }


        public byte[] TestBeforePicture { get; set; }
        public byte[] Test500AfterPicture { get; set; }
        public byte[] Test2000AfterPicture { get; set; }
    }
    public class Window_RandomTumblePillingTest
    {
        public string ReportNo { get; set; }
        public byte[] TestFaceSideBeforePicture { get; set; }
        public byte[] TestFaceSideAfterPicture { get; set; }
        public byte[] TestBackSideBeforePicture { get; set; }
        public byte[] TestBackSideAfterPicture { get; set; }
    }
    public class Window_SinglePicture
    {

        public string PKey_1 { get; set; }
        public string PKey_2 { get; set; }
        public string PKey_3 { get; set; }

        public string PKey_1_val { get; set; }
        public string PKey_2_val { get; set; }
        public string PKey_3_val { get; set; }

        public string Table { get; set; }
        public string ColumnName { get; set; }


        public byte[] Image { get; set; }
    }

    public class Window_TestFailMail
    {
        public string FactoryID { get; set; }

        public string GroupName { get; set; }
        public string Type { get; set; }

        public string ToAddress { get; set; }

        public string CcAddress { get; set; }
    }

    public class Window_Operation
    {
        public string Operation { get; set; }

        public string EmployeeID { get; set; }
        public string Name { get; set; }
    }
    public class Window_AreaCode
    {
        public string AreaCode { get; set; }
    }
    public class Window_FabricRefNo
    {
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string SCIRefno { get; set; }
        public string Refno { get; set; }
        public string Color { get; set; }
        public string SuppID { get; set; }
    }
    public class Window_InkType
    {
        public string InkType { get; set; }
    }

    public class Window_RollDyelot
    {
        public string Roll { get; set; }
        public string Dyelot { get; set; }
    }
    public class Window_BrandBulkTestItem
    {
        public long Ukey { get; set; }
        public string BrandID { get; set; }
        public string Category { get; set; }
        public string TestClassify { get; set; }
        public string DocType { get; set; }
        public string TestItem { get; set; }
    }

}
