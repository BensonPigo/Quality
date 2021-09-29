using ADOHelper.DBToolKit;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using ToolKit;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    /*(MockupWashProvider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/31  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/31  1.00    Admin        Create
    /// </history>
    public class MockupWashProvider : SQLDAL, IMockupWashProvider
    {
        private IMockupWashDetailProvider _MockupWashDetailProvider;
        #region 底層連線
        public MockupWashProvider(string ConString) : base(ConString) { }
        public MockupWashProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<MockupWash> Get(MockupWash Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ReportNo" + Environment.NewLine);
            SbSql.Append("        ,POID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,Article" + Environment.NewLine);
            SbSql.Append("        ,ArtworkTypeID" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,T1Subcon" + Environment.NewLine);
            SbSql.Append("        ,T2Supplier" + Environment.NewLine);
            SbSql.Append("        ,TestDate" + Environment.NewLine);
            SbSql.Append("        ,ReceivedDate" + Environment.NewLine);
            SbSql.Append("        ,ReleasedDate" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Technician" + Environment.NewLine);
            SbSql.Append("        ,MR" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,OtherMethod" + Environment.NewLine);
            SbSql.Append("        ,MethodID" + Environment.NewLine);
            SbSql.Append("        ,TestingMethod" + Environment.NewLine);
            SbSql.Append("        ,HTPlate" + Environment.NewLine);
            SbSql.Append("        ,HTFlim" + Environment.NewLine);
            SbSql.Append("        ,HTTime" + Environment.NewLine);
            SbSql.Append("        ,HTPressure" + Environment.NewLine);
            SbSql.Append("        ,HTPellOff" + Environment.NewLine);
            SbSql.Append("        ,HT2ndPressnoreverse" + Environment.NewLine);
            SbSql.Append("        ,HT2ndPressreversed" + Environment.NewLine);
            SbSql.Append("        ,HTCoolingTime" + Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture" + Environment.NewLine);
            SbSql.Append("        ,Type" + Environment.NewLine);
            SbSql.Append("FROM [MockupWash]" + Environment.NewLine);

            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupWash>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(MockupWash Item, string Mdivision, out string NewReportNo)
        {
            NewReportNo = GetID(Mdivision + "WA", "MockupWash", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [MockupWash]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ReportNo" + Environment.NewLine);
            SbSql.Append("        ,POID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,Article" + Environment.NewLine);
            SbSql.Append("        ,ArtworkTypeID" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,T1Subcon" + Environment.NewLine);
            SbSql.Append("        ,T2Supplier" + Environment.NewLine);
            SbSql.Append("        ,TestDate" + Environment.NewLine);
            SbSql.Append("        ,ReceivedDate" + Environment.NewLine);
            SbSql.Append("        ,ReleasedDate" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Technician" + Environment.NewLine);
            SbSql.Append("        ,MR" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,OtherMethod" + Environment.NewLine);
            SbSql.Append("        ,MethodID" + Environment.NewLine);
            SbSql.Append("        ,TestingMethod" + Environment.NewLine);
            SbSql.Append("        ,HTPlate" + Environment.NewLine);
            SbSql.Append("        ,HTFlim" + Environment.NewLine);
            SbSql.Append("        ,HTTime" + Environment.NewLine);
            SbSql.Append("        ,HTPressure" + Environment.NewLine);
            SbSql.Append("        ,HTPellOff" + Environment.NewLine);
            SbSql.Append("        ,HT2ndPressnoreverse" + Environment.NewLine);
            SbSql.Append("        ,HT2ndPressreversed" + Environment.NewLine);
            SbSql.Append("        ,HTCoolingTime" + Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture" + Environment.NewLine);
            SbSql.Append("        ,Type" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ReportNo"); objParameter.Add("@ReportNo", DbType.String, NewReportNo);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, HttpUtility.HtmlDecode(Item.POID) ?? string.Empty);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, HttpUtility.HtmlDecode(Item.StyleID) ?? string.Empty);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, HttpUtility.HtmlDecode(Item.SeasonID) ?? string.Empty);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, HttpUtility.HtmlDecode(Item.BrandID) ?? string.Empty);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, HttpUtility.HtmlDecode(Item.Article) ?? string.Empty);
            SbSql.Append("        ,@ArtworkTypeID"); objParameter.Add("@ArtworkTypeID", DbType.String, HttpUtility.HtmlDecode(Item.ArtworkTypeID) ?? string.Empty);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(Item.Remark) ?? string.Empty);
            SbSql.Append("        ,@T1Subcon"); objParameter.Add("@T1Subcon", DbType.String, HttpUtility.HtmlDecode(Item.T1Subcon) ?? string.Empty);
            SbSql.Append("        ,@T2Supplier"); objParameter.Add("@T2Supplier", DbType.String, HttpUtility.HtmlDecode(Item.T2Supplier) ?? string.Empty);
            SbSql.Append("        ,@TestDate"); objParameter.Add("@TestDate", DbType.Date, Item.TestDate);
            SbSql.Append("        ,@ReceivedDate"); objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate);
            SbSql.Append("        ,@ReleasedDate"); objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, HttpUtility.HtmlDecode(Item.Result) ?? string.Empty);
            SbSql.Append("        ,@Technician"); objParameter.Add("@Technician", DbType.String, HttpUtility.HtmlDecode(Item.Technician) ?? string.Empty);
            SbSql.Append("        ,@MR"); objParameter.Add("@MR", DbType.String, HttpUtility.HtmlDecode(Item.MR) ?? string.Empty);
            SbSql.Append("        ,GETDATE()");
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, HttpUtility.HtmlDecode(Item.AddName) ?? string.Empty);
            SbSql.Append("        ,@OtherMethod"); objParameter.Add("@OtherMethod", DbType.Boolean, Item.OtherMethod);
            SbSql.Append("        ,@MethodID"); objParameter.Add("@MethodID", DbType.String, HttpUtility.HtmlDecode(Item.MethodID) ?? string.Empty);
            SbSql.Append("        ,@TestingMethod"); objParameter.Add("@TestingMethod", DbType.String, HttpUtility.HtmlDecode(Item.TestingMethod) ?? string.Empty);
            SbSql.Append("        ,@HTPlate"); objParameter.Add("@HTPlate", DbType.Int32, Item.HTPlate);
            SbSql.Append("        ,@HTFlim"); objParameter.Add("@HTFlim", DbType.Int32, Item.HTFlim);
            SbSql.Append("        ,@HTTime"); objParameter.Add("@HTTime", DbType.Int32, Item.HTTime);
            SbSql.Append("        ,@HTPressure"); objParameter.Add("@HTPressure", DbType.Decimal, Item.HTPressure);
            SbSql.Append("        ,@HTPellOff"); objParameter.Add("@HTPellOff", DbType.String, HttpUtility.HtmlDecode(Item.HTPellOff) ?? string.Empty);
            SbSql.Append("        ,@HT2ndPressnoreverse"); objParameter.Add("@HT2ndPressnoreverse", DbType.Int32, Item.HT2ndPressnoreverse);
            SbSql.Append("        ,@HT2ndPressreversed"); objParameter.Add("@HT2ndPressreversed", DbType.Int32, Item.HT2ndPressreversed);
            SbSql.Append("        ,@HTCoolingTime"); objParameter.Add("@HTCoolingTime", DbType.String, HttpUtility.HtmlDecode(Item.HTCoolingTime) ?? string.Empty);

            SbSql.Append("        ,@TestBeforePicture");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            SbSql.Append("        ,@TestAfterPicture");
            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, HttpUtility.HtmlDecode(Item.Type) ?? string.Empty);
            SbSql.Append(")" + Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public void Update(MockupWash_ViewModel Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [MockupWash]" + Environment.NewLine);
            SbSql.Append("SET EditDate=GETDATE()" + Environment.NewLine);
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName" + Environment.NewLine); objParameter.Add("@EditName", DbType.String, HttpUtility.HtmlDecode(Item.EditName)); }
            if (Item.POID != null) { SbSql.Append(",POID=@POID" + Environment.NewLine); objParameter.Add("@POID", DbType.String, HttpUtility.HtmlDecode(Item.POID) ?? string.Empty); }
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID" + Environment.NewLine); objParameter.Add("@StyleID", DbType.String, HttpUtility.HtmlDecode(Item.StyleID) ?? string.Empty); }
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID" + Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, HttpUtility.HtmlDecode(Item.SeasonID) ?? string.Empty); }
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID" + Environment.NewLine); objParameter.Add("@BrandID", DbType.String, HttpUtility.HtmlDecode(Item.BrandID) ?? string.Empty); }
            if (Item.Article != null) { SbSql.Append(",Article=@Article" + Environment.NewLine); objParameter.Add("@Article", DbType.String, HttpUtility.HtmlDecode(Item.Article) ?? string.Empty); }
            if (Item.ArtworkTypeID != null) { SbSql.Append(",ArtworkTypeID=@ArtworkTypeID" + Environment.NewLine); objParameter.Add("@ArtworkTypeID", DbType.String, HttpUtility.HtmlDecode(Item.ArtworkTypeID) ?? string.Empty); }
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark" + Environment.NewLine); objParameter.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(Item.Remark) ?? string.Empty); }
            if (Item.T1Subcon != null) { SbSql.Append(",T1Subcon=@T1Subcon" + Environment.NewLine); objParameter.Add("@T1Subcon", DbType.String, HttpUtility.HtmlDecode(Item.T1Subcon) ?? string.Empty); }
            if (Item.T2Supplier != null) { SbSql.Append(",T2Supplier=@T2Supplier" + Environment.NewLine); objParameter.Add("@T2Supplier", DbType.String, HttpUtility.HtmlDecode(Item.T2Supplier) ?? string.Empty); }
            if (Item.TestDate != null) { SbSql.Append(",TestDate=@TestDate" + Environment.NewLine); objParameter.Add("@TestDate", DbType.Date, Item.TestDate); }
            if (Item.ReceivedDate != null) { SbSql.Append(",ReceivedDate=@ReceivedDate" + Environment.NewLine); objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate); }
            if (Item.ReleasedDate != null) { SbSql.Append(",ReleasedDate=@ReleasedDate" + Environment.NewLine); objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate); }
            if (Item.Result != null) { SbSql.Append(",Result=@Result" + Environment.NewLine); objParameter.Add("@Result", DbType.String, HttpUtility.HtmlDecode(Item.Result) ?? string.Empty); }
            if (Item.Technician != null) { SbSql.Append(",Technician=@Technician" + Environment.NewLine); objParameter.Add("@Technician", DbType.String, HttpUtility.HtmlDecode(Item.Technician) ?? string.Empty); }
            if (Item.MR != null) { SbSql.Append(",MR=@MR" + Environment.NewLine); objParameter.Add("@MR", DbType.String, HttpUtility.HtmlDecode(Item.MR) ?? string.Empty); }

            SbSql.Append($@"
,OtherMethod=@OtherMethod
,MethodID=iif(@OtherMethod = 0, '', @MethodID)
,TestingMethod=iif(@OtherMethod = 0, @TestingMethod, '')
");
            objParameter.Add("@OtherMethod", DbType.Boolean, Item.OtherMethod);
            objParameter.Add("@MethodID", DbType.String, HttpUtility.HtmlDecode(Item.MethodID) ?? string.Empty);
            objParameter.Add("@TestingMethod", DbType.String, HttpUtility.HtmlDecode(Item.TestingMethod) ?? string.Empty);

            if (Item.HTPlate != null) { SbSql.Append(",HTPlate=@HTPlate" + Environment.NewLine); objParameter.Add("@HTPlate", DbType.Int32, Item.HTPlate); }
            if (Item.HTFlim != null) { SbSql.Append(",HTFlim=@HTFlim" + Environment.NewLine); objParameter.Add("@HTFlim", DbType.Int32, Item.HTFlim); }
            if (Item.HTTime != null) { SbSql.Append(",HTTime=@HTTime" + Environment.NewLine); objParameter.Add("@HTTime", DbType.Int32, Item.HTTime); }
            if (Item.HTPressure != null) { SbSql.Append(",HTPressure=@HTPressure" + Environment.NewLine); objParameter.Add("@HTPressure", DbType.Decimal, Item.HTPressure); }
            if (Item.HTPellOff != null) { SbSql.Append(",HTPellOff=@HTPellOff" + Environment.NewLine); objParameter.Add("@HTPellOff", DbType.String, HttpUtility.HtmlDecode(Item.HTPellOff) ?? string.Empty); }
            if (Item.HT2ndPressnoreverse != null) { SbSql.Append(",HT2ndPressnoreverse=@HT2ndPressnoreverse" + Environment.NewLine); objParameter.Add("@HT2ndPressnoreverse", DbType.Int32, Item.HT2ndPressnoreverse); }
            if (Item.HT2ndPressreversed != null) { SbSql.Append(",HT2ndPressreversed=@HT2ndPressreversed" + Environment.NewLine); objParameter.Add("@HT2ndPressreversed", DbType.Int32, Item.HT2ndPressreversed); }
            if (Item.HTCoolingTime != null) { SbSql.Append(",HTCoolingTime=@HTCoolingTime" + Environment.NewLine); objParameter.Add("@HTCoolingTime", DbType.String, HttpUtility.HtmlDecode(Item.HTCoolingTime) ?? string.Empty); }

            // 圖檔
            SbSql.Append(@"
,TestBeforePicture=@TestBeforePicture
,TestAfterPicture=@TestAfterPicture
");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (Item.Type != null) { SbSql.Append(",Type=@Type" + Environment.NewLine); objParameter.Add("@Type", DbType.String, HttpUtility.HtmlDecode(Item.Type) ?? string.Empty); }
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, HttpUtility.HtmlDecode(Item.ReportNo));

            _MockupWashDetailProvider = new MockupWashDetailProvider(Common.ProductionDataAccessLayer);
            var oldWashData = _MockupWashDetailProvider.GetMockupWash_Detail(new MockupWash_Detail() { ReportNo = Item.ReportNo }).ToList();

            List<MockupWash_Detail_ViewModel> needUpdateDetailList =
                PublicClass.CompareListValue<MockupWash_Detail_ViewModel>(
                    Item.MockupWash_Detail,
                    oldWashData,
                    "Ukey",
                    "TypeofPrint,Design,ArtworkColor,AccessoryRefno,FabricRefNo,FabricColor,Result,Remark");

            string insertDetail = @"
INSERT INTO [dbo].[MockupWash_Detail]
           (ReportNo
           ,TypeofPrint
           ,Design
           ,ArtworkColor
           ,AccessoryRefno
           ,FabricRefNo
           ,FabricColor
           ,Result
           ,Remark
           ,EditName
           ,EditDate
)
     VALUES
           (@ReportNo
           ,@TypeofPrint
           ,@Design
           ,@ArtworkColor
           ,@AccessoryRefno
           ,@FabricRefNo
           ,@FabricColor
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE()
)
";
            string deleteDetail = @"
delete MockupWash_Detail where Ukey = @Ukey
";
            string updateDetail = @"
UPDATE [dbo].[MockupWash_Detail]
   SET 
       [TypeofPrint] =    @TypeofPrint
      ,[Design] =         @Design
      ,[ArtworkColor] =   @ArtworkColor
      ,[AccessoryRefno] = @AccessoryRefno
      ,[FabricRefNo] =    @FabricRefNo
      ,[FabricColor] =    @FabricColor
      ,[Result] =         @Result
      ,[Remark] =         @Remark
      ,[EditName] =       @EditName
      ,[EditDate] = GETDATE()
WHERE UKey = @Ukey
";

            DataTable dtResult = ExecuteDataTable(CommandType.Text, SbSql.ToString(), objParameter);
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                switch (detailItem.StateType)
                {
                    case DatabaseObject.Public.CompareStateType.Add:
                        listDetailPar.Add("@ReportNo", DbType.String, Item.ReportNo);
                        listDetailPar.Add("@TypeofPrint", DbType.String, HttpUtility.HtmlDecode(detailItem.TypeofPrint) ?? string.Empty);
                        listDetailPar.Add("@Design", DbType.String, HttpUtility.HtmlDecode(detailItem.Design) ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, HttpUtility.HtmlDecode(detailItem.ArtworkColor) ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricRefNo) ?? string.Empty);
                        listDetailPar.Add("@AccessoryRefno", DbType.String, HttpUtility.HtmlDecode(detailItem.AccessoryRefno) ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricColor) ?? string.Empty);
                        listDetailPar.Add("@Result", DbType.String, HttpUtility.HtmlDecode(detailItem.Result) ?? string.Empty);
                        listDetailPar.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(detailItem.Remark) ?? string.Empty);
                        listDetailPar.Add("@EditName", DbType.String, HttpUtility.HtmlDecode(detailItem.EditName) ?? string.Empty);

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add("@TypeofPrint", DbType.String, HttpUtility.HtmlDecode(detailItem.TypeofPrint) ?? string.Empty);
                        listDetailPar.Add("@Design", DbType.String, HttpUtility.HtmlDecode(detailItem.Design) ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, HttpUtility.HtmlDecode(detailItem.ArtworkColor) ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricRefNo) ?? string.Empty);
                        listDetailPar.Add("@AccessoryRefno", DbType.String, HttpUtility.HtmlDecode(detailItem.AccessoryRefno) ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricColor) ?? string.Empty);
                        listDetailPar.Add("@Result", DbType.String, HttpUtility.HtmlDecode(detailItem.Result) ?? string.Empty);
                        listDetailPar.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(detailItem.Remark) ?? string.Empty);
                        listDetailPar.Add("@EditName", DbType.String, HttpUtility.HtmlDecode(detailItem.EditName) ?? string.Empty);
                        listDetailPar.Add("@ukey", detailItem.Ukey);

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Delete:
                        listDetailPar.Add("@ukey", detailItem.Ukey);

                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.None:
                        break;
                    default:
                        break;
                }
            }
        }

        public int Delete(MockupWash Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [MockupWash]" + Environment.NewLine);
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<MockupWash_ViewModel> GetMockupWashReportNoList(MockupWash_Request Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT ReportNo
FROM MockupWash m
");
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            /*
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }
            */

            if (!string.IsNullOrEmpty(Item.BrandID))
            {
                SbSql.Append("And BrandID = @BrandID" + Environment.NewLine);
                objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            }
            if (!string.IsNullOrEmpty(Item.SeasonID))
            {
                SbSql.Append("And SeasonID = @SeasonID" + Environment.NewLine);
                objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            }
            if (!string.IsNullOrEmpty(Item.StyleID))
            {
                SbSql.Append("And StyleID = @StyleID" + Environment.NewLine);
                objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            }
            if (!string.IsNullOrEmpty(Item.Article))
            {
                SbSql.Append("And Article = @Article" + Environment.NewLine);
                objParameter.Add("@Article", DbType.String, Item.Article);
            }
            if (!string.IsNullOrEmpty(Item.Type))
            {
                SbSql.Append("And Type = @Type" + Environment.NewLine);
                objParameter.Add("@Type", DbType.String, Item.Type);
            }

            SbSql.Append("Order by ReportNo");
            return ExecuteList<MockupWash_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<MockupWash_ViewModel> GetMockupWash(MockupWash_Request Item, bool istop1)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string top1 = string.Empty;
            if (istop1)
            {
                top1 = "top 1";
            }
            SbSql.Append($@"
SELECT {top1}
         ReportNo
        ,POID
        ,StyleID
        ,SeasonID
        ,BrandID
        ,Article
        ,ArtworkTypeID
        ,Remark
        ,T1Subcon
		,T1SubconAbb = (select Abb from LocalSupp where ID = T1Subcon)
        ,T2Supplier
		,T2SupplierAbb = (select top 1 Abb from (select Abb from LocalSupp where ID = m.T2Supplier and Junk = 0 union select AbbEN from Supp where ID = m.T2Supplier and Junk = 0)x)
        ,TestDate
        ,ReceivedDate
        ,ReleasedDate
        ,Result
        ,Technician
        ,TechnicianName = Technician_ne.Name
        ,TechnicianExtNo = Technician_ne.ExtNo
        ,MR
		,MRName = MR_ne.Name
		,MRExtNo = MR_ne.Extno
        ,MRMail = MR_ne.EMail
		,LastEditName = iif(EditName <> '', Concat (EditName, '-', EditName.Name, ' ', Format(EditDate,'yyyy/MM/dd HH:mm:ss')), Concat (AddName, '-', AddName.Name, ' ', Format(AddDate,'yyyy/MM/dd HH:mm:ss')))
		,m.OtherMethod
        ,m.MethodID
        ,m.TestingMethod
		,m.HTPlate
		,m.HTPellOff
		,m.HTFlim
		,m.HT2ndPressnoreverse
		,m.HTTime
		,m.HT2ndPressreversed
		,m.HTPressure
		,m.HTCoolingTime
        ,Type
        ,TestBeforePicture
        ,TestAfterPicture
        ,AddDate
        ,AddName
        ,EditDate
        ,EditName
        ,Signature = (select t.Signature from Technician t where t.ID = Technician)
FROM MockupWash m
outer apply (select Name, ExtNo from pass1 p inner join Technician t on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo, EMail from pass1 where id = m.MR) MR_ne
outer apply (select Name from Pass1 where id = m.AddName) AddName
outer apply (select Name from Pass1 where id = m.EditName) EditName
");
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }
            if (!string.IsNullOrEmpty(Item.BrandID))
            {
                SbSql.Append("And BrandID = @BrandID" + Environment.NewLine);
                objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            }
            if (!string.IsNullOrEmpty(Item.SeasonID))
            {
                SbSql.Append("And SeasonID = @SeasonID" + Environment.NewLine);
                objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            }
            if (!string.IsNullOrEmpty(Item.StyleID))
            {
                SbSql.Append("And StyleID = @StyleID" + Environment.NewLine);
                objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            }
            if (!string.IsNullOrEmpty(Item.Article))
            {
                SbSql.Append("And Article = @Article" + Environment.NewLine);
                objParameter.Add("@Article", DbType.String, Item.Article);
            }
            if (!string.IsNullOrEmpty(Item.Type))
            {
                SbSql.Append("And Type = @Type" + Environment.NewLine);
                objParameter.Add("@Type", DbType.String, Item.Type);
            }

            SbSql.Append("Order by ReportNo");
            return ExecuteList<MockupWash_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable GetMockupWashFailMailContentData(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo);

            SbSql.Append($@"
SELECT 
         [Report No] = ReportNo
        ,[SP#] = POID
        ,[Style] = StyleID
        ,[Brand] = BrandID
        ,[Season] = SeasonID
        ,[Article] = Article
        ,[Artwork] = ArtworkTypeID
        ,[Remark] = Remark
        ,[T1 Subcon Name] = Concat(T1Subcon, '-' + (select Abb from LocalSupp where ID = T1Subcon))
        ,[T2 Supplier Name] = Concat(T2Supplier, '-' + (select top 1 Abb from (select Abb from LocalSupp where ID = m.T2Supplier and Junk = 0 union select AbbEN from Supp where ID = m.T2Supplier and Junk = 0)x))
        ,[Test Date] = format(TestDate,'yyyy/MM/dd')
        ,[Received Date] = format(ReceivedDate,'yyyy/MM/dd')
        ,[Released Date] = format(ReleasedDate,'yyyy/MM/dd')
        ,[Result] = Result
        ,[Technician] = Concat(Technician, '-', Technician_ne.Name, ' Ext.', Technician_ne.ExtNo)
        ,[MR] = Concat(MR, '-', MR_ne.Name, ' Ext.', MR_ne.ExtNo)
FROM MockupWash m
outer apply (select Name, ExtNo from pass1 p inner join Technician t on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo from pass1 where id = m.MR) MR_ne
where ReportNo = @ReportNo
");
            return ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
