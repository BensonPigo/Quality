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
using ToolKit;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class MockupCrockingProvider : SQLDAL, IMockupCrockingProvider
    {
        private IMockupCrockingDetailProvider _MockupCrockingDetailProvider;
        #region 底層連線
        public MockupCrockingProvider(string ConString) : base(ConString) { }
        public MockupCrockingProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<MockupCrocking> Get(MockupCrocking Item)
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
            SbSql.Append("        ,TestDate" + Environment.NewLine);
            SbSql.Append("        ,ReceivedDate" + Environment.NewLine);
            SbSql.Append("        ,ReleasedDate" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Technician" + Environment.NewLine);
            SbSql.Append("        ,MR" + Environment.NewLine);
            SbSql.Append("        ,Type" + Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("FROM [MockupCrocking]" + Environment.NewLine);

            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupCrocking>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(MockupCrocking Item, string Mdivision, out string NewReportNo)
        {
            NewReportNo = GetID(Mdivision + "CK", "MockupCrocking", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [MockupCrocking]" + Environment.NewLine);
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
            SbSql.Append("        ,TestDate" + Environment.NewLine);
            SbSql.Append("        ,ReceivedDate" + Environment.NewLine);
            SbSql.Append("        ,ReleasedDate" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Technician" + Environment.NewLine);
            SbSql.Append("        ,MR" + Environment.NewLine);
            SbSql.Append("        ,Type" + Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ReportNo"); objParameter.Add("@ReportNo", DbType.String, NewReportNo);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID ?? string.Empty);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, Item.StyleID ?? string.Empty);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID ?? string.Empty);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID ?? string.Empty);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article ?? string.Empty);
            SbSql.Append("        ,@ArtworkTypeID"); objParameter.Add("@ArtworkTypeID", DbType.String, Item.ArtworkTypeID ?? string.Empty);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark ?? string.Empty);
            SbSql.Append("        ,@T1Subcon"); objParameter.Add("@T1Subcon", DbType.String, Item.T1Subcon ?? string.Empty);
            SbSql.Append("        ,@TestDate"); objParameter.Add("@TestDate", DbType.Date, Item.TestDate);
            SbSql.Append("        ,@ReceivedDate"); objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate);
            SbSql.Append("        ,@ReleasedDate"); objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result ?? string.Empty);
            SbSql.Append("        ,@Technician"); objParameter.Add("@Technician", DbType.String, Item.Technician ?? string.Empty);
            SbSql.Append("        ,@MR"); objParameter.Add("@MR", DbType.String, Item.MR ?? string.Empty);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type ?? string.Empty);

            SbSql.Append("        ,@TestBeforePicture");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            SbSql.Append("        ,@TestAfterPicture");
            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("        ,GETDATE()");
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName ?? string.Empty);
            SbSql.Append(")" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public void Update(MockupCrocking_ViewModel Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [MockupCrocking]" + Environment.NewLine);
            SbSql.Append("SET EditDate=GETDATE()" + Environment.NewLine);
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName" + Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName); }

            if (Item.POID != null) { SbSql.Append(",POID=@POID" + Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID); }
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID" + Environment.NewLine); objParameter.Add("@StyleID", DbType.String, Item.StyleID); }
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID" + Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID); }
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID" + Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID); }
            if (Item.Article != null) { SbSql.Append(",Article=@Article" + Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article); }
            if (Item.ArtworkTypeID != null) { SbSql.Append(",ArtworkTypeID=@ArtworkTypeID" + Environment.NewLine); objParameter.Add("@ArtworkTypeID", DbType.String, Item.ArtworkTypeID); }
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark" + Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark); }
            if (Item.T1Subcon != null) { SbSql.Append(",T1Subcon=@T1Subcon" + Environment.NewLine); objParameter.Add("@T1Subcon", DbType.String, Item.T1Subcon); }
            if (Item.TestDate != null) { SbSql.Append(",TestDate=@TestDate" + Environment.NewLine); objParameter.Add("@TestDate", DbType.Date, Item.TestDate); }
            if (Item.ReceivedDate != null) { SbSql.Append(",ReceivedDate=@ReceivedDate" + Environment.NewLine); objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate); }
            if (Item.ReleasedDate != null) { SbSql.Append(",ReleasedDate=@ReleasedDate" + Environment.NewLine); objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate); }
            if (Item.Result != null) { SbSql.Append(",Result=@Result" + Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result); }
            if (Item.Technician != null) { SbSql.Append(",Technician=@Technician" + Environment.NewLine); objParameter.Add("@Technician", DbType.String, Item.Technician); }
            if (Item.MR != null) { SbSql.Append(",MR=@MR" + Environment.NewLine); objParameter.Add("@MR", DbType.String, Item.MR); }
            if (Item.Type != null) { SbSql.Append(",Type=@Type" + Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type); }

            // 圖檔
            SbSql.Append(@"
,TestBeforePicture=@TestBeforePicture
,TestAfterPicture=@TestAfterPicture
");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);


            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
            var oldCrockingData = _MockupCrockingDetailProvider.GetMockupCrocking_Detail(new MockupCrocking_Detail() { ReportNo = Item.ReportNo }).ToList();

            List<MockupCrocking_Detail_ViewModel> needUpdateDetailList =
                PublicClass.CompareListValue<MockupCrocking_Detail_ViewModel>(
                    Item.MockupCrocking_Detail,
                    oldCrockingData,
                    "Ukey",
                    "Design,ArtworkColor,FabricRefNo,FabricColor,DryScale,WetScale,Result,Remark");

            string insertDetail = @"
INSERT INTO [dbo].[MockupCrocking_Detail]
           (ReportNo
           ,Design
           ,ArtworkColor
           ,FabricRefNo
           ,FabricColor
           ,DryScale
           ,WetScale
           ,Result
           ,Remark
           ,EditName
           ,EditDate
		   )
     VALUES
           (@ReportNo
           ,@Design
           ,@ArtworkColor
           ,@FabricRefNo
           ,@FabricColor
           ,@DryScale
           ,@WetScale
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE()
		   )
";
            string deleteDetail = @"
delete MockupCrocking_Detail where Ukey = @Ukey
";
            string updateDetail = @"
UPDATE [dbo].[MockupCrocking_Detail]
   SET [Design] =       @Design
      ,[ArtworkColor] = @ArtworkColor
      ,[FabricRefNo] =  @FabricRefNo
      ,[FabricColor] =  @FabricColor
      ,[DryScale] =     @DryScale
      ,[WetScale] =     @WetScale
      ,[Result] =       @Result
      ,[Remark] =       @Remark
      ,[EditName] =     @EditName
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
                        listDetailPar.Add("@Design", DbType.String, detailItem.Design ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, detailItem.ArtworkColor ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, detailItem.FabricRefNo ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, detailItem.FabricColor ?? string.Empty);
                        listDetailPar.Add("@DryScale", DbType.String, detailItem.DryScale ?? string.Empty);
                        listDetailPar.Add("@WetScale", DbType.String, detailItem.WetScale ?? string.Empty);
                        listDetailPar.Add("@Result", DbType.String, detailItem.Result ?? string.Empty);
                        listDetailPar.Add("@Remark", DbType.String, detailItem.Remark ?? string.Empty);
                        listDetailPar.Add("@EditName", DbType.String, detailItem.EditName ?? string.Empty);

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add("@Design", DbType.String, detailItem.Design ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, detailItem.ArtworkColor ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, detailItem.FabricRefNo ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, detailItem.FabricColor ?? string.Empty);
                        listDetailPar.Add("@DryScale", DbType.String, detailItem.DryScale ?? string.Empty);
                        listDetailPar.Add("@WetScale", DbType.String, detailItem.WetScale ?? string.Empty);
                        listDetailPar.Add("@Result", DbType.String, detailItem.Result ?? string.Empty);
                        listDetailPar.Add("@Remark", DbType.String, detailItem.Remark ?? string.Empty);
                        listDetailPar.Add("@EditName", DbType.String, detailItem.EditName ?? string.Empty);
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

        public int Delete(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE [MockupCrocking]" + Environment.NewLine);
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        public IList<MockupCrocking_ViewModel> GetMockupCrockingReportNoList(MockupCrocking_Request Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT ReportNo
FROM [MockupCrocking] m
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
            return ExecuteList<MockupCrocking_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<MockupCrocking_ViewModel> GetMockupCrocking(MockupCrocking_Request Item, bool istop1)
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
        ,Type
        ,TestBeforePicture
        ,TestAfterPicture
        ,AddDate
        ,AddName
        ,EditDate
        ,EditName
        ,SignaturePic = (select PicPath from system) + (select t.SignaturePic from Technician t where t.ID = Technician)
		,LastEditName = iif(EditName <> '', Concat (EditName, '-', EditName.Name, ' ', Format(EditDate,'yyyy/MM/dd HH:mm:ss')), Concat (AddName, '-', AddName.Name, ' ', Format(AddDate,'yyyy/MM/dd HH:mm:ss')))
FROM [MockupCrocking] m
outer apply (select Name, ExtNo from pass1 p inner join Technician t on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo from pass1 where id = m.MR) MR_ne
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

            return ExecuteList<MockupCrocking_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable GetMockupCrockingFailMailContentData(string ReportNo)
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
        ,[Test Date] = format(TestDate,'yyyy/MM/dd')
        ,[Received Date] = format(ReceivedDate,'yyyy/MM/dd')
        ,[Released Date] = format(ReleasedDate,'yyyy/MM/dd')
        ,[Result] = Result
        ,[Technician] = Concat(Technician, '-', Technician_ne.Name, ' Ext.', Technician_ne.ExtNo)
        ,[MR] = Concat(MR, '-', MR_ne.Name, ' Ext.', MR_ne.ExtNo)
FROM MockupCrocking m
outer apply (select Name, ExtNo from pass1 p inner join Technician t on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo from pass1 where id = m.MR) MR_ne
where ReportNo = @ReportNo
");
            return ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
