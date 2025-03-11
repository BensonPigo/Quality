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
    public class MockupCrockingProvider : SQLDAL, IMockupCrockingProvider
    {
        #region 底層連線
        public MockupCrockingProvider(string ConString) : base(ConString) { }
        public MockupCrockingProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public string GetFactoryNameEN(string ReportNo,string FactoryID)
        {
            string factoryNameEN = string.Empty;
            SQLParameterCollection objParameter = new SQLParameterCollection 
            {
                { "@ReportNo", DbType.String, ReportNo } ,
                { "@FactoryID",DbType.String, FactoryID } ,
            };
            string sql = $@"
            SELECT
            o.FactoryID
			INTO #tmp
            FROM MockupCrocking M WITH(NOLOCK)
            INNER JOIN Orders O WITH(NOLOCK) ON O.ID = M.POID
            WHERE M.ReportNo = @ReportNo

            SELECT
            F.NameEN
            FROM Factory F WITH(NOLOCK)
            LEFT JOIN #TMP T WITH(NOLOCK) ON T.FactoryID = F.ID
            WHERE F.ID = IIF((SELECT count(1) from #tmp) > 0 ,T.FactoryID,@FactoryID)";

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sql, objParameter);
            factoryNameEN = dt.Rows[0]["NameEN"].ToString();
            return factoryNameEN;
        }
        #region MockupCrocking CUD

        public int CreateMockupCrocking(MockupCrocking Item, string Mdivision, out string NewReportNo)
        {
            NewReportNo = GetID(Mdivision + "CK", "MockupCrocking", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SET XACT_ABORT ON" + Environment.NewLine);
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
            //SbSql.Append("        ,TestBeforePicture" + Environment.NewLine);
            //SbSql.Append("        ,TestAfterPicture" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
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
            SbSql.Append("        ,@TestDate"); objParameter.Add("@TestDate", DbType.Date, Item.TestDate);
            SbSql.Append("        ,@ReceivedDate"); objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate);
            SbSql.Append("        ,@ReleasedDate"); objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, HttpUtility.HtmlDecode(Item.Result) ?? string.Empty);
            SbSql.Append("        ,@Technician"); objParameter.Add("@Technician", DbType.String, HttpUtility.HtmlDecode(Item.Technician) ?? string.Empty);
            SbSql.Append("        ,@MR"); objParameter.Add("@MR", DbType.String, HttpUtility.HtmlDecode(Item.MR) ?? string.Empty);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, HttpUtility.HtmlDecode(Item.Type) ?? string.Empty);

            // 2022/01/10 PMSFile上線，因此去掉Image寫入Production DB的部分
            //SbSql.Append("        ,@TestBeforePicture"); 
            //SbSql.Append("        ,@TestAfterPicture");

            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("        ,GETDATE()");
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, HttpUtility.HtmlDecode(Item.AddName) ?? string.Empty);
            SbSql.Append(")" + Environment.NewLine);

            SbSql.Append($@"
IF EXISTS(
    SELECT 1 FROM SciPMSFile_MockupCrocking WHERE ReportNo = @ReportNo
)
BEGIN
    UPDATE SciPMSFile_MockupCrocking
    SET TestBeforePicture = @TestBeforePicture , TestAfterPicture = @TestAfterPicture
    WHERE ReportNo = @ReportNo
END
ELSE
BEGIN
    INSERT INTO SciPMSFile_MockupCrocking (ReportNo,TestBeforePicture,TestAfterPicture)
    VALUES(@ReportNo,@TestBeforePicture,@TestAfterPicture)
END
");




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public void UpdateMockupCrocking(MockupCrocking_ViewModel Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append($@"
SET XACT_ABORT ON

-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分

UPDATE [MockupCrocking] SET
    EditDate = GETDATE()
    ,EditName=@EditName
    ,POID=@POID
    ,StyleID=@StyleID
    ,SeasonID=@SeasonID
    ,BrandID=@BrandID
    ,Article=@Article
    ,ArtworkTypeID=@ArtworkTypeID
    ,Remark=@Remark
    ,T1Subcon=@T1Subcon
    ,TestDate=@TestDate
    ,ReceivedDate=@ReceivedDate
    ,ReleasedDate=@ReleasedDate
    ,Result=@Result
    ,Technician=@Technician
    ,MR=@MR

WHERE ReportNo = @ReportNo


if not exists (select 1 from SciPMSFile_MockupCrocking where ReportNo = @ReportNo)
begin
    INSERT INTO SciPMSFile_MockupCrocking (ReportNo,TestBeforePicture,TestAfterPicture)
    VALUES (@ReportNo,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    UPDATE SciPMSFile_MockupCrocking SET
        TestBeforePicture=@TestBeforePicture
        ,TestAfterPicture=@TestAfterPicture
    WHERE ReportNo = @ReportNo
end

" + Environment.NewLine);
            objParameter.Add("@EditName", DbType.String, HttpUtility.HtmlDecode(Item.EditName) ?? string.Empty);
            objParameter.Add("@POID", DbType.String, HttpUtility.HtmlDecode(Item.POID) ?? string.Empty);
            objParameter.Add("@StyleID", DbType.String, HttpUtility.HtmlDecode(Item.StyleID) ?? string.Empty);
            objParameter.Add("@SeasonID", DbType.String, HttpUtility.HtmlDecode(Item.SeasonID) ?? string.Empty);
            objParameter.Add("@BrandID", DbType.String, HttpUtility.HtmlDecode(Item.BrandID) ?? string.Empty);
            objParameter.Add("@Article", DbType.String, HttpUtility.HtmlDecode(Item.Article) ?? string.Empty);
            objParameter.Add("@ArtworkTypeID", DbType.String, HttpUtility.HtmlDecode(Item.ArtworkTypeID) ?? string.Empty);
            objParameter.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(Item.Remark) ?? string.Empty);
            objParameter.Add("@T1Subcon", DbType.String, HttpUtility.HtmlDecode(Item.T1Subcon) ?? string.Empty);
            objParameter.Add("@TestDate", DbType.Date, Item.TestDate);
            objParameter.Add("@ReceivedDate", DbType.Date, Item.ReceivedDate);
            objParameter.Add("@ReleasedDate", DbType.Date, Item.ReleasedDate);
            objParameter.Add("@Result", DbType.String, HttpUtility.HtmlDecode(Item.Result) ?? string.Empty);
            objParameter.Add("@Technician", DbType.String, HttpUtility.HtmlDecode(Item.Technician) ?? string.Empty);
            objParameter.Add("@MR", DbType.String, HttpUtility.HtmlDecode(Item.MR) ?? string.Empty);
            objParameter.Add("@Type", DbType.String, HttpUtility.HtmlDecode(Item.Type) ?? string.Empty);
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);


            //_MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
            var oldCrockingData = this.GetMockupCrocking_Detail(new MockupCrocking_Detail() { ReportNo = Item.ReportNo }).ToList();

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

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                switch (detailItem.StateType)
                {
                    case DatabaseObject.Public.CompareStateType.Add:
                        listDetailPar.Add("@ReportNo", DbType.String, HttpUtility.HtmlDecode(Item.ReportNo));
                        listDetailPar.Add("@Design", DbType.String, HttpUtility.HtmlDecode(detailItem.Design) ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, HttpUtility.HtmlDecode(detailItem.ArtworkColor) ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricRefNo) ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricColor) ?? string.Empty);
                        listDetailPar.Add("@DryScale", DbType.String, HttpUtility.HtmlDecode(detailItem.DryScale) ?? string.Empty);
                        listDetailPar.Add("@WetScale", DbType.String, HttpUtility.HtmlDecode(detailItem.WetScale) ?? string.Empty);
                        listDetailPar.Add("@Result", DbType.String, HttpUtility.HtmlDecode(detailItem.Result) ?? string.Empty);
                        listDetailPar.Add("@Remark", DbType.String, HttpUtility.HtmlDecode(detailItem.Remark) ?? string.Empty);
                        listDetailPar.Add("@EditName", DbType.String, HttpUtility.HtmlDecode(detailItem.EditName) ?? string.Empty);

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add("@Design", DbType.String, HttpUtility.HtmlDecode(detailItem.Design) ?? string.Empty);
                        listDetailPar.Add("@ArtworkColor", DbType.String, HttpUtility.HtmlDecode(detailItem.ArtworkColor) ?? string.Empty);
                        listDetailPar.Add("@FabricRefNo", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricRefNo) ?? string.Empty);
                        listDetailPar.Add("@FabricColor", DbType.String, HttpUtility.HtmlDecode(detailItem.FabricColor) ?? string.Empty);
                        listDetailPar.Add("@DryScale", DbType.String, HttpUtility.HtmlDecode(detailItem.DryScale) ?? string.Empty);
                        listDetailPar.Add("@WetScale", DbType.String, HttpUtility.HtmlDecode(detailItem.WetScale) ?? string.Empty);
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

        public int DeleteMockupCrocking(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SET XACT_ABORT ON" + Environment.NewLine);
            SbSql.Append("DELETE [MockupCrocking]" + Environment.NewLine);
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);

            SbSql.Append(@"DELETE SciPMSFile_MockupCrocking" + Environment.NewLine);
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);

            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        #region MockupCrocking_Detail CD

        public int CreateDetail(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [MockupCrocking_Detail]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ReportNo" + Environment.NewLine);
            SbSql.Append("        ,Design" + Environment.NewLine);
            SbSql.Append("        ,ArtworkColor" + Environment.NewLine);
            SbSql.Append("        ,FabricRefNo" + Environment.NewLine);
            SbSql.Append("        ,FabricColor" + Environment.NewLine);
            SbSql.Append("        ,DryScale" + Environment.NewLine);
            SbSql.Append("        ,WetScale" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ReportNo"); objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            SbSql.Append("        ,@Design"); objParameter.Add("@Design", DbType.String, Item.Design ?? string.Empty);
            SbSql.Append("        ,@ArtworkColor"); objParameter.Add("@ArtworkColor", DbType.String, Item.ArtworkColor ?? string.Empty);
            SbSql.Append("        ,@FabricRefNo"); objParameter.Add("@FabricRefNo", DbType.String, Item.FabricRefNo ?? string.Empty);
            SbSql.Append("        ,@FabricColor"); objParameter.Add("@FabricColor", DbType.String, Item.FabricColor ?? string.Empty);
            SbSql.Append("        ,@DryScale"); objParameter.Add("@DryScale", DbType.String, Item.DryScale ?? string.Empty);
            SbSql.Append("        ,@WetScale"); objParameter.Add("@WetScale", DbType.String, Item.WetScale ?? string.Empty);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result ?? string.Empty);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark ?? string.Empty);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName ?? string.Empty);
            SbSql.Append("        ,GETDATE()");
            SbSql.Append(")" + Environment.NewLine);


            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int DeleteDetail(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [MockupCrocking_Detail]" + Environment.NewLine);
            if (string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("Where Ukey = @Ukey" + Environment.NewLine);
                objParameter.Add("@Ukey", Item.Ukey);
            }
            else
            {
                SbSql.Append("Where ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        public IList<MockupCrocking_ViewModel> GetMockupCrockingReportNoList(MockupCrocking_Request Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT ReportNo
FROM [MockupCrocking] m WITH(NOLOCK)
");
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

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
          m.ReportNo
        ,POID
        ,StyleID
        ,SeasonID
        ,BrandID
        ,Article
        ,ArtworkTypeID
        ,Remark
        ,T1Subcon
		,T1SubconAbb = (select Abb from LocalSupp WITH(NOLOCK) where ID = T1Subcon)
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
        ,Type
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_MockupCrocking mi WITH(NOLOCK) where m.ReportNo=mi.ReportNo)
        ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_MockupCrocking mi WITH(NOLOCK) where m.ReportNo=mi.ReportNo)
        ,AddDate
        ,AddName
        ,EditDate
        ,EditName
        ,Signature = (select t.Signature from Technician t WITH(NOLOCK) where t.ID = Technician)
		,LastEditName = iif(EditName <> '', Concat (EditName, '-', EditName.Name, ' ', Format(EditDate,'yyyy/MM/dd HH:mm:ss')), Concat (AddName, '-', AddName.Name, ' ', Format(AddDate,'yyyy/MM/dd HH:mm:ss')))
FROM [MockupCrocking] m WITH(NOLOCK)
outer apply (select Name, ExtNo from pass1 p WITH(NOLOCK) inner join Technician t WITH(NOLOCK) on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo, EMail from pass1 WITH(NOLOCK) where id = m.MR) MR_ne
outer apply (select Name from Pass1 WITH(NOLOCK) where id = m.AddName) AddName
outer apply (select Name from Pass1 WITH(NOLOCK) where id = m.EditName) EditName
");
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And m.ReportNo = @ReportNo" + Environment.NewLine);
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

        public IList<MockupCrocking_Detail_ViewModel> GetMockupCrocking_Detail(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT
         md.ReportNo
        ,md.Ukey
        ,md.Design
        ,md.ArtworkColor
        ,md.FabricRefNo
        ,md.FabricColor
        ,md.DryScale
        ,md.WetScale
        ,md.Result
        ,md.Remark
        ,md.EditName
        ,md.EditDate
		,ArtworkColorName = (select stuff((select concat(';', Name) from Color WITH(NOLOCK) where ID in (select Data from SplitString(md.ArtworkColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
        ,FabricColorName = (select stuff((select concat(';', Name) from Color WITH(NOLOCK) where ID in (select Data from SplitString(md.FabricColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
		,LastUpdate = iif(isnull(md.EditName, '') <> '', Concat(md.EditName, '-' + Format(md.EditDate,'yyyy/MM/dd HH:mm:ss')), Format(md.EditDate,'yyyy/MM/dd HH:mm:ss'))
FROM [MockupCrocking_Detail] md WITH(NOLOCK)
inner join MockupCrocking m WITH(NOLOCK) on m.ReportNo = md.ReportNo
Where 1=1
");
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And md.ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupCrocking_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable GetMockupCrockingFailMailContentData(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo);

            SbSql.Append($@"
SELECT 
         [Report No] = m.ReportNo
        ,[SP#] = POID
        ,[Style] = StyleID
        ,[Brand] = BrandID
        ,[Season] = SeasonID
        ,[Article] = Article
        ,[Artwork] = ArtworkTypeID
        ,[Remark] = Remark
        ,[T1 Subcon Name] = Concat(T1Subcon, '-' + (select Abb from LocalSupp WITH(NOLOCK) where ID = T1Subcon))
        ,[Test Date] = format(TestDate,'yyyy/MM/dd')
        ,[Received Date] = format(ReceivedDate,'yyyy/MM/dd')
        ,[Released Date] = format(ReleasedDate,'yyyy/MM/dd')
        ,[Result] = Result
        ,[Technician] = Concat(Technician, '-', Technician_ne.Name, ' Ext.', Technician_ne.ExtNo)
        ,[MR] = Concat(MR, '-', MR_ne.Name, ' Ext.', MR_ne.ExtNo)
FROM MockupCrocking m WITH(NOLOCK)
outer apply (select Name, ExtNo from pass1 p WITH(NOLOCK) inner join Technician t WITH(NOLOCK) on t.ID = p.ID where t.id = m.Technician) Technician_ne
outer apply (select Name, ExtNo from pass1 WITH(NOLOCK) where id = m.MR) MR_ne
where m.ReportNo = @ReportNo
");
            return ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
