using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class MockupWashDetailProvider : SQLDAL, IMockupWashDetailProvider
    {
        #region 底層連線
        public MockupWashDetailProvider(string ConString) : base(ConString) { }
        public MockupWashDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<MockupWash_Detail> Get(MockupWash_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ReportNo" + Environment.NewLine);
            SbSql.Append("        ,Ukey" + Environment.NewLine);
            SbSql.Append("        ,TypeofPrint" + Environment.NewLine);
            SbSql.Append("        ,Design" + Environment.NewLine);
            SbSql.Append("        ,ArtworkColor" + Environment.NewLine);
            SbSql.Append("        ,FabricRefNo" + Environment.NewLine);
            SbSql.Append("        ,FabricColor" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("FROM [MockupWash_Detail] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("Where 1=1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupWash_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(MockupWash_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [MockupWash_Detail]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ReportNo" + Environment.NewLine);
            SbSql.Append("        ,TypeofPrint" + Environment.NewLine);
            SbSql.Append("        ,Design" + Environment.NewLine);
            SbSql.Append("        ,ArtworkColor" + Environment.NewLine);
            SbSql.Append("        ,FabricRefNo" + Environment.NewLine);
            SbSql.Append("        ,AccessoryRefno" + Environment.NewLine);
            SbSql.Append("        ,FabricColor" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ReportNo"); objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            SbSql.Append("        ,@TypeofPrint"); objParameter.Add("@TypeofPrint", DbType.String, Item.TypeofPrint ?? string.Empty);
            SbSql.Append("        ,@Design"); objParameter.Add("@Design", DbType.String, Item.Design ?? string.Empty);
            SbSql.Append("        ,@ArtworkColor"); objParameter.Add("@ArtworkColor", DbType.String, Item.ArtworkColor ?? string.Empty);
            SbSql.Append("        ,@FabricRefNo"); objParameter.Add("@FabricRefNo", DbType.String, Item.FabricRefNo ?? string.Empty);
            SbSql.Append("        ,@AccessoryRefno"); objParameter.Add("@AccessoryRefno", DbType.String, Item.AccessoryRefno ?? string.Empty);
            SbSql.Append("        ,@FabricColor"); objParameter.Add("@FabricColor", DbType.String, Item.FabricColor ?? string.Empty);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result ?? string.Empty);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark ?? string.Empty);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName ?? string.Empty);
            SbSql.Append("        ,GETDATE()");
            SbSql.Append(")" + Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update(MockupWash_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [MockupWash_Detail]" + Environment.NewLine);
            SbSql.Append("SET EditDate=GETDATE()" + Environment.NewLine);
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName" + Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName); }
            if (Item.TypeofPrint != null) { SbSql.Append(",TypeofPrint=@TypeofPrint" + Environment.NewLine); objParameter.Add("@TypeofPrint", DbType.String, Item.TypeofPrint); }
            if (Item.Design != null) { SbSql.Append(",Design=@Design" + Environment.NewLine); objParameter.Add("@Design", DbType.String, Item.Design); }
            if (Item.ArtworkColor != null) { SbSql.Append(",ArtworkColor=@ArtworkColor" + Environment.NewLine); objParameter.Add("@ArtworkColor", DbType.String, Item.ArtworkColor); }
            if (Item.FabricRefNo != null) { SbSql.Append(",FabricRefNo=@FabricRefNo" + Environment.NewLine); objParameter.Add("@FabricRefNo", DbType.String, Item.FabricRefNo); }
            if (Item.FabricColor != null) { SbSql.Append(",FabricColor=@FabricColor" + Environment.NewLine); objParameter.Add("@FabricColor", DbType.String, Item.FabricColor); }
            if (Item.Result != null) { SbSql.Append(",Result=@Result" + Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result); }
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark" + Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark); }
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
            SbSql.Append("And Ukey = @Ukey" + Environment.NewLine);
            objParameter.Add("@Ukey", DbType.Int64, Item.Ukey);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Delete(MockupWash_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [MockupWash_Detail]" + Environment.NewLine);
            if (string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("Where Ukey = @Ukey" + Environment.NewLine);
                objParameter.Add("@Ukey", DbType.Int64, Item.Ukey);
            }
            else
            {
                SbSql.Append("Where ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<MockupWash_Detail_ViewModel> GetMockupWash_Detail(MockupWash_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT
         md.ReportNo
        ,md.Ukey
		,md.TypeofPrint
        ,md.Design
        ,md.ArtworkColor
        ,md.AccessoryRefno
        ,md.FabricRefNo
        ,md.FabricColor
        ,md.Result
        ,md.Remark
        ,md.EditName
        ,md.EditDate
		,ArtworkColorName = (select stuff((select concat(';', Name) from Color WITH(NOLOCK) where ID in (select Data from SplitString(md.ArtworkColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
        ,FabricColorName = (select stuff((select concat(';', Name) from Color WITH(NOLOCK) where ID in (select Data from SplitString(md.FabricColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
		,LastUpdate = iif(isnull(md.EditName, '') <> '', Concat(md.EditName, '-' + Format(md.EditDate,'yyyy/MM/dd HH:mm:ss')), Format(md.EditDate,'yyyy/MM/dd HH:mm:ss'))
FROM MockupWash_Detail md WITH(NOLOCK)
inner join MockupWash m WITH(NOLOCK) on m.ReportNo = md.ReportNo
Where 1=1
");
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And md.ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupWash_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
