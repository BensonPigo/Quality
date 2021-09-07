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

    public class MockupCrockingDetailProvider : SQLDAL, IMockupCrockingDetailProvider
    {
        #region 底層連線
        public MockupCrockingDetailProvider(string ConString) : base(ConString) { }
        public MockupCrockingDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<MockupCrocking_Detail> Get(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         md.ReportNo" + Environment.NewLine);
            SbSql.Append("        ,md.Ukey" + Environment.NewLine);
            SbSql.Append("        ,md.Design" + Environment.NewLine);
            SbSql.Append("        ,md.ArtworkColor" + Environment.NewLine);
            SbSql.Append("        ,md.FabricRefNo" + Environment.NewLine);
            SbSql.Append("        ,md.FabricColor" + Environment.NewLine);
            SbSql.Append("        ,md.DryScale" + Environment.NewLine);
            SbSql.Append("        ,md.WetScale" + Environment.NewLine);
            SbSql.Append("        ,md.Result" + Environment.NewLine);
            SbSql.Append("        ,md.Remark" + Environment.NewLine);
            SbSql.Append("        ,md.EditName" + Environment.NewLine);
            SbSql.Append("        ,md.EditDate" + Environment.NewLine);
            SbSql.Append("FROM [MockupCrocking_Detail] md" + Environment.NewLine);
            SbSql.Append("inner join MockupCrocking m on m.ReportNo = md.ReportNo" + Environment.NewLine);
            SbSql.Append("Where 1=1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And md.ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupCrocking_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(MockupCrocking_Detail Item)
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
            SbSql.Append("        ,@Design"); objParameter.Add("@Design", DbType.String, Item.Design);
            SbSql.Append("        ,@ArtworkColor"); objParameter.Add("@ArtworkColor", DbType.String, Item.ArtworkColor);
            SbSql.Append("        ,@FabricRefNo"); objParameter.Add("@FabricRefNo", DbType.String, Item.FabricRefNo);
            SbSql.Append("        ,@FabricColor"); objParameter.Add("@FabricColor", DbType.String, Item.FabricColor);
            SbSql.Append("        ,@DryScale"); objParameter.Add("@DryScale", DbType.String, Item.DryScale);
            SbSql.Append("        ,@WetScale"); objParameter.Add("@WetScale", DbType.String, Item.WetScale);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,GETDATE()");
            SbSql.Append(")" + Environment.NewLine);


            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [MockupCrocking_Detail]" + Environment.NewLine);
            SbSql.Append("SET EditDate=GETDATE()" + Environment.NewLine);
            if (Item.Design != null) { SbSql.Append(",Design=@Design" + Environment.NewLine); objParameter.Add("@Design", DbType.String, Item.Design); }
            if (Item.ArtworkColor != null) { SbSql.Append(",ArtworkColor=@ArtworkColor" + Environment.NewLine); objParameter.Add("@ArtworkColor", DbType.String, Item.ArtworkColor); }
            if (Item.FabricRefNo != null) { SbSql.Append(",FabricRefNo=@FabricRefNo" + Environment.NewLine); objParameter.Add("@FabricRefNo", DbType.String, Item.FabricRefNo); }
            if (Item.FabricColor != null) { SbSql.Append(",FabricColor=@FabricColor" + Environment.NewLine); objParameter.Add("@FabricColor", DbType.String, Item.FabricColor); }
            if (Item.DryScale != null) { SbSql.Append(",DryScale=@DryScale" + Environment.NewLine); objParameter.Add("@DryScale", DbType.String, Item.DryScale); }
            if (Item.WetScale != null) { SbSql.Append(",WetScale=@WetScale" + Environment.NewLine); objParameter.Add("@WetScale", DbType.String, Item.WetScale); }
            if (Item.Result != null) { SbSql.Append(",Result=@Result" + Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result); }
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark" + Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark); }
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName" + Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName); }
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
            SbSql.Append("And Ukey = @Ukey" + Environment.NewLine);
            objParameter.Add("@Ukey", DbType.Int64, Item.Ukey);


            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Delete(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [MockupCrocking_Detail]" + Environment.NewLine);
            SbSql.Append("Where Ukey = @Ukey" + Environment.NewLine);
            objParameter.Add("@Ukey", DbType.Int64, Item.Ukey);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

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
		,ArtworkColorName = (select stuff((select concat(';', Name) from Color where ID in (select Data from SplitString(md.ArtworkColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
        ,FabricColorName = (select stuff((select concat(';', Name) from Color where ID in (select Data from SplitString(md.FabricColor,';')) and BrandID = m.BrandID for xml path('')),1,1,''))
		,LastUpdate = Concat (md.EditName, '-', Format(md.EditDate,'yyyy/MM/dd HH:mm:ss'))
FROM [MockupCrocking_Detail] md
inner join MockupCrocking m on m.ReportNo = md.ReportNo
Where 1=1
");
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And md.ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupCrocking_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
