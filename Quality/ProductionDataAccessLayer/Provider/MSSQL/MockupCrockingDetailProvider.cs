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
            SbSql.Append("FROM [MockupCrocking_Detail] md WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("inner join MockupCrocking m WITH(NOLOCK) on m.ReportNo = md.ReportNo" + Environment.NewLine);
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


        public int Delete(MockupCrocking_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [MockupCrocking_Detail]" + Environment.NewLine);
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
    }
}
