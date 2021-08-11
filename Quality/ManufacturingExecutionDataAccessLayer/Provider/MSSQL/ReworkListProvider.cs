using ADOHelper.Template.MSSQL;
using ManufacturingExecutionDataAccessLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Text;
using System.Data;
using System.Collections.Generic;
using ADOHelper.Utility;
using DatabaseObject.ViewModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class ReworkListProvider : SQLDAL, IReworkListProvider
    {
        #region 底層連線
        public ReworkListProvider(string conString) : base(conString) { }
        public ReworkListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<ReworkList> Get(ReworkList Item)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlcmd = $@"
select distinct 
[ReworkNo] = ins.ReworkCardNo+'/'+ins.FixType
,[SPNo] = ins.OrderId
,[PONo] = o.CustPONo
,o.StyleID
,ins.Size
,ins.Article
,[Defect] =  defect.code
, Reworked = 'Pass'
,[AddDefect] = 'Reject'
,Status='Action'
,ins.ReworkCardNo
,ins.ReworkCardType
,ins.FactoryID
,ins.Line
from RFT_Inspection ins with(nolock) 
left join [dbo].[SciProduction_Orders] o with(nolock) 
on ins.OrderId=o.ID
outer apply(
	select code = (
		select CONCAT(AreaCode,'-',DefectCode,'-',PMS_RFTBACriteriaID,'-',Description) 
		from (
				select distinct [Description] = dp.Description
				,AreaCode,DefectCode,PMS_RFTBACriteriaID
				from RFT_Inspection_Detail ins2 with(nolock) 
				left join Production.dbo.DropDownList dp with (nolock) on ins2.PMS_RFTRespID = dp.id
				and dp.Type = 'PMS_RFTResp'
				where ins2.ID=ins.ID
			)s
		for xml path('')
	)
)defect
where ins.Status='Reject'
and ins.Line='{Item.Line}'
and ins.FactoryID = '{Item.FactoryID}'
order by ReworkCardType desc, ins.AddDate
";

            return ExecuteList<ReworkList>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<ReworkList_ViewModel> GetReworkListFilter(ReworkList_ViewModel Item, string key)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID", DbType.String, Item.FactoryID } ,
                { "@Line", DbType.String, Item.Line } ,
                { "@Status", DbType.DateTime, Item.Status } ,
            };

            switch (key)
            {
                case "OrderID":
                    SbSql.Append("select SP = 'All'" + Environment.NewLine);
                    SbSql.Append("union all" + Environment.NewLine);

                    SbSql.Append("SELECT distinct" + Environment.NewLine);
                    SbSql.Append(" SP = r.OrderID" + Environment.NewLine);
                    break;
                case "StyleID":
                    SbSql.Append("select Style = 'All'" + Environment.NewLine);
                    SbSql.Append("union all" + Environment.NewLine);
                    SbSql.Append("SELECT distinct" + Environment.NewLine);
                    SbSql.Append(" Style = s.ID" + Environment.NewLine);
                    break;
                case "Article":
                    SbSql.Append("select Article = 'All'" + Environment.NewLine);
                    SbSql.Append("union all" + Environment.NewLine);
                    SbSql.Append("SELECT distinct" + Environment.NewLine);
                    SbSql.Append(" r.Article" + Environment.NewLine);
                    break;
                case "Size":
                    SbSql.Append("select Size = 'All'" + Environment.NewLine);
                    SbSql.Append("union all" + Environment.NewLine);
                    SbSql.Append("SELECT distinct" + Environment.NewLine);
                    SbSql.Append(" r.Size" + Environment.NewLine);
                    break;
                default:
                    break;
            }

            SbSql.Append("FROM [RFT_Inspection] r" + Environment.NewLine);
            SbSql.Append("left join Production.dbo.[Style] s on r.StyleUkey = s.Ukey" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.rft_Inspection.FactoryID.ToString())) { SbSql.Append("And r.FactoryID = @FactoryID" + Environment.NewLine); }

            if (!string.IsNullOrEmpty(Item.rft_Inspection.Line)) { SbSql.Append("And r.Line = @Line" + Environment.NewLine); }

            if (!string.IsNullOrEmpty(Item.rft_Inspection.Status.ToString())) { SbSql.Append("And r.Status = @Status" + Environment.NewLine); }

            return ExecuteList<ReworkList_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
