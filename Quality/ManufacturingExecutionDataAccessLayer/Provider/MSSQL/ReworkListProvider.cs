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
        public IList<ReworkList_ViewModel> Get(ReworkList_ViewModel Item)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlcmd = $@"
select  
    [ReworkNo] = ins.ReworkCardNo+'/'+ins.FixType
    ,ins.OrderID
    ,[POID] = o.CustPONO
    ,[Style] = o.StyleID
    ,ins.Size
    ,ins.Article
    ,[Defect] =  defect.code
    ,ins.ID
    ,ins.ReworkCardNo
    ,ins.ReworkCardType
    ,ins.FactoryID
    ,ins.Line
    ,ins.FixType
from RFT_Inspection ins with(nolock) 
left join [dbo].[SciProduction_Orders] o with(nolock) 
on ins.OrderId=o.ID
outer apply(
	select code = (
		select CONCAT(AreaCode,'-',DefectCode,'-',PMS_RFTBACriteriaID,'-',Description,';') 
		from (
				select distinct [Description] = dp.Description
				,AreaCode,DefectCode,PMS_RFTBACriteriaID
				from RFT_Inspection_Detail ins2 with(nolock) 
				left join MainServer.Production.dbo.DropDownList dp with (nolock) on ins2.PMS_RFTRespID = dp.id
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

            return ExecuteList<ReworkList_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<ReworkList_ViewModel> GetReworkListFilter(ReworkList_ViewModel Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID", DbType.String, Item.FactoryID } ,
                { "@Line", DbType.String, Item.Line } ,
                { "@Status", DbType.String, Item.Status } ,
            }; 

            SbSql.Append("SELECT OrderID = 'All', Style = 'All', Article = 'All', Size = 'All'" + Environment.NewLine);
            SbSql.Append("union all" + Environment.NewLine);
            SbSql.Append("SELECT distinct r.OrderID, Style = s.ID, r.Article, r.Size" + Environment.NewLine);
            SbSql.Append("FROM [RFT_Inspection] r" + Environment.NewLine);
            SbSql.Append("left join MainServer.Production.dbo.[Style] s on r.StyleUkey = s.Ukey" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.FactoryID)) { SbSql.Append("And r.FactoryID = @FactoryID" + Environment.NewLine); }

            if (!string.IsNullOrEmpty(Item.Line)) { SbSql.Append("And r.Line = @Line" + Environment.NewLine); }

            if (!string.IsNullOrEmpty(Item.Status)) { SbSql.Append("And r.Status = @Status" + Environment.NewLine); }

            return ExecuteList<ReworkList_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        
        #endregion
    }
}
