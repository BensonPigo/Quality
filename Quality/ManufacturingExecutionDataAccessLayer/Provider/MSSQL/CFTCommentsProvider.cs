using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class CFTCommentsProvider : SQLDAL, ICFTCommentsProvider
    {
        #region 底層連線
        public CFTCommentsProvider(string conString) : base(conString) { }
        public CFTCommentsProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<CFTComments_ViewModel> Get_CFT_Orders(CFTComments_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            string where;
            if (Req.OrderID != null || !string.IsNullOrEmpty(Req.OrderID))
            {
                where = @"
and o.id = @OrderID";
            }
            else
            {
                where = @"
and o.StyleID = @StyleID
and o.BrandID = @BrandID
and o.SeasonID = @SeasonID";
            }

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append($@"
select OrderID = o.ID
, o.StyleID
, o.BrandID
, o.SeasonID
, SampleStage = OrderTypeID
from [Production].[dbo].Orders o with(nolock)
where 1=1
and o.junk = 0
and o.Category = 'S'
--and o.OnSiteSample != 1
{where}
");
            return ExecuteList<CFTComments_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<CFTComments_Result> Get_CFT_OrderComments(CFTComments_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            string where;
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                where = @"
and o.id = @OrderID";
            }
            else
            {
                where = @"
and o.StyleID = @StyleID
and o.BrandID = @BrandID
and o.SeasonID = @SeasonID";
            }

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append($@"
select r.OrderID
	,o.OrderTypeID
	, PMS_RFTCommentsID = d.Description
	, r.Comnments
into #tmpBase
from MainServer.[Production].[dbo].Orders o with(nolock)
inner join RFT_OrderComments r WITH(NOLOCK) on r.OrderID = o.id
left join Production.dbo.DropdownList d WITH(NOLOCK) ON d.Type='PMS_RFTComments' AND d.ID = r.PMS_RFTCommentsID
where 1=1
and o.junk = 0
and o.Category = 'S'
--and o.OnSiteSample != 1
{where}

select SampleStage = a.OrderTypeID, CommentsCategory = a.PMS_RFTCommentsID, c.Comnments
from (select distinct OrderTypeID,PMS_RFTCommentsID from #tmpBase) a
outer apply(
	select Comnments = stuff((
		select concat('*', Comnments)
		from #tmpBase b
		where a.OrderTypeID = b.OrderTypeID and a.PMS_RFTCommentsID = b.PMS_RFTCommentsID
		and Comnments <> ''
		for xml path('')
	),1,1,'')
)c
order by  a.OrderTypeID, a.PMS_RFTCommentsID

drop table #tmpBase
");
            return ExecuteList<CFTComments_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable Get_CFT_OrderComments_DataTable(CFTComments_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            string where;
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                where = @"
and o.id = @OrderID";
            }
            else
            {
                where = @"
and o.StyleID = @StyleID
and o.BrandID = @BrandID
and o.SeasonID = @SeasonID";
            }

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append($@"
select r.OrderID,o.OrderTypeID, r.PMS_RFTCommentsID, r.Comnments
into #tmpBase
from [Production].[dbo].Orders o with(nolock)
inner join RFT_OrderComments r on r.OrderID = o.id
where 1=1
and o.junk = 0
and o.Category = 'S'
--and o.OnSiteSample != 1
{where}

select SampleStage = a.OrderTypeID, CommentsCategory = (select Name from [Production].[dbo].DropDownList WITH(NOLOCK) where id = a.PMS_RFTCommentsID and type = 'PMS_RFTComments'), c.Comnments
from (select distinct OrderTypeID,PMS_RFTCommentsID from #tmpBase) a
outer apply(
	select Comnments = stuff((
		select concat(char(10), Comnments)
		from #tmpBase b
		where a.OrderTypeID = b.OrderTypeID and a.PMS_RFTCommentsID = b.PMS_RFTCommentsID
		and Comnments <> ''
		for xml path('')
	),1,1,'')
)c
");
            return ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
