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

        public IList<CFTComments_where> GetCFT_Orders(CFTComments_where CFTComments)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, CFTComments.OrderID);
            objParameter.Add("@StyleID", DbType.String, CFTComments.StyleID);
            objParameter.Add("@BrandID", DbType.String, CFTComments.BrandID);
            objParameter.Add("@SeasonID", DbType.String, CFTComments.SeasonID);
            string where;
            if (!string.IsNullOrEmpty(CFTComments.OrderID))
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
select OrderID = o.ID, o.StyleID, o.BrandID, o.SeasonID
from [Production].[dbo].Orders o with(nolock)
where 1=1
and o.junk = 0
and o.Category = 'S'
and o.OnSiteSample != 1
{where}
");
            return ExecuteList<CFTComments_where>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<CFTComments_ViewModel> GetRFT_OrderComments(CFTComments_where CFTComments)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, CFTComments.OrderID);
            objParameter.Add("@StyleID", DbType.String, CFTComments.StyleID);
            objParameter.Add("@BrandID", DbType.String, CFTComments.BrandID);
            objParameter.Add("@SeasonID", DbType.String, CFTComments.SeasonID);
            string where;
            if (!string.IsNullOrEmpty(CFTComments.OrderID))
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
and o.OnSiteSample != 1
{where}

select SampleStage = a.OrderTypeID, CommentsCategory = a.PMS_RFTCommentsID, c.Comnments
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
            return ExecuteList<CFTComments_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
