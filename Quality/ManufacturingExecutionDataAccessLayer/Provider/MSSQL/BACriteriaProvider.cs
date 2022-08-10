using ADOHelper.Template.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Utility;
using DatabaseObject.ResultModel;
using System.Web.Mvc;
using System.Data;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class BACriteriaProvider : SQLDAL, IBACriteriaProvider
    {
        #region 底層連線
        public BACriteriaProvider(string ConString) : base(ConString) { }
        public BACriteriaProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public int Check_SampleRFTInspection_Count(BACriteria_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
            };
            string sqlcmd = $@"
Select count(1)
from SampleRFTInspection i WITH(NOLOCK) 
where 1=1
";
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                sqlcmd += $@" AND OrderID = @OrderID " + Environment.NewLine;
                objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            }

            if (!string.IsNullOrEmpty(Req.InspectionDateStart) && !string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                sqlcmd += $@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate between @sDate and @eDate)" + Environment.NewLine;
                objParameter.Add("@sDate", DbType.String, Req.InspectionDateStart);
                objParameter.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateStart))
            {
                sqlcmd += $@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate >= @sDate)" + Environment.NewLine;
                objParameter.Add("@sDate", DbType.String, Req.InspectionDateStart);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                sqlcmd += $@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate <= @eDate)" + Environment.NewLine;
                objParameter.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }

            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToInt32(result == null ? 0 : result);
        }
        public IList<BACriteria_Result> Get_BACriteria_Result(BACriteria_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
select OrderID=o.ID
	,o.OrderTypeID
	,o.Qty
	,InspectedQty = ISNULL(Inspected.Qty, 0)
	,BAProduct = ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0)
	,BACriteria = ROUND( IIF( ISNULL(Inspected.Qty, 0) = 0 , 0 ,(ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0) ) *1.0 / ISNULL(Inspected.Qty, 0)*1.0 * 5) ,1)
from Production.dbo.Orders o WITH(NOLOCK)
OUTER APPLY(
	select Qty = COUNT(1)
	from RFT_Inspection i WITH(NOLOCK)
	where i.OrderID = o.ID
)Inspected
OUTER APPLY(
	select Qty =  COUNT(DISTINCT i.ID)
	from RFT_Inspection i WITH(NOLOCK)
	inner join RFT_Inspection_Detail id WITH(NOLOCK) on i.ID = id.ID
	WHERE id.PMS_RFTBACriteriaID BETWEEN 'C2' AND 'C9'
	AND i.OrderID = o.ID
)BAProduct

where o.Junk = 0
AND o.Category = 'S'
--AND o.OnSiteSample != 1
");
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql.Append($@" AND o.ID = @OrderID" + Environment.NewLine);
                paras.Add("@OrderID", DbType.String, Req.OrderID);
            }

            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql.Append($@" AND o.SeasonID = @SeasonID" + Environment.NewLine);
                paras.Add("@SeasonID", DbType.String, Req.SeasonID);
            }

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql.Append($@" AND o.BrandID = @BrandID" + Environment.NewLine);
                paras.Add("@BrandID", DbType.String, Req.BrandID);
            }

            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql.Append($@" AND o.StyleID = @StyleID" + Environment.NewLine);
                paras.Add("@StyleID", DbType.String, Req.StyleID);
            }

            if (!string.IsNullOrEmpty(Req.InspectionDateStart) && !string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from RFT_Inspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate between @sDate and @eDate)" + Environment.NewLine);
                paras.Add("@sDate", DbType.String, Req.InspectionDateStart);
                paras.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateStart))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from RFT_Inspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate >= @sDate)" + Environment.NewLine);
                paras.Add("@sDate", DbType.String, Req.InspectionDateStart);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from RFT_Inspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate <= @eDate)" + Environment.NewLine);
                paras.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }
            
            return ExecuteList<BACriteria_Result>(CommandType.Text, SbSql.ToString(), paras);
        }
        public IList<BACriteria_Result> Get_BACriteria_Result_SampleRFTInspection(BACriteria_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
select OrderID=o.ID
	,o.OrderTypeID
	,o.Qty
	,InspectedQty =  Inspected.Qty
	,BAProduct = ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0)
	,BACriteria = ROUND( IIF( ISNULL(Inspected.Qty, 0) = 0 
					, 0 ,
					(ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0) ) *1.0 / ISNULL(Inspected.Qty, 0)*1.0 * 5) 
				,1)
from Production.dbo.Orders o WITH(NOLOCK)
OUTER APPLY(
	SELECT Qty = SUM(InspectedQty) FROM (
		select InspectedQty = CASE WHEN s.InspectionStage = '100%' THEN o.Qty
									WHEN s.InspectionStage='AQL' THEN s.SampleSize
								ELSE  o.Qty
							END
		from SampleRFTInspection s WITH(NOLOCK)
		where o.ID=s.OrderID
	)zz
)Inspected
OUTER APPLY(
	select Qty =  SUM(sn.Qty)
	from SampleRFTInspection s WITH(NOLOCK)
	inner join SampleRFTInspection_NonBACriteria sn WITH(NOLOCK) on s.ID =  sn.SampleRFTInspectionUkey and Qty > 0
	WHERE sn.BACriteria BETWEEN 'C2' AND 'C9'
	AND s.OrderID = o.ID
)BAProduct

where o.Junk = 0
AND o.Category = 'S'
--AND o.OnSiteSample != 1
");
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql.Append($@" AND o.ID = @OrderID" + Environment.NewLine);
                paras.Add("@OrderID", DbType.String, Req.OrderID);
            }

            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql.Append($@" AND o.SeasonID = @SeasonID" + Environment.NewLine);
                paras.Add("@SeasonID", DbType.String, Req.SeasonID);
            }

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql.Append($@" AND o.BrandID = @BrandID" + Environment.NewLine);
                paras.Add("@BrandID", DbType.String, Req.BrandID);
            }

            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql.Append($@" AND o.StyleID = @StyleID" + Environment.NewLine);
                paras.Add("@StyleID", DbType.String, Req.StyleID);
            }

            if (!string.IsNullOrEmpty(Req.InspectionDateStart) && !string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate between @sDate and @eDate)" + Environment.NewLine);
                paras.Add("@sDate", DbType.String, Req.InspectionDateStart);
                paras.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateStart))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate >= @sDate)" + Environment.NewLine);
                paras.Add("@sDate", DbType.String, Req.InspectionDateStart);
            }
            else if (!string.IsNullOrEmpty(Req.InspectionDateEnd))
            {
                SbSql.Append($@" AND EXISTS (Select 1 from SampleRFTInspection i WITH(NOLOCK) where i.OrderID = o.ID and i.InspectionDate <= @eDate)" + Environment.NewLine);
                paras.Add("@eDate", DbType.String, Req.InspectionDateEnd);
            }

            return ExecuteList<BACriteria_Result>(CommandType.Text, SbSql.ToString(), paras);
        }
    }
}
