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


namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class BACriteriaProvider : SQLDAL, IBACriteriaProvider
    {
        #region 底層連線
        public BACriteriaProvider(string ConString) : base(ConString) { }
        public BACriteriaProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<BACriteria_Result> Get_BACriteria_Result(string OrderID, string StyleID, string BrandID, string SeasonID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
select OrderID=o.ID
	,o.OrderTypeID
	,o.Qty
	,InspectedQty = ISNULL(Inspected.Qty, 0)
	,BAProduct = ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0)
	,BACriteria = CAST( IIF( ISNULL(Inspected.Qty, 0) = 0 , 0 ,(ISNULL(Inspected.Qty, 0) - ISNULL(BAProduct.Qty, 0) ) *1.0 / ISNULL(Inspected.Qty, 0)*1.0 * 5) as INT)
from SciProduction_Orders o
OUTER APPLY(
	select Qty = COUNT(1)
	from RFT_Inspection i
	where i.OrderID = o.ID
)Inspected
OUTER APPLY(
	select Qty =  COUNT(DISTINCT i.ID)
	from RFT_Inspection i
	inner join RFT_Inspection_Detail id on i.ID = id.ID
	WHERE id.PMS_RFTBACriteriaID BETWEEN 'C2' AND 'C9'
	AND i.OrderID = o.ID
)BAProduct

where o.Junk = 0
AND o.Category = 'S'
AND o.OnSiteSample != 1
");
            if (!string.IsNullOrEmpty(OrderID))
            {
                SbSql.Append($@" AND o.ID = @OrderID" + Environment.NewLine);
                paras.Add("@OrderID", DbType.String, OrderID);
            }
            if (!string.IsNullOrEmpty(SeasonID))
            {
                SbSql.Append($@" AND o.SeasonID = @SeasonID" + Environment.NewLine);
                paras.Add("@SeasonID", DbType.String, SeasonID);
            }
            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@" AND o.BrandID = @BrandID" + Environment.NewLine);
                paras.Add("@BrandID", DbType.String, BrandID);
            }
            if (!string.IsNullOrEmpty(StyleID))
            {
                SbSql.Append($@" AND o.StyleID = @StyleID" + Environment.NewLine);
                paras.Add("@StyleID", DbType.String, StyleID);
            }

            return ExecuteList<BACriteria_Result>(CommandType.Text, SbSql.ToString(), paras);
        }
    }
}
