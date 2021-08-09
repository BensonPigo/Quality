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
    public class InspectionProvider : SQLDAL, IInspectionProvider
    {
        #region 底層連線
        public InspectionProvider(string conString) : base(conString) { }
        public InspectionProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<Inspection_ViewModel> GetSelectItemData(Inspection_ViewModel inspection_ViewModel)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FtyGroup", DbType.String, inspection_ViewModel.FactoryID } ,
                { "@ID", DbType.String, inspection_ViewModel.OrderID } ,
                { "@StyleID", DbType.String, inspection_ViewModel.StyleID } ,
                { "@Article", DbType.String, inspection_ViewModel.Article } ,
                { "@SizeCode", DbType.String, inspection_ViewModel.Size } ,
                { "@Location", DbType.String, inspection_ViewModel.ProductType } ,
            };

            SbSql.Append(
                @"
select  [OrderID] = o.ID
	, o.StyleID
	, oq.Article
	, [Size] = oq.SizeCode
	, [ProductType] = ol.Location
from [Production].[dbo].Orders o with(nolock)
inner join [Production].[dbo].Order_Qty oq with(nolock) on o.ID = oq.ID
inner join [Production].[dbo].Order_Location ol with(nolock) on o.ID = ol.OrderId
outer apply (
	select InspectionQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Article = oq.Article
	and r.Size = oq.SizeCode
	and r.Location = ol.Location
	and Status <> 'Dispose'
)r
where r.InspectionQty < oq.Qty
and o.Category = 'S'
and o.OnSiteSample != 1
and o.Junk = 0
and o.PulloutComplete = 1 " + Environment.NewLine);

            if (!string.IsNullOrEmpty(inspection_ViewModel.FactoryID)) { SbSql.Append("and o.FtyGroup = @FtyGroup" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.OrderID)) { SbSql.Append("and o.ID = @ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.StyleID)) { SbSql.Append("and o.StyleID = @StyleID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Article)) { SbSql.Append("and oq.Article = @Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Size)) { SbSql.Append("and oq.SizeCode = @SizeCode" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.ProductType)) { SbSql.Append("and ol.Location = @Location" + Environment.NewLine); }

            return ExecuteList<Inspection_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<Inspection_ViewModel> CheckSelectItemData(Inspection_ViewModel inspection_ViewModel)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FtyGroup", DbType.String, inspection_ViewModel.FactoryID } ,
                { "@ID", DbType.String, inspection_ViewModel.OrderID } ,
                { "@StyleID", DbType.String, inspection_ViewModel.StyleID } ,
                { "@Article", DbType.String, inspection_ViewModel.Article } ,
                { "@SizeCode", DbType.String, inspection_ViewModel.Size } ,
                { "@Location", DbType.String, inspection_ViewModel.ProductType } ,
            };

            SbSql.Append(
                @"
select  [OrderID] = o.ID
	, o.StyleID
	, oq.Article
	, [Size] = oq.SizeCode
	, [ProductType] = ol.Location
	, [ProductTypePMS] = n.ID
	, [Brand] = o.BrandID
	, [Season] = o.SeasonID
	, [SampleStage] = o.OrderTypeID
	, [OriginalLine] = o.SewLine
	, [SizeQty] = cast(oq.Qty as varchar)
	, [SizeBalanceQty] = cast(r_Size.SizeBalanceQty as varchar)
	, [OrderQty] = cast(o.qty as varchar)
	, [OrderBalanceQty] = cast(r_Order.OrderBalanceQty as varchar)
from [Production].[dbo].Orders o with(nolock)
inner join [Production].[dbo].Style s with(nolock) on o.StyleUkey = s.Ukey
inner join [Production].[dbo].Order_Qty oq with(nolock) on o.ID = oq.ID
inner join [Production].[dbo].Order_Location ol with(nolock) on o.ID = ol.OrderId
left join [Production].[dbo].NewCDCode n with(nolock) on n.Classifty = 'ApparelType' and n.TypeName = s.ApparelType
outer apply (
	select SizeBalanceQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Article = oq.Article
	and r.Size = oq.SizeCode
	and r.Location = ol.Location
	and Status in ('Passs', 'Fixed')
)r_Size
outer apply (
	select OrderBalanceQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Location = ol.Location
	and Status in ('Passs', 'Fixed')
)r_Order
where r_Size.SizeBalanceQty < oq.Qty
and o.Category = 'S'
and o.OnSiteSample != 1
and o.Junk = 0
and o.PulloutComplete = 1 " + Environment.NewLine);

            if (!string.IsNullOrEmpty(inspection_ViewModel.FactoryID)) { SbSql.Append("and o.FtyGroup = @FtyGroup" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.OrderID)) { SbSql.Append("and o.ID = @ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.StyleID)) { SbSql.Append("and o.StyleID = @StyleID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Article)) { SbSql.Append("and oq.Article = @Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Size)) { SbSql.Append("and oq.SizeCode = @SizeCode" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.ProductType)) { SbSql.Append("and ol.Location = @Location" + Environment.NewLine); }

            return ExecuteList<Inspection_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
