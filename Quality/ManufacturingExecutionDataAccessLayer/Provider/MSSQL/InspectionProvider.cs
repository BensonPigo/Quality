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
                { "@BrandID", DbType.String, inspection_ViewModel.Brand } ,
            };

            SbSql.Append(
                @"
-- 塞入Location

select  ID, cnt = count(1),row = ROW_NUMBER()over(order by id)into #tmpfrom Production.dbo.Orders awhere PulloutComplete = 0and Category = 's'and OnSiteSample = 0and junk = 0and brandid = @BrandIDand FtyGroup = @FtyGroupand qty > 0
and not exists (		select 1		from Production.dbo.Order_Location b		where a.ID = b.OrderId	)and    exists (			    select 1	    from Production.dbo.Style_Location b	    where a.StyleUkey = b.StyleUkey    )
group by ID


declare @cnt int = (select count(1) from #tmp)
declare @OrderID varchar(16)
declare @Num int = 1;while @Num <= @cntbegin		set @OrderID = (select top 1 id from #tmp where row = @Num)	exec Production.dbo.Ins_OrderLocation @OrderID	set @Num = @Num + 1enddrop table #tmp


-- 撈資料
select  [OrderID] = o.ID
	, o.StyleID
	, oq.Article
	, [Size] = oq.SizeCode
	, [ProductType] = ol.Location
from [Production].[dbo].Orders o with(nolock)
inner join [Production].[dbo].Order_Qty oq with(nolock) on o.ID = oq.ID
inner join [Production].[dbo].Order_Location ol with(nolock) on o.ID = ol.OrderId
left join (	
	select OrderID,Article,Size,Location, InspectionQty = count(1)
	from RFT_Inspection r with(nolock)	
	where Status <> 'Dispose'
	group by OrderID,Article,Size,Location
) r on r.OrderID = o.ID 
    and r.Article = oq.Article
    and r.Size = oq.SizeCode
    and r.Location = ol.Location
outer apply (
	select SizeBalanceQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Article = oq.Article
	and r.Size = oq.SizeCode
	and r.Location = ol.Location
	and Status in ('Pass', 'Fixed')
)r_Size
where r_Size.SizeBalanceQty < oq.Qty
and o.Category = 'S'
and o.OnSiteSample != 1
and o.Junk = 0
and o.PulloutComplete = 0 " + Environment.NewLine);

            if (!string.IsNullOrEmpty(inspection_ViewModel.FactoryID)) { SbSql.Append("and o.FtyGroup = @FtyGroup" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.OrderID)) { SbSql.Append("and o.ID = @ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.StyleID)) { SbSql.Append("and o.StyleID = @StyleID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Article)) { SbSql.Append("and oq.Article = @Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Size)) { SbSql.Append("and oq.SizeCode = @SizeCode" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.ProductType)) { SbSql.Append("and ol.Location = @Location" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Brand)) { SbSql.Append("and o.BrandID = @BrandID" + Environment.NewLine); }

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
                { "@BrandID", DbType.String, inspection_ViewModel.Brand } ,
            };

            SbSql.Append(
                @"
select  [OrderID] = o.ID
	, o.StyleID
	, oq.Article
	, [Size] = oq.SizeCode
	, [ProductType] = ol.Location
	, [ProductTypePMS] = ApparelType.Name
	, [Brand] = o.BrandID
	, [Season] = o.SeasonID
	, [SampleStage] = o.OrderTypeID
	, [OriginalLine] = o.SewLine
	, [SizeQty] = cast(oq.Qty as varchar)
	, [SizeBalanceQty] = cast(oq.Qty - r_Size.SizeBalanceQty as varchar)
	, [OrderQty] = cast(o.qty as varchar)
	, [OrderBalanceQty] = cast(o.qty - r_Order.OrderBalanceQty as varchar)
from [Production].[dbo].Orders o with(nolock)
inner join [Production].[dbo].Style s with(nolock) on o.StyleUkey = s.Ukey
inner join [Production].[dbo].Order_Qty oq with(nolock) on o.ID = oq.ID
inner join [Production].[dbo].Order_Location ol with(nolock) on o.ID = ol.OrderId
outer apply (
	select SizeBalanceQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Article = oq.Article
	and r.Size = oq.SizeCode
	and r.Location = ol.Location
	and Status in ('Pass', 'Fixed')
)r_Size
outer apply (
	select OrderBalanceQty = count(*)
	from RFT_Inspection r with(nolock)
	where r.OrderID = o.ID
	and r.Location = ol.Location
	and Status in ('Pass', 'Fixed')
)r_Order
outer apply(
	select r.Name 
	from [Production].[dbo].Style s with(nolock)
	inner join [Production].[dbo].Reason r with(nolock) on r.ID = s.ApparelType 
        and r.ReasonTypeID = 'Style_Apparel_Type'	
	where s.Ukey = o.StyleUkey
)ApparelType
where 1 = 1
--r_Size.SizeBalanceQty < oq.Qty
and o.Category = 'S'
and o.OnSiteSample != 1
and o.Junk = 0
and o.PulloutComplete = 0 " + Environment.NewLine);

            if (!string.IsNullOrEmpty(inspection_ViewModel.FactoryID)) { SbSql.Append("and o.FtyGroup = @FtyGroup" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.OrderID)) { SbSql.Append("and o.ID = @ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.StyleID)) { SbSql.Append("and o.StyleID = @StyleID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Article)) { SbSql.Append("and oq.Article = @Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Size)) { SbSql.Append("and oq.SizeCode = @SizeCode" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.ProductType)) { SbSql.Append("and ol.Location = @Location" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(inspection_ViewModel.Brand)) { SbSql.Append("and o.BrandID = @BrandID" + Environment.NewLine); }

            return ExecuteList<Inspection_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
