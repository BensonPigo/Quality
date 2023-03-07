using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class SearchListProvider : SQLDAL , ISearchListProvider
    {
        #region 底層連線
        public SearchListProvider(string conString) : base(conString) { }
        public SearchListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SelectListItem> GetTypeDatasource(string Pass1ID)
        {
            string SbSql = $@"
select Text = '', Value = ''
UNION
select Text = IIF(md.FunctionName IS NOt NULL , md.FunctionName,m.FunctionName)
	 , Value = IIF(md.FunctionName IS NOt NULL , md.FunctionName,m.FunctionName)
from Quality_Pass1 p WITH(NOLOCK)
inner join Quality_Position pp WITH(NOLOCK) on p.Position=pp.ID
inner join Quality_Pass2 p2 WITH(NOLOCK) on p2.PositionID=pp.ID 
inner join Quality_Menu m WITH(NOLOCK) on m.ID=p2.MenuID
left join Quality_Menu_detail md WITH(NOLOCK) on md.ID=m.ID AND md.Type=p.BulkFGT_Brand
where ModuleName='Bulk FGT'
and FunctionSeq not in (10,20)
and p.ID='{Pass1ID}'
";
            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, new SQLParameterCollection());
        }

        public IList<SearchList_Result> Get_SearchList(SearchList_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID },
                { "@SeasonID", DbType.String, Req.SeasonID },
                { "@StyleID", DbType.String, Req.StyleID },
                { "@Article", DbType.String, Req.Article },
                { "@MDivisionID", DbType.String, Req.MDivisionID },

                { "@ReceivedDate_s", DbType.Date, Req.ReceivedDate_s },
                { "@ReceivedDate_e", DbType.Date, Req.ReceivedDate_e },
                { "@ReportDate_s", DbType.Date, Req.ReportDate_s },
                { "@ReportDate_e", DbType.Date, Req.ReportDate_e },
                { "@WhseArrival_s", DbType.Date, Req.WhseArrival_s },
                { "@WhseArrival_e", DbType.Date, Req.WhseArrival_e },
            };

            string sqlWhseArrival = string.Empty;
            StringBuilder SbSql = new StringBuilder();

            #region Fabric Crocking & Shrinkage Test (504, 405)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type1 = $@"
select Type = 'Fabric Crocking & Shrinkage Test (504, 405)'
        , ReportNo=''
		, OrderID = o.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, Article = ''
		, Artwork = ''
		, [Result] = f_Result.Result
		, [TestDate] = f_TestDate.TestDate
	    , ReceivedDate = NULL
	    , ReportDate = NULL
        , AddName = '' ----AddName不會是單一個人，故不顯示
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = p.ID
outer apply (
	select TestDate = MAX(f.TestDate)
	from (
		select TestDate = (
				SELECT MAX(TestDate) FROM (
					SELECT TestDate = f.CrockingDate
					UNION
					SELECT TestDate = f.HeatDate
					UNION 
					SELECt TestDate = f.WashDate
				)tmp
			)
		from FIR_Laboratory f
		where f.POID = p.ID
	)f
)f_TestDate
outer apply (
	select Result = case Max(case f.Result 
				 when 'Fail' then 2
				 when 'Pass' then 1
				 else 0
			   end)
			when 2 then 'Fail'
			when 1 then 'Pass'
			else ''
			end
	from FIR_Laboratory f
	where f.POID = p.ID
)f_Result
{sqlWhseArrival}
WHERE exists (select 1 from FIR_Laboratory f WITH(NOLOCK) WHERE f.POID = p.ID)
and f_Result.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type1 += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type1 += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type1 += "AND o.StyleID = @StyleID ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type1 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type1 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Garment Test (450, 451, 701, 710)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Orders o WITH(NOLOCK)
	inner join Export_Detail ed WITH(NOLOCK) on o.POID = ed.PoID
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.ID = g.OrderID
)e ";
            }

            string type2 = $@"
select  Type = 'Garment Test (450, 451, 701, 710)'
        , gd.ReportNo
		, gd.OrderID
		, StyleID
		, BrandID
		, SeasonID
		, Article
		, Artwork = ''
		, Result= IIF(gd.Result='P','Pass', IIF(gd.Result='F','Fail',''))
		, TestDate = gd.InspDate
	    , ReceivedDate = NULL
	    , ReportDate = NULL
        , AddName = ISNULL(pa.Name, ma.Name)
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) ON g.ID= gd.ID
left join Production.dbo.Pass1 pa on gd.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on gd.AddName = ma.ID
{sqlWhseArrival}
WHERE gd.Result <> ''
";

            if (!string.IsNullOrEmpty(Req.MDivisionID))
            {
                type2 += "AND g.MDivisionID = @MDivisionID ";
            }
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type2 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type2 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type2 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type2 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type2 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type2 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Mockup Crocking Test  (504)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Orders o WITH(NOLOCK)
	inner join Export_Detail ed WITH(NOLOCK) on o.POID = ed.PoID
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.ID = m.POID
)e ";
            }

            string type3 = $@"
select DISTINCT  Type = 'Mockup Crocking Test  (504)'
        , m.ReportNo
		, OrderID = m.POID
		, m.StyleID
		, m.BrandID
		, m.SeasonID
		, m.Article
		, Artwork = m.ArtworkTypeID
		, m.Result
		, m.TestDate 
        , m.ReceivedDate
        , ReportDate = m.ReleasedDate
        , AddName = ISNULL(pa.Name, ma.Name)
from MockupCrocking m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
WHERE m.Result <> ''
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type3 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type3 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type3 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type3 += "AND Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                type3 += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type3 += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                type3 += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type3 += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type3 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type3 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Mockup Oven Test (514)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Orders o WITH(NOLOCK)
	inner join Export_Detail ed WITH(NOLOCK) on o.POID = ed.PoID
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.ID = m.POID
)e ";
            }

            string type4 = $@"
select DISTINCT Type = 'Mockup Oven Test (514)'
	, m.ReportNo
	, OrderID = m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = m.ArtworkTypeID
	, m.Result
	, m.TestDate
    , m.ReceivedDate
    , ReportDate = m.ReleasedDate
    , AddName = ISNULL(pa.Name, ma.Name)
from MockupOven m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Type = 'B'
and m.Result <> ''
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type4 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type4 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type4 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type4 += "AND Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                type4 += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type4 += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                type4 += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type4 += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type4 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type4 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Mockup Wash Test (701)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Orders o WITH(NOLOCK)
	inner join Export_Detail ed WITH(NOLOCK) on o.POID = ed.PoID
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.ID = m.POID
)e ";
            }

            string type5 = $@"
select DISTINCT Type = 'Mockup Wash Test (701)'
	, m.ReportNo
	, OrderID = m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = m.ArtworkTypeID
	, m.Result
	, m.TestDate
    , m.ReceivedDate
    , ReportDate = m.ReleasedDate
    , AddName = ISNULL(pa.Name, ma.Name)
from MockupWash m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Type = 'B' 
and m.Result <> ''
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type5 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type5 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type5 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type5 += "AND Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                type5 += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type5 += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                type5 += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type5 += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type5 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type5 += " AND @WhseArrival_e >= e.WhseArrival ";
            }

            #endregion

            #region Fabric Oven Test (515) 
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type6 = $@"
select DISTINCT Type= 'Fabric Oven Test (515)'
        , f.ReportNo
		, OrderID = o.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, f.Article  
		, Artwork = ''
		, Result = f.Result
		, TestDate = f.InspDate
	    , ReceivedDate = NULL
	    , ReportDate = NULL
        , AddName = ISNULL(pa.Name, ma.Name)
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
inner join Oven f WITH(NOLOCK) ON f.POID = p.ID
left join Production.dbo.Pass1 pa on f.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on f.AddName = ma.ID
{sqlWhseArrival}
where f.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type6 += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type6 += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type6 += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type6 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type6 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type6 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Washing Fastness (501)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type7 = $@"
select DISTINCT Type= 'Washing Fastness (501)'
        , ReportNo = f.ID
		, OrderID = o.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, f.Article 
		, Artwork = ''
		, Result = f.Result
		, TestDate = f.InspDate
	    , ReceivedDate = NULL
	    , ReportDate = NULL
        , AddName = ISNULL(pa.Name, ma.Name)
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
INNER JOIN ColorFastness f WITH(NOLOCK) ON f.POID = p.ID
left join Production.dbo.Pass1 pa on f.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on f.AddName = ma.ID
{sqlWhseArrival}
WHERE f.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type7 += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type7 += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type7 += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type7 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type7 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type7 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Accessory Oven & Wash Test (515, 701)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK) 
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type8 = $@"
select Type = 'Accessory Oven & Wash Test (515, 701)'
        , ReportNo = ''
		, OrderID = o.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, Article = ''
		, [Artwork] = ''
		, [Result] = f_Result.Result
		, [TestDate] = f_TestDate.TestDate
		, ReceivedDate = NULL
		, ReportDate = NULL
        , AddName = '' ----不會只有一個人，故空著
from PO p
inner join Orders o WITH(NOLOCK) ON o.ID = p.ID
outer apply (
	select TestDate = MAX(f.TestDate)
	from (
		select TestDate = (
				SELECT MAX(TestDate) FROM (
					SELECT TestDate = f.OvenDate
					UNION 
					SELECt TestDate = f.WashDate
				)tmp
			)
		from AIR_Laboratory f
		where f.POID = p.ID
	)f
)f_TestDate
outer apply (
	select Result = case Max(case f.Result 
				 when 'Fail' then 2
				 when 'Pass' then 1
				 else 0
			   end)
			when 2 then 'Fail'
			when 1 then 'Pass'
			else ''
			end
	from AIR_Laboratory f
	where f.POID = p.ID
)f_Result
{sqlWhseArrival}
WHERE exists (select 1 from AIR_Laboratory f WITH(NOLOCK) WHERE f.POID = p.ID)
and f_Result.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type8 += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type8 += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type8 += "AND o.StyleID = @StyleID ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type8 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type8 += " AND @WhseArrival_e >= e.WhseArrival ";
            }

            #endregion

            #region Pulling test for Snap/Botton/Rivet (437)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Orders o WITH(NOLOCK)
	inner join Export_Detail ed WITH(NOLOCK) on o.POID = ed.PoID
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.ID = m.POID
)e ";
            }

            string type9 = $@"
select DISTINCT Type = 'Pulling test for Snap/Botton/Rivet (437)'
	, m.ReportNo
	, OrderID = m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = ''
	, m.Result
	, m.TestDate
	, ReceivedDate = NULL
	, ReportDate = NULL
    , AddName = ISNULL(pa.Name, ma.Name)
from [ExtendServer].ManufacturingExecution.dbo.PullingTest m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type9 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type9 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID)) 
            {
                type9 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type9 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type9 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type9 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Water Fastness Test(503)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type11 = $@"
select DISTINCT Type= 'Water Fastness Test(503)'
        , w.ReportNo
		, OrderID = w.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, Article 
		, Artwork = ''
		, w.Result
		, TestDate = w.InspDate
	    , ReceivedDate =NULL
	    , ReportDate =NULL
        , AddName = ISNULL(pa.Name, ma.Name)
from WaterFastness w WITH (NOLOCK) 
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
left join Production.dbo.Pass1 pa on w.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on w.AddName = ma.ID
{sqlWhseArrival}
WHERE w.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type11 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type11 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type11 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type11 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type11 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type11 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Perspiration Fastness Test(502)
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type12 = $@"
select DISTINCT Type= 'Perspiration Fastness (502)'
        , w.ReportNo
		, OrderID = w.POID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, Article 
		, Artwork = ''
		, w.Result
		, TestDate = w.InspDate
	    , ReceivedDate = NULL
	    , ReportDate = NULL
        , AddName = ISNULL(pa.Name, ma.Name)
from PerspirationFastness w WITH (NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
left join Production.dbo.Pass1 pa on w.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on w.AddName = ma.ID
{sqlWhseArrival}
WHERE w.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type12 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type12 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type12 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type12 += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type12 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type12 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            #region Daily HT Wash Test
            string type13 = $@"
select DISTINCT Type= 'Daily HT Wash Test'
        , h.ReportNo
		, h. OrderID
		, h.StyleID
		, h.BrandID
		, h.SeasonID
		, h.Article 
		, h.Line 
		, Artwork = ''
		, h.Result
		, TestDate = h.ReportDate
	    , ReceivedDate = h.ReceivedDate
	    , ReportDate = h.ReportDate
        , AddName = ISNULL(pa.Name, ma.Name)
from [ExtendServer].ManufacturingExecution.dbo.HeatTransferWash h WITH (NOLOCK)
left join Production.dbo.Pass1 pa on h.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on h.AddName = ma.ID
WHERE h.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type13 += "AND h.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type13 += "AND h.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type13 += "AND h.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type13 += "AND h.Article = @Article ";
            }
            if (!string.IsNullOrEmpty(Req.Line))
            {
                type13 += "AND h.Line = @Line ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                type13 += " AND @ReceivedDate_s <= h.ReceivedDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type13 += " AND h.ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                type13 += " AND @ReportDate_s <= h.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type13 += " AND h.ReportDate <= @ReportDate_e ";
            }
            #endregion

            #region Daily Bulk Moistur Test
            if (Req.WhseArrival_s.HasValue || Req.WhseArrival_e.HasValue)
            {
                sqlWhseArrival = @" 
outer apply (
	select WhseArrival = MAX(e.WhseArrival)
	from Export_Detail ed WITH(NOLOCK)
	inner join Export e WITH(NOLOCK) on e.ID = ed.ID
	where o.POID = ed.PoID
)e ";
            }

            string type14 = $@"
select DISTINCT Type= 'Daily Bulk Moisture Test'
        , h.ReportNo
		, h. OrderID
		, o.StyleID
		, o.BrandID
		, o.SeasonID
		, h.Article 
		, h.Line
		, Artwork = ''
		, h.Result
		, TestDate = h.AddDate
	    , ReceivedDate = NULL
	    , ReportDate = h.ReportDate
        , AddName = ISNULL(pa.Name, ma.Name)
from [ExtendServer].ManufacturingExecution.dbo.BulkMoistureTest h WITH (NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = h.OrderID
left join Production.dbo.Pass1 pa on h.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on h.AddName = ma.ID
{sqlWhseArrival}
WHERE h.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type14 += "AND h.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type14 += "AND h.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type14 += "AND h.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type14 += "AND h.Article = @Article ";
            }
            if (!string.IsNullOrEmpty(Req.Line))
            {
                type14 += "AND h.Line = @Line ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                type14 += " AND @ReportDate_s <= h.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type14 += " AND h.ReportDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                type14 += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                type14 += " AND @WhseArrival_e >= e.WhseArrival ";
            }
            #endregion

            switch (Req.Type)
            {
                case string a when a.Contains("Fabric Crocking & Shrinkage Test"):
                    SbSql.Append(type1);
                    break;
                case string a when a.Contains("Garment Test"):
                    SbSql.Append(type2);
                    break;
                case string a when a.Contains("Mockup Crocking Test"):
                    SbSql.Append(type3);
                    break;
                case string a when a.Contains("Mockup Oven Test"):
                    SbSql.Append(type4);
                    break;
                case string a when a.Contains("Mockup Wash Test"):
                    SbSql.Append(type5);
                    break;
                case string a when a.Contains("Fabric Oven Test"):
                    SbSql.Append(type6);
                    break;
                case string a when a.Contains("Washing Fastness"):
                    SbSql.Append(type7);
                    break;
                case string a when a.Contains("Accessory Oven & Wash Test"):
                    SbSql.Append(type8);
                    break;
                case string a when a.Contains("Pulling test for Snap/Button/Rivet"):
                    SbSql.Append(type9);
                    break;
                case string a when a.Contains("Water Fastness Test"):
                    SbSql.Append(type11);
                    break;
                case string a when a.Contains("Perspiration Fastness Test"):
                    SbSql.Append(type12);
                    break;
                case string a when a.Contains("Daily HT Wash"):
                    SbSql.Append(type13);
                    break;
                case string a when a.Contains("Daily Moisture"):
                    SbSql.Append(type14);
                    break;
                default:
                    SbSql.Append(
                        type1 + " union all " + 
                        type2 + " union all " +
                        type3 + " union all " +
                        type4 + " union all " +
                        type5 + " union all " +
                        type6 + " union all " +
                        type7 + " union all " +
                        type8 + " union all " +
                        type9 + " union all " +
                        type11 + " union all " +
                        type12 + " union all " +
                        type13 + " union all " +
                        type14);
                    break;
            }

            return ExecuteList<SearchList_Result>(CommandType.Text, SbSql.ToString(), objParameter, 90);
        }
    }
}
