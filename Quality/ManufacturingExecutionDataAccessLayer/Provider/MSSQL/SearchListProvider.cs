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
-- select Text = '', Value = ''
-- UNION
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
            string tmptable = string.Empty;
            string where = string.Empty;

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

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
		, AIComment = IIF(f_Result.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Fabric Crocking & Shrinkage Test',0,o.StyleID,o.BrandID,o.SeasonID) ) ,'')
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
{where}
";

            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.MDivisionID))
            {
                where += "AND g.MDivisionID = @MDivisionID ";
            }
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND g.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND g.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND g.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND g.Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
		, AIComment = IIF(gd.WashResult='F' ,(select dbo.GetQualityWebAIComment('Garment Wash Test',0,g.StyleID,g.BrandID,g.SeasonID)),'')
                        --+ CHAR(10)+CHAR(13)
                        + IIF(gd.SeamBreakageResult='F' ,(select dbo.GetQualityWebAIComment('Seam Breakage',0,g.StyleID,g.BrandID,g.SeasonID)),'')
                        --+ CHAR(10)+CHAR(13)
                        + IIF(gd.OdourResult='F' ,(select dbo.GetQualityWebAIComment('Odour Test',0,g.StyleID,g.BrandID,g.SeasonID)),'')
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) ON g.ID= gd.ID
left join Production.dbo.Pass1 pa on gd.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on gd.AddName = ma.ID

{sqlWhseArrival}
WHERE gd.Result <> ''
{where}

";

            // 重置
            where = string.Empty;
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

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                where += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                where += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                where += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                where += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
		, AIComment = IIF(m.Result = 'Fail' ,(select dbo.GetQualityWebAIComment('Mockup Crocking Test',0,m.StyleID,m.BrandID,m.SeasonID)) ,'')
from MockupCrocking m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
WHERE m.Result <> ''
{where}

";

            // 重置
            tmptable = string.Empty;
            where = string.Empty;
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

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND m.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND m.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND m.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                where += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                where += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                where += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                where += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	, AIComment = IIF(m.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Mockup Oven Test',0,m.StyleID,m.BrandID,m.SeasonID) ),'')
from MockupOven m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Type = 'B'
and m.Result <> ''
{where}
";

            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND m.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND m.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND m.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND m.Article = @Article ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                where += "AND ReceivedDate >= @ReceivedDate_s ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                where += "AND ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                where += "AND ReleasedDate >= @ReportDate_s ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                where += "AND ReleasedDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	, AIComment = IIF(m.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Mockup Wash Test',0,m.StyleID,m.BrandID,m.SeasonID) ),'')
from MockupWash m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Type = 'B' 
and m.Result <> ''
{where}
";
            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	    , AIComment = IIF(f.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Fabric Oven Test',0,o.StyleID,o.BrandID,o.SeasonID) ),'')
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
inner join Oven f WITH(NOLOCK) ON f.POID = p.ID
left join Production.dbo.Pass1 pa on f.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on f.AddName = ma.ID
{sqlWhseArrival}
where f.Result <> ''
{where}

";

            // 重
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	    , AIComment = IIF(f.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Washing Fastness',0,o.StyleID,o.BrandID,o.SeasonID) ),'')
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
INNER JOIN ColorFastness f WITH(NOLOCK) ON f.POID = p.ID
left join Production.dbo.Pass1 pa on f.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on f.AddName = ma.ID
{sqlWhseArrival}
WHERE f.Result <> ''
{where}
";

            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	    , AIComment = IIF(f_Result.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Accessory Oven & Wash Test',0,o.StyleID,o.BrandID,o.SeasonID) ),'')
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
{where}
";

            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND m.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND m.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND m.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND m.Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	, AIComment = IIF(m.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Pulling test for Snap/Button/Rivet',0,m.StyleID,m.BrandID,m.SeasonID) ),'')
from [ExtendServer].ManufacturingExecution.dbo.PullingTest m WITH(NOLOCK)
left join Production.dbo.Pass1 pa on m.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on m.AddName = ma.ID
{sqlWhseArrival}
where m.Result <> ''
{where}

";

            // 重置
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND w.Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
            }

            tmptable = $@"----Search List查詢
DECLARE @IsRRLR as bit = 0

----取得AIComment
select  ad.Type,ad.IsRRLR,ad.Comment
INTO #AIComment
from ExtendServer.ManufacturingExecution.dbo.AIComment_Detail ad
where ad.AICommentUkey in (
	select Ukey from ExtendServer.ManufacturingExecution.dbo.AIComment where FunctionName='QualityWeb'
)
and ad.Type ='Water Fastness Test'

select TOP 1 @IsRRLR = isRRLR from #AIComment

----取得 存在RR/LR的ReportNo
select DISTINCT w.ReportNo
INTO #RRLR_ACH
from WaterFastness w WITH (NOLOCK) 
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
inner join Style s ON s.ID = o.StyleID and s.BrandID=o.BrandID and s.SeasonID = o.SeasonID
inner join Style_RRLR_Report srr on s.Ukey=srr.StyleUkey
{sqlWhseArrival}
WHERE w.Result='Fail'
and RRRemark like '%ACH%'
{where}

----取得 存在RR/LR的LR的ReportNo
select DISTINCT w.ReportNo
INTO #RRLR_CF
from WaterFastness w WITH (NOLOCK) 
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
inner join Style s ON s.ID = o.StyleID and s.BrandID=o.BrandID and s.SeasonID = o.SeasonID
inner join Style_RRLR_Report srr on s.Ukey=srr.StyleUkey
{sqlWhseArrival}
WHERE w.Result='Fail'
and RRRemark like '%CF%'
{where}
";

            string type11 = $@"
{tmptable}

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
	    , AIComment = IIF(w.Result = 'Fail' ,AIComment.Val,'')
from WaterFastness w WITH (NOLOCK) 
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
left join Production.dbo.Pass1 pa on w.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on w.AddName = ma.ID
outer apply(
    -----只秀出有Fail的AIComment訊息，再加上RR/LR訊息
	SELECT Val=
		ISNULL( (select Comment from #AIComment where w.Result = 'Fail' ) ,'')
		+CHAR(10)+CHAR(13)+
		+ (CASE  WHEN @IsRRLR = 0 THEN ''
                 WHEN ((select COUNT(1) from #RRLR_ACH s where s.ReportNo = w.ReportNo )>0 and (select COUNT(1) from #RRLR_CF s where s.ReportNo = w.ReportNo ) > 0) THEN 'There is RR/LR (With shade achievability issue, please ensure shading within tolerance as agreement. Lower color fastness waring, please check if need to apply tissue paper.)'
				 WHEN ((select COUNT(1) from #RRLR_ACH s where s.ReportNo = w.ReportNo ) > 0) THEN 'With shade achievability issue, please ensure shading within tolerance as agreement.'
				 WHEN ((select COUNT(1) from #RRLR_CF s where s.ReportNo = w.ReportNo ) > 0) THEN 'Lower color fastness waring, please check if need to apply tissue paper.'
				ELSE''
			END
		)
)AIComment
{sqlWhseArrival}
WHERE w.Result <> ''
{where}

drop table #AIComment,#RRLR_ACH,#RRLR_CF
";

            // 重置
            tmptable = string.Empty;
            where = string.Empty;
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

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND o.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND o.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND o.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND w.Article = @Article ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	    , AIComment = IIF(w.Result = 'Fail' ,( select dbo.GetQualityWebAIComment('Perspiration Fastness Test',0,o.StyleID,o.BrandID,o.SeasonID) ),'')
from PerspirationFastness w WITH (NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
left join Production.dbo.Pass1 pa on w.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on w.AddName = ma.ID
{sqlWhseArrival}
WHERE w.Result <> ''
{where}
";

            // 重置
            where = string.Empty;
            #endregion

            #region Daily HT Wash Test

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND h.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND h.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND h.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND h.Article = @Article ";
            }
            if (!string.IsNullOrEmpty(Req.Line))
            {
                where += "AND h.Line = @Line ";
            }
            if (Req.ReceivedDate_s.HasValue)
            {
                where += " AND @ReceivedDate_s <= h.ReceivedDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                where += " AND h.ReceivedDate <= @ReceivedDate_e ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                where += " AND @ReportDate_s <= h.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                where += " AND h.ReportDate <= @ReportDate_e ";
            }

            string type13 = $@"
select DISTINCT Type= 'Daily HT Wash Test'
        , h.ReportNo
		, h. OrderID
		, h.StyleID
		, h.BrandID
		, h.SeasonID
		, h.Article 
		, Artwork = ''
		, h.Result
		, TestDate = h.ReportDate
	    , ReceivedDate = h.ReceivedDate
	    , ReportDate = h.ReportDate
        , AddName = ISNULL(pa.Name, ma.Name)
	    , AIComment = ''
from [ExtendServer].ManufacturingExecution.dbo.HeatTransferWash h WITH (NOLOCK)
left join Production.dbo.Pass1 pa on h.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on h.AddName = ma.ID
WHERE h.Result <> ''
{where}
";

            // 重置
            tmptable = string.Empty;
            where = string.Empty;
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
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                where += "AND h.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                where += "AND h.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                where += "AND h.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                where += "AND h.Article = @Article ";
            }
            if (!string.IsNullOrEmpty(Req.Line))
            {
                where += "AND h.Line = @Line ";
            }
            if (Req.ReportDate_s.HasValue)
            {
                where += " AND @ReportDate_s <= h.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                where += " AND h.ReportDate <= @ReportDate_e ";
            }
            if (Req.WhseArrival_s.HasValue)
            {
                where += " AND @WhseArrival_s <= e.WhseArrival ";
            }
            if (Req.WhseArrival_e.HasValue)
            {
                where += " AND @WhseArrival_e >= e.WhseArrival ";
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
	    , AIComment = ''
from [ExtendServer].ManufacturingExecution.dbo.BulkMoistureTest h WITH (NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = h.OrderID
left join Production.dbo.Pass1 pa on h.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on h.AddName = ma.ID
{sqlWhseArrival}
WHERE h.Result <> ''
{where}
";

            // 重置
            tmptable = string.Empty;
            where = string.Empty;
            #endregion

            #region AgingHydrolysisTest (461)

            string type15 = $@"
select Type= 'Accelerated Aging by Hydrolysis (461)'
	, a.ReportNo
	, b.OrderID
	, b.StyleID
	, b.BrandID
	, b.SeasonID
	, b.Article
	, Line = ''
	, Artwork = ''
	, a.Result	
	, TestDate = Cast( NULL as date)
	, ReceivedDate = a.ReceivedDate
	, ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)

from [ExtendServer].ManufacturingExecution.dbo.AgingHydrolysisTest_Detail a
inner join [ExtendServer].ManufacturingExecution.dbo.AgingHydrolysisTest b on b.ID = a.AgingHydrolysisTestID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type15 += "AND BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type15 += "AND SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type15 += "AND StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type15 += "AND Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type15 += " AND @ReceivedDate_s <= ReceivedDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type15 += " AND ReceivedDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type15 += " AND @ReportDate_s <= ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type15 += " AND ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Phenolic Yellow Test (510)

            string type16 = $@"
select Type= 'Phenolic Yellowing Test (510)'
	, a.ReportNo
	, a.OrderID
	, a.StyleID
	, a.BrandID
	, a.SeasonID
	, a.Article
	, Line = ''
	, Artwork = ''
	, a.Result	
	, TestDate = Cast( NULL as date)
	, ReceivedDate = a.SubmitDate
	, ReportDate = a.ReportDate
    , AddName = ISNULL(ma.Name, pa.Name)
from [ExtendServer].ManufacturingExecution.dbo.PhenolicYellowTest a WITH (NOLOCK) 
left join Production.dbo.Pass1 pa on a.AddName = pa.ID
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 ma on a.AddName = ma.ID
WHERE a.Result <> '' AND　ａ.ReportDate IS NOT NULL
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type16 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type16 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type16 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type16 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type16 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type16 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type16 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type16 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Saliva Fastness Test (519)

            string type17 = $@"
select Type= 'Saliva Fastness Test (519)'
	, a.ReportNo
	, a.OrderID
	, a.StyleID
	, a.BrandID
	, a.SeasonID
	, a.Article
	, Line = ''
	, Artwork = ''
	, a.Result	
	, TestDate = Cast( NULL as date)
	, ReceivedDate = a.SubmitDate
	, ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)

from [ExtendServer].ManufacturingExecution.dbo.SalivaFastnessTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type17 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type17 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type17 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type17 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type17 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type17 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type17 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type17 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region T-Peel Strength Test (438)

            string type18 = $@"
select Type= 'T-Peel Strength Test (438)'
	, a.ReportNo
	, a.OrderID
	, a.StyleID
	, a.BrandID
	, a.SeasonID
	, a.Article
	, Line = ''
	, Artwork = ''
	, a.Result	
	, TestDate = Cast( NULL as date)
	, ReceivedDate = a.SubmitDate
	, ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)

from [ExtendServer].ManufacturingExecution.dbo.TPeelStrengthTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.Result <> ''
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type18 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type18 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type18 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type18 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type18 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type18 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type18 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type18 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Hydrostatic Pressure Waterproof Test (602)

            string type19 = $@"
select Type= 'Hydrostatic Pressure Waterproof Test (602)'
    , a.ReportNo
    , a.OrderID
    , a.StyleID
    , a.BrandID
    , a.SeasonID
    , a.Article
    , Line = ''
    , Artwork = ''
    , a.Result  
    , TestDate = Cast( NULL as date)
    , ReceivedDate = a.SubmitDate
    , ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)
from [ExtendServer].ManufacturingExecution.dbo.HydrostaticPressureWaterproofTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type19 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type19 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type19 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type19 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type19 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type19 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type19 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type19 += " AND a.ReportDate <= @ReportDate_e ";
            }
            #endregion

            #region Martindale Pilling Test (452)

            string type20 = $@"
select Type= 'Martindale Pilling Test (452)'
    , a.ReportNo
    , a.OrderID
    , a.StyleID
    , a.BrandID
    , a.SeasonID
    , a.Article
    , Line = ''
    , Artwork = ''
    , a.Result  
    , TestDate = Cast( NULL as date)
    , ReceivedDate = a.SubmitDate
    , ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)
from [ExtendServer].ManufacturingExecution.dbo.MartindalePillingTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type20 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type20 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type20 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type20 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type20 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type20 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type20 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type20 += " AND a.ReportDate <= @ReportDate_e ";
            }
            #endregion

            #region Random Tumble Pilling Test (407)

            string type21 = $@"
select Type= 'Random Tumble Pilling Test (407)'
    , a.ReportNo
    , a.OrderID
    , a.StyleID
    , a.BrandID
    , a.SeasonID
    , a.Article
    , Line = ''
    , Artwork = ''
    , a.Result  
    , TestDate = Cast( NULL as date)
    , ReceivedDate = a.SubmitDate
    , ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)

from [ExtendServer].ManufacturingExecution.dbo.RandomTumblePillingTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type21 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type21 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type21 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type21 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type21 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type21 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type21 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type21 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Residue/Ageing Test for Sticker (434)
            string type22 = $@"
select Type= 'Residue/Ageing Test for Sticker (434)'
	, a.ReportNo
	, a.OrderID
	, a.StyleID
	, a.BrandID
	, a.SeasonID
	, a.Article
	, Line = ''
	, Artwork = ''
	, a.Result	
	, TestDate = Cast( NULL as date)
	, ReceivedDate = a.SubmitDate
	, ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)
from [ExtendServer].ManufacturingExecution.dbo.StickerTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type22 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type22 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type22 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type22 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type22 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type22 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type22 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type22 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Water Absorbency Test (604)
            string type23 = $@"
select Type= 'Water Absorbency Test (604)'
    , a.ReportNo
    , a.OrderID
    , a.StyleID
    , a.BrandID
    , a.SeasonID
    , a.Article
    , Line = ''
    , Artwork = ''
    , a.Result  
    , TestDate = Cast( NULL as date)
    , ReceivedDate = a.SubmitDate
    , ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)
from [ExtendServer].ManufacturingExecution.dbo.WaterAbsorbencyTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type23 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type23 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type23 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type23 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type23 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type23 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type23 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type23 += " AND a.ReportDate <= @ReportDate_e ";
            }

            #endregion

            #region Evaporation Rate Test(617)
            string type24 = $@"
select Type= 'Evaporation Rate Test (617)'
    , a.ReportNo
    , a.OrderID
    , a.StyleID
    , a.BrandID
    , a.SeasonID
    , a.Article
    , Line = ''
    , Artwork = ''
    , a.Result  
    , TestDate = Cast( NULL as date)
    , ReceivedDate = a.SubmitDate
    , ReportDate = a.ReportDate
    , AddName = ISNULL(mp.Name, pp.Name)
from [ExtendServer].ManufacturingExecution.dbo.EvaporationRateTest a
left join [ExtendServer].ManufacturingExecution.dbo.Pass1 mp on a.EditName = mp.ID
left join Pass1 pp on a.EditName = pp.ID
WHERE a.ReportDate IS NOT NULL
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type24 += "AND a.BrandID = @BrandID ";
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type24 += "AND a.SeasonID = @SeasonID ";
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                type24 += "AND a.StyleID = @StyleID ";
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type24 += "AND a.Article = @Article ";
            }

            if (Req.ReceivedDate_s.HasValue)
            {
                type24 += " AND @ReceivedDate_s <= a.SubmitDate ";
            }
            if (Req.ReceivedDate_e.HasValue)
            {
                type24 += " AND a.SubmitDate <= @ReceivedDate_e ";
            }

            if (Req.ReportDate_s.HasValue)
            {
                type24 += " AND @ReportDate_s <= a.ReportDate ";
            }
            if (Req.ReportDate_e.HasValue)
            {
                type24 += " AND a.ReportDate <= @ReportDate_e ";
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
                case string a when a.Contains("Accelerated Aging by Hydrolysis"):
                    SbSql.Append(type15);
                    break;
                case string a when a.Contains("Phenolic Yellowing Test"):
                    SbSql.Append(type16);
                    break;
                case string a when a.Contains("Saliva Fastness Test"):
                    SbSql.Append(type17);
                    break;
                case string a when a.Contains("T-Peel Strength Test"):
                    SbSql.Append(type18);
                    break;
                case string a when a.Contains("Hydrostatic Pressure Waterproof Test"):
                    SbSql.Append(type19);
                    break;
                case string a when a.Contains("Martindale Pilling Test"):
                    SbSql.Append(type20);
                    break;
                case string a when a.Contains("Random Tumble Pilling Test"):
                    SbSql.Append(type21);
                    break;
                case string a when a.Contains("Residue/Ageing Test for Sticker"):
                    SbSql.Append(type22);
                    break;
                case string a when a.Contains("Water Absorbency Test"):
                    SbSql.Append(type23);
                    break;
                case string a when a.Contains("Evaporation Rate Test"):
                    SbSql.Append(type24);
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
                        type14 + " union all " +
                        type15 + " union all " +
                        type16 + " union all " +
                        type17 + " union all " +
                        type18 + " union all " +
                        type19 + " union all " +
                        type20 + " union all " +
                        type21 + " union all " +
                        type22 + " union all " +
                        type23 + " union all " +
                        type24);
                    break;
            }

            return ExecuteList<SearchList_Result>(CommandType.Text, SbSql.ToString(), objParameter, 90);
        }
    }
}
