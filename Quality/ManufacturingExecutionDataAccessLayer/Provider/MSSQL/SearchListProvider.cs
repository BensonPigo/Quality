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
            };

            StringBuilder SbSql = new StringBuilder();

            #region Fabric Crocking & Shrinkage Test (504, 405)
            string type1 = $@"
select Type = 'Fabric Crocking & Shrinkage Test (504, 405)'
        , ReportNo=''
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article = ''
		,Artwork = ''
		, [Result] = f_Result.Result
		, [TestDate] = f_TestDate.TestDate
	    ,ReceivedDate =NULL
	    ,ReportDate =NULL
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
WHERE exists (select 1 from FIR_Laboratory f WITH(NOLOCK) WHERE f.POID = p.ID)
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
            #endregion

            #region Garment Test (450, 451, 701, 710)
            string type2 = $@"
select  Type = 'Garment Test (450, 451, 701, 710)'
        ,ReportNo = cast(gd.No as varchar(50))
		,gd.OrderID
		,StyleID
		,BrandID
		,SeasonID
		,Article
		,Artwork = ''
		,Result= IIF(gd.Result='P','Pass', IIF(gd.Result='F','Fail',''))
		,TestDate = gd.InspDate
	,ReceivedDate =NULL
	,ReportDate =NULL
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) ON g.ID= gd.ID
WHERE 1=1 
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
            #endregion

            #region Mockup Crocking Test  (504)
            string type3 = $@"
select DISTINCT  Type = 'Mockup Crocking Test  (504)'
        ,ReportNo
		,OrderID = POID
		,StyleID
		,BrandID
		,SeasonID
		,Article
		,Artwork = ArtworkTypeID
		,Result
		,TestDate 
        , ReceivedDate
        , ReportDate = ReleasedDate
from MockupCrocking  WITH(NOLOCK)
WHERE 1=1 
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
            #endregion

            #region Mockup Oven Test (514)
            string type4 = $@"
select DISTINCT Type = 'Mockup Oven Test (514)'
	, m.ReportNo
	, m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = m.ArtworkTypeID
	, m.Result
	, m.TestDate
    , ReceivedDate
    , ReportDate = ReleasedDate
from MockupOven m WITH(NOLOCK)
where m.Type = 'B'
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
            #endregion

            #region Mockup Wash Test (701)
            string type5 = $@"
select DISTINCT Type = 'Mockup Wash Test (701)'
	, m.ReportNo
	, m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = m.ArtworkTypeID
	, m.Result
	, m.TestDate
    , ReceivedDate
    , ReportDate = ReleasedDate
from MockupWash m WITH(NOLOCK)
where m.Type = 'B' 
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
            #endregion

            #region Fabric Oven Test (515) 
            string type6 = $@"
select DISTINCT Type= 'Fabric Oven Test (515)'
        ,ReportNo = cast(f.TestNo as varchar(50))
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,f.Article  
		,Artwork = ''
		,Result=f.Result
		,TestDate = f.InspDate
	    ,ReceivedDate =NULL
	    ,ReportDate =NULL
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
inner join Oven f WITH(NOLOCK) ON f.POID = p.ID
where 1=1 
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
            #endregion

            #region Washing Fastness (501)
            string type7 = $@"
select DISTINCT Type= 'Washing Fastness (501)'
        ,ReportNo = cast(f.TestNo as varchar(50))
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,f.Article 
		,Artwork = ''
		,Result=f.Result
		,TestDate = f.InspDate

	    ,ReceivedDate =NULL
	    ,ReportDate =NULL
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
INNER JOIN ColorFastness f WITH(NOLOCK) ON f.POID = p.ID
WHERE 1=1 
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
            #endregion

            #region Accessory Oven & Wash Test (515, 701)
            string type8 = $@"
select Type = 'Accessory Oven & Wash Test (515, 701)'
        ,ReportNo=''
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article = ''
		,[Artwork] = ''
		, [Result] = f_Result.Result
		, [TestDate] = f_TestDate.TestDate
		,ReceivedDate =NULL
		,ReportDate =NULL
from PO p
inner join Orders o WITH(NOLOCK) ON o.POID = p.ID
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
WHERE exists (select 1 from AIR_Laboratory f WITH(NOLOCK) WHERE f.POID = p.ID)
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

            #endregion

            #region Pulling test for Snap/Botton/Rivet (437)
            string type9 = $@"
select DISTINCT Type = 'Pulling test for Snap/Botton/Rivet (437)'
	, m.ReportNo
	, m.POID
	, m.StyleID
	, m.BrandID
	, m.SeasonID
	, m.Article
	, [Artwork] = ''
	, m.Result
	, m.TestDate
	,ReceivedDate =NULL
	,ReportDate =NULL
from [ExtendServer].ManufacturingExecution.dbo.PullingTest m  WITH(NOLOCK)
where 1=1 
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

            #endregion

            #region Water Fastness Test(503)
            string type11 = $@"
select DISTINCT Type= 'Water Fastness Test(503)'
        , ReportNo=''
		, w.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article 
		,Artwork = ''
		,w.Result
		,TestDate = w.InspDate

	    ,ReceivedDate =NULL
	    ,ReportDate =NULL
from WaterFastness w WITH (NOLOCK) 
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
WHERE 1=1 
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
            #endregion

            #region Perspiration Fastness Test(502)
            string type12 = $@"
select DISTINCT Type= 'Perspiration Fastness (502)'
        , ReportNo=''
		, w.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article 
		,Artwork = ''
		,w.Result
		,TestDate = w.InspDate

	    ,ReceivedDate =NULL
	    ,ReportDate =NULL
from PerspirationFastness w WITH (NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = w.POID
WHERE 1=1 
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
                case string a when a.Contains("Water Fastness Test(503)"):
                    SbSql.Append(type11);
                    break;
                case string a when a.Contains("Perspiration Fastness Test(502)"):
                    SbSql.Append(type12);
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
                        type12);
                    break;
            }

            return ExecuteList<SearchList_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
