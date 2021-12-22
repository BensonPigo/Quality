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
    public class RFTPerLineProvider : SQLDAL, IRFTPerLineProvider
    {
        #region 底層連線
        public RFTPerLineProvider(string conString) : base(conString) { }
        public RFTPerLineProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<MonthlyRFT> GetMonthlyRFT(string FactoryID, string Year, int monthint)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@FactoryID", DbType.String, FactoryID);
            objParameter.Add("@Year", DbType.String, Year);
            objParameter.Add("@Month", DbType.Int32, monthint);

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append(@"
select
	 rft.Line,
	 rft.Status
into #tmpRft
from RFT_Inspection rft WITH(NOLOCK)
where rft.FactoryID = @FactoryID
and Year(rft.InspectionDate) = @Year
and Month(rft.InspectionDate) = @Month

select
	Month = DateName(Month, DateAdd(Month, @Month, -1)),
	Line = s.ID,
	RFT = cast(iif(exist.Line is null, null ,iif(ttl.ct = 0, 0, round((isnull(pass.ct, 0) * 1.0 / ttl.ct) * 100, 2))) as decimal(5,2)) 
from SciProduction_SewingLine s WITH(NOLOCK)
outer apply(select ct = count(1) from #tmpRft where Line = s.ID )ttl
outer apply(select ct = count(1) from #tmpRft where Line = s.ID and Status = 'Pass')pass
left join (
	select distinct Line
	from #tmpRft
)exist on  Line = s.ID 
where s.Junk = 0
and s.FactoryID = @FactoryID

drop table #tmpRft
");
            return ExecuteList<MonthlyRFT>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<DailyRFT> GetDailyRFT(string FactoryID, string Year, int monthint)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@FactoryID", DbType.String, FactoryID);
            objParameter.Add("@Year", DbType.String, Year);
            objParameter.Add("@Month", DbType.Int32, monthint);
            StringBuilder SbSql = new StringBuilder();
            SbSql.Append(@"
DECLARE @date DATETIME = concat(@Year,'/',@Month,'/','1')
DECLARE @d1 int = 1
DECLARE @d2 int =(SELECT day(DATEADD(Month, ((YEAR(@date) - 1900) * 12) + Month(@date), -1)))

;WITH cte AS (
    SELECT [date] = @d1
    UNION ALL
    SELECT [date] + 1 FROM cte WHERE ([date] < @d2)
)
SELECT [date]
into #tmpAllday
FROM cte

select
	 rft.Line,
	 Date = DAY(rft.InspectionDate),
	 rft.Status
into #tmpRft
from RFT_Inspection rft WITH(NOLOCK)
where rft.FactoryID = @FactoryID
and Year(rft.InspectionDate) = @Year
and Month(rft.InspectionDate) = @Month

select
	Date = d.date,
	Month = DateName(Month, DateAdd(Month, @Month, -1)),
	Line = s.ID,
	RFT = cast(iif(exist.Line is null, null ,iif(ttl.ct = 0, 0, round((isnull(pass.ct, 0) * 1.0 / ttl.ct) * 100, 2))) as decimal(5,2))
from SciProduction_SewingLine s WITH(NOLOCK)
cross join #tmpAllday d
outer apply(select ct = count(1) from #tmpRft where Line = s.ID and Date = d.date)ttl
outer apply(select ct = count(1) from #tmpRft where Line = s.ID and Date = d.date and Status = 'Pass')pass
left join (
	select distinct Line, Date
	from #tmpRft
)exist on exist.Line = s.ID and exist.Date = d.date
where s.Junk = 0
and s.FactoryID = @FactoryID

drop table #tmpRft,#tmpAllday
");
            return ExecuteList<DailyRFT>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
