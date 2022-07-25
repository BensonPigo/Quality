using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.EtoEFlowChart;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class EtoEFlowChartProvider : SQLDAL
    {
        #region 底層連線
        public EtoEFlowChartProvider(string ConString) : base(ConString) { }
        public EtoEFlowChartProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public long Get_StyleUkey(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
            };
            string sqlcmd = @"
select Ukey 
from Style 
where ID = @StyleID 
AND BrandID = @BrandID
AND SeasonID = @SeasonID
";
            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToInt64(result == null ? 0 : result);
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public decimal Get_Development_SampleRFT(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = @"
select SampleRftRate = ROUND( ISNULL( 1.0 * SUM(IIF(Status='Pass',1,0)) / count(StyleUkey) * 100.0 ,0) ,2)

from RFT_Inspection rft WITH(NOLOCK)
where StyleUkey = @StyleUkey

";
            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToDecimal(result == null ? 0 : result);
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Development_SampleRFT> Get_Development_SampleRFT_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = @"
select distinct OrderID
    ,SampleStage = o.OrderTypeID
    ,s.BuyReadyDate
from RFT_Inspection rft WITH(NOLOCK)
INNER JOIN SciProduction_Orders o ON  rft.OrderID=o.ID
INNER JOIN MainServer.Production.dbo.Style_Article s ON  s.StyleUkey=rft.StyleUkey
where rft.StyleUkey = @StyleUkey

";
            var result = ExecuteList<Development_SampleRFT>(CommandType.Text, sqlcmd, objParameter) ?? new List<Development_SampleRFT>();
            return result.ToList();
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Development_SampleRFT_CFTComments> Get_Development_SampleRFT_Detail_CFTComments(List<string> OrderIDList)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string orders = "'" + string.Join("','", OrderIDList) + "'";

            string sqlcmd = $@"
----DECLARE @OrderID as varchar(13) =N'22062072BBS'

select r.OrderID
    ,o.OrderTypeID
    , PMS_RFTCommentsID = d.Description
    , r.Comnments
into #tmpBase
from SciProduction_Orders o with(nolock)
inner join RFT_OrderComments r WITH(NOLOCK) on r.OrderID = o.id
left join SciProduction_DropdownList d WITH(NOLOCK) ON d.Type='PMS_RFTComments' AND d.ID = r.PMS_RFTCommentsID
where 1=1
and o.junk = 0
and o.Category = 'S'
and o.id IN ({orders})

select SampleStage = a.OrderTypeID
	, CommentCategory = a.PMS_RFTCommentsID
	, Comment = c.Comnments
	,a.OrderID
from (select distinct OrderID,OrderTypeID,PMS_RFTCommentsID from #tmpBase) a
outer apply(
    select Comnments = stuff((
        select concat('*', Comnments)
        from #tmpBase b
        where a.OrderTypeID = b.OrderTypeID and a.PMS_RFTCommentsID = b.PMS_RFTCommentsID and a.OrderID=b.OrderID
        and Comnments <> ''
        for xml path('')
    ),1,1,'')
)c
order by a.OrderID, a.OrderTypeID, a.PMS_RFTCommentsID

drop table #tmpBase

";

            var result = ExecuteList<Development_SampleRFT_CFTComments>(CommandType.Text, sqlcmd, objParameter) ?? new List<Development_SampleRFT_CFTComments>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Development_TestResult> Get_Development_TestResult_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
------正式機執行

select * from (
    select * from (
        select TOP 1 Category = 'Fabric Crocking & Shrinkage Test (504, 405)'
                ,ReportNo
                ,Article 
                ,TestDate = (
                    SELECT MAX (TestDate) FROM (
                        SELECT TestDate = CrockingTestDate
                        UNION
                        SELECT TestDate = HeatTestDate
                        UNION 
                        SELECt TestDate = WashTestDate
                    )tmp
                )
                ,TestResult=Result 
        from Quality.dbo.FabricTest WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        ORDER BY TestDate DESC
    )a
    UNION
    select * from (
        select TOP 1 Category = 'Garment Test (450, 451, 701, 710)'
                ,ReportNo = gd.ReportNo--cast(gd.No as varchar(50))
                ,g.Article
                ,TestDate =  gd.InspDate
                ,TestResult= IIF(gd.Result='P','Pass', IIF(gd.Result='F','Fail',''))
        from Quality.dbo.SampleGarmentTest g WITH (NOLOCK)
        inner join SampleGarmentTest_Detail gd WITH (NOLOCK) ON g.ID= gd.ID
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        ORDER BY gd.InspDate DESC
    )a
    UNION
    select * from (
        select TOP 1  Category = 'Mockup Crocking Test  (504)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.MockupCrocking  WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select TOP 1  Category = 'Mockup Oven Test  (514)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.MockupOven  WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select TOP 1  Category = 'Mockup Wash Test  (701)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.MockupOven  WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select TOP 1  Category = 'Fabric Oven Test (515)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.FabricOven   WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select TOP 1  Category = 'Washing Fastness (501)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.FabricColorFastness    WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select Top 1 Category = 'Accessory Oven & Wash Test (515, 701)'
                ,ReportNo
                ,Article
                ,TestDate = (
                    SELECT MAX (TestDate) FROM (
                        SELECT TestDate = WashTestDate
                        UNION
                        SELECT TestDate = OvenTestDate
                    )tmp
                )
                ,TestResult = Result
        from Quality.dbo.AccessoryTest     WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select Top 1 Category = 'Pulling test for Snap/Botton/Rivet (437)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.PullingTest      WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select Top 1 Category = 'Water Fastness Test(503)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.WaterFastness WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
    UNION
    select * from (
        select Top 1 Category = 'Perspiration Fastness (502)'
                ,ReportNo
                ,Article
                ,TestDate 
                ,TestResult = Result
        from Quality.dbo.PerspirationFastness  WITH (NOLOCK)
        WHERE StyleID=@StyleID 
            AND BrandID=@BrandID 
            AND SeasonID=@SeasonID 
            {(string.IsNullOrEmpty(Req.Article) ? string.Empty : "AND Article = @Article")}
        order by TestDate desc
    )a
)b
where TestResult = 'Fail'

";

            var result = ExecuteList<Development_TestResult>(CommandType.Text, sqlcmd, objParameter) ?? new List<Development_TestResult>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Development_RRLR> Get_Development_RRLR_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"

--DECLARE @StyleID as varchar(30)='ARWRS23007'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='23SS'

select Refno
        ,SuppID
        ,ColorID
        ,RR = IIF(a.RR=1,'V','')
        ,RRRemark
        ,LR= IIF(a.LR=1,'V','')
from Style_RRLR_Report a
inner join Style b on a.StyleUkey = b.Ukey
where b.ID = @StyleID AND b.BrandID = @BrandID AND b.SeasonID = @SeasonID
";

            var result = ExecuteList<Development_RRLR>(CommandType.Text, sqlcmd, objParameter) ?? new List<Development_RRLR>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Development_FD> Get_Development_FD_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"

--DECLARE @StyleID as varchar(30)='ARWRS23007'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='23SS'

select StyleID = s.ID
    ,s.BrandID
    ,s.SeasonID
    ,Article = Article.Val
    ,ExpectionFormStatus = d.Name
    ,s.ExpectionFormDate
    ,s.ExpectionFormRemark
from  Style s
left join DropDownList d ON d.Type = 'FactoryDisclaimer' AND s.ExpectionFormStatus = d.ID
outer apply(
    select Val = STUFF((
        select DISTINCt ',' + Article
        from Style_Article sa WITH(NOLOCK)
        where sa.StyleUkey = s.Ukey
        for xml path('')
        ),1,1,'')
)Article

where s.ID = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
AND  s.ExpectionFormDate >= DATEADD(Year,-2, GETDATE())

";

            var result = ExecuteList<Development_FD>(CommandType.Text, sqlcmd, objParameter) ?? new List<Development_FD>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public string Get_Warehouse_PhysicalInspection(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='S1805GHTT213'
--DECLARE @BrandID as varchar(30)='ADIDAS'
--DECLARE @SeasonID as varchar(30)='21SS'

SELECT distinct
    fir.POID
    ,s.Ukey
    ,SEQ1
    ,SEQ2
    ,Refno
    ,Physical
    ,s.ID ,s.SeasonID ,s.BrandID
INTO #tmp
FROM FIR fir
left join orders o on o.POID = fir.poid
left join Style s on s.Ukey = o.StyleUkey
where Physical <> ''
    AND s.BrandID=@BrandID 
    AND s.SeasonID = @SeasonID
    and s.ID= @StyleID 
order by seq1,seq2

select Result = CASE WHEN SUM(IIF(Physical='Fail',1,0) ) > 0 THEN 'Fail'
                     WHEN SUM(IIF(Physical='Pass',1,0) ) = COUNT(1) AND COUNT(1) > 0 THEN 'Pass'
                     ELSE 'N/A'
                END
from #tmp

drop table #tmp


";
            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return result == null ? string.Empty : result.ToString();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Warehouse_PhysicalInspection> Get_Warehouse_PhysicalInspection_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='S1805GHTT213'
--DECLARE @BrandID as varchar(30)='ADIDAS'
--DECLARE @SeasonID as varchar(30)='21SS'

SELECT distinct
    OrderID = fir.POID
    ,Seq1
    ,Seq2
    ,Refno
    ,Result = Physical

FROM FIR fir
left join orders o on o.POID = fir.poid
left join Style s on s.Ukey = o.StyleUkey
where Physical <> ''
AND s.BrandID=@BrandID 
AND s.SeasonID = @SeasonID
and s.ID= @StyleID 

order by seq1,seq2

";

            var result = ExecuteList<Warehouse_PhysicalInspection>(CommandType.Text, sqlcmd, objParameter) ?? new List<Warehouse_PhysicalInspection>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Production_SubprocessRFT> Get_Production_SubprocessRFT_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

-- 哪些綁包去做這工段
select SubProcessID
    ,spir.BundleNo 
    ,RejectQty
    ,BundleQty = ISNULL(bd.Qty,0)
into #tmp_subprocess_bundle
from SubProInsRecord spir　
left join Bundle_Detail bd on spir.BundleNo=bd.BundleNo
WHERE 1=1
and spir.BundleNo  in (
    select BundleNo 
    from Bundle_Detail_Order bdo 
    left join orders od on od.id = bdo.OrderID 
    where od.StyleUkey=@StyleUkey
)

select SubProcessID
    , RejectRate = ROUND( IIF(SUM(BundleQty) = 0 , 0 ,　1.0 * (SUM(BundleQty) - SUM(RejectQty)) / SUM(BundleQty)) ,2)
    , SummaryRate = ROUND( (select RejectRate = IIF(SUM(BundleQty) = 0 , 0 ,　1.0 * (SUM(BundleQty) - SUM(RejectQty)) / SUM(BundleQty)) from #tmp_subprocess_bundle) ,2)
from #tmp_subprocess_bundle
group by SubProcessID

drop table #tmp_subprocess_bundle
";

            var result = ExecuteList<Production_SubprocessRFT>(CommandType.Text, sqlcmd, objParameter) ?? new List<Production_SubprocessRFT>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Production_TestResult> Get_Production_TestResult_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

select　b.WashResult
into #tmp
from GarmentTest a
left join GarmentTest_Detail b on a.ID= b.ID
where StyleID=@StyleID AND BrandID=@BrandID AND SeasonID=@SeasonID


select　c.TestResult
into #tmp2
from GarmentTest a
left join GarmentTest_Detail b on a.ID= b.ID
left join GarmentTest_Detail_FGPT c on b.ID=c.ID
where StyleID=@StyleID AND BrandID=@BrandID AND SeasonID=@SeasonID

SELECT FGWTResult = (
    select FGWTResult = CASE WHEN SUM(IIF(WashResult='F',1,0) ) > 0 THEN 'Fail'
                         WHEN SUM(IIF(WashResult='P',1,0) ) = COUNT(1) AND COUNT(1) > 0 THEN 'Pass'
                         ELSE 'N/A'
                    END
    from #tmp
), FGPTResult = (
    select FGPTResult = CASE WHEN SUM(IIF(TestResult='F',1,0) ) > 0 THEN 'Fail'
                            WHEN SUM(IIF(TestResult='P',1,0) ) = COUNT(1) AND COUNT(1) > 0 THEN 'Pass'
                            ELSE 'N/A'
                    END
    from #tmp2
)

drop table #tmp ,#tmp2
";

            var result = ExecuteList<Production_TestResult>(CommandType.Text, sqlcmd, objParameter) ?? new List<Production_TestResult>();
            return result.ToList();
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public decimal Get_Production_InlineRFT(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

select [InlineRFT] = iif(
    (isnull(SUM(PassWIP), 0) + isnull(SUM(RejectWIP), 0)) = 0
    , 0
    , ROUND(  isnull(SUM(PassWIP), 0) * 1.0 / (isnull(SUM(PassWIP), 0) + isnull(SUM(RejectWIP), 0)) ,2)
)
from InlineInspection  with (nolock)
where StyleUkey = @StyleUkey
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}
";
            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToDecimal(result == null ? 0 : result);
        }
        
        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Production_InlineRFT> Get_Production_InlineRFT_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

select
    ord.BrandID as [Brand]
    , ins.FactoryID as [Factory]
    , ins.Line
    , ins.Team 
    , ins.Shift
    , ord.custpono as [PO#]
    , ord.styleid as [Style]
    , ins.OrderId as [SP#]
    , ins.Article
    , cast(Ins.AddDate as date) as [First Inspection Date]
    , format(ins.AddDate,'yyyy/MM/dd HH:mm') as [Inspected Time]
    , Inspection_QC.Name as [Inspected QC]
    , [Destination] = Cou.Alias
    , sty.ProductType
    , [Product Type] = 
        case when ins.Location = 'T' then 'TOP'
            when ins.Location = 'B' then 'BOTTOM'
            when ins.Location = 'I' then 'INNER'
            when ins.Location = 'O' then 'OUTER'
            else ''
        end 
    , ins.AddName
    , ins.PassWIP
    , ins.RejectWIP
into #tmp_src
from InlineInspection ins  
inner join sciproduction_orders ord on ins.OrderId=ord.id
inner join SciProduction_Factory fac on ins.FactoryID=fac.ID
left join SciProduction_Style ps on ps.Ukey = ord.StyleUkey
left join InlineInspection_Detail ind on ins.id=ind.InlineInspectionID
left join SciProduction_SewingLine s on s.FactoryID = ins.FactoryID and s.ID = ins.Line
left join SciProduction_Country Cou on ord.Dest = Cou.ID
outer apply(select FirstName, LastName from InlineEmployee with(nolock) where InlineEmployee.EmployeeID= ins.SewerID) InlineEmployee
outer apply(select name from pass1 with(nolock) where pass1.id= ins.AddName) Inspection_QC
outer apply(select name from pass1 with(nolock) where pass1.id= ins.EditName) Inspection_fixQC
Outer apply (
    SELECT ProductType = r2.Name
        , FabricType = r1.Name
        , Lining
        , Gender
        , Construction = d1.Name
    FROM SciProduction_Style s WITH(NOLOCK)
    left join SciProduction_DropDownList d1 WITH(NOLOCK) on d1.type= 'StyleConstruction' and d1.ID = s.Construction
    left join SciProduction_Reason r1 WITH(NOLOCK) on r1.ReasonTypeID= 'Fabric_Kind' and r1.ID = s.FabricType
    left join SciProduction_Reason r2 WITH(NOLOCK) on r2.ReasonTypeID= 'Style_Apparel_Type' and r2.ID = s.ApparelType
    where s.Ukey = ord.StyleUkey
)sty
where ord.StyleUkey=@StyleUkey 
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}

-- summary
select t.[First Inspection Date]
    ,t.Factory
    ,t.Brand
    ,t.Style
    ,t.[PO#]
    ,t.[SP#]
    ,t.Article   
    ,t.[Destination]
    ,t.ProductType
    ,t.Team
    ,AddName=(select stuff(
                            (select distinct concat(',',AddName) 
                               from #tmp_src t2 
                              where 1=1
                                and t2.[First Inspection Date] = t.[First Inspection Date] 
                                and t2.Factory = t.Factory 
                                and t2.[SP#] = t.[SP#] 
                                and t2.Team = t.Team 
                                and t2.Shift = t.Shift 
                                and t2.Line = t.Line for xml path('') ),1,1,'')
                          )
    ,t.Shift
    ,t.Line
    ,t.PassWIP
    ,t.RejectWIP
into #tmp_Summy_first
from #tmp_src t

select t.[First Inspection Date]
      ,t.Factory
      ,t.Brand
      ,t.Style
      ,t.[PO#]
      ,t.[SP#]
      ,t.[Destination]
      ,t.ProductType
      ,t.Team
      ,t.AddName
      ,t.Shift
      ,t.Line
      ,t.Article
      ,[TtlQty]=SUM(t.PassWIP+t.RejectWIP) 
      ,[PassQty] = Sum(t.PassWIP)
      ,[RejectQty] = Sum(t.RejectWIP)
into #tmp_Summy_Final
from 
(select t.[First Inspection Date]
          ,t.Factory
          ,t.Brand
          ,t.Style
          ,t.[PO#]
          ,t.[SP#]
          ,t.[Destination]
          ,t.ProductType
          ,t.Team
          ,t.AddName
          ,t.Shift
          ,t.Line
          ,sum(t.PassWIP) as [PassWIP]
          ,sum(t.RejectWIP) as [RejectWIP]
          ,t.Article
    from #tmp_Summy_first t
    group by t.[First Inspection Date],t.Factory,t.Brand,t.Style,t.[PO#],t.[SP#]
,t.[Destination],t.Team,t.AddName,t.Shift,t.Line, t.Article ,t.ProductType )t
group by t.[First Inspection Date],t.Factory,t.Brand,t.Style,t.[PO#],t.[SP#]
        ,t.[Destination],t.Team,t.AddName,t.Shift,t.Line,t.Article ,t.ProductType

select FirstInspDate = t.[First Inspection Date]
      ,t.Factory
      ,POID = t.[PO#]
      ,OrderID = t.[SP#]
      ,t.Article
      ,t.[Destination]
      ,t.ProductType
      ,QCName = t.AddName
      --,t.TtlQty as [Inspected Qty]
      ,RejectQty = t.RejectQty
      ,RFTRate = ROUND( (t.PassQty *1.0) / (t.TtlQty *1.0) *100 ,2)
from #tmp_Summy_Final t
Order by t.[First Inspection Date], t.Factory, t.Brand, t.[SP#],t.Article, t.Line, t.Team


drop table #tmp_src ,#tmp_Summy_first , #tmp_Summy_Final

";

            var result = ExecuteList<Production_InlineRFT>(CommandType.Text, sqlcmd, objParameter) ?? new List<Production_InlineRFT>();
            return result.ToList();
        }
        
        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public decimal Get_Production_EndlineWFT(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

select *
into #tmpInspection
from Inspection
where StyleUkey=@StyleUkey
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}

select EndlineWFT = ROUND( IIF(COUNT(1) = 0,0, 1.0 * SUM(IIF(Status <>'Pass',1,0)) / COUNT(1)) ,2)
from #tmpInspection

drop table #tmpInspection
";
            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToDecimal(result == null ? 0 : result);
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Production_EndlineWFT> Get_Production_EndlineWFT_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"

--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @Article as varchar(30)='015157'
--DECLARE @StyleUkey as bigint = 89408

select
    ins.InspectionDate
    ,[FirstInspectionDate] = cast(Ins.AddDate as date)
    ,[Factory] = ins.FactoryID
    ,[Brand] = ord.BrandID
    ,[Style] = ord.styleid
    ,[PO#] = ord.custpono
    ,[SP#] = ins.OrderId
    ,ins.Article
    ,[Destination] = Cou.Alias
    ,sty.ProductType
    ,ins.AddName
    ,[PassQty] = SUM(IIF(ins.Status = ('Pass') , 1 ,0))
    ,[RejectAndFixedQty] = SUM(IIF(ins.Status in ('Reject','Fixed','Dispose') , 1 ,0))
    ,[TtlQty] = SUM(IIF(ins.Status = ('Pass') , 1 ,0)) + SUM(IIF(ins.Status in ('Reject','Fixed','Dispose') , 1 ,0))
into #tmp_summy_first
from Inspection ins  
inner join sciproduction_orders ord on ins.OrderId=ord.id
inner join SciProduction_Factory fac on ins.FactoryID=fac.ID
left join SciProduction_Style ps on ps.Ukey = ord.StyleUkey
left join [dbo].[SciProduction_SewingLine] s on s.FactoryID = ins.FactoryID and s.ID = ins.Line
left join [dbo].[SciProduction_Country] Cou on ord.Dest = Cou.ID
Outer apply (
    SELECT ProductType = r2.Name
        , FabricType = r1.Name
        , Lining
        , Gender
        , Construction = d1.Name
    FROM SciProduction_Style s WITH(NOLOCK)
    left join SciProduction_DropDownList d1 WITH(NOLOCK) on d1.type= 'StyleConstruction' and d1.ID = s.Construction
    left join SciProduction_Reason r1 WITH(NOLOCK) on r1.ReasonTypeID= 'Fabric_Kind' and r1.ID = s.FabricType
    left join SciProduction_Reason r2 WITH(NOLOCK) on r2.ReasonTypeID= 'Style_Apparel_Type' and r2.ID = s.ApparelType
    where s.Ukey = ord.StyleUkey
)sty
where 1=1 and ord.StyleUkey=@StyleUkey
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}

group by ins.InspectionDate
	,cast(Ins.AddDate as date)
	,ins.FactoryID
	,ord.BrandID
	,ord.styleid
	,ord.custpono
	,ins.OrderId
	,ins.Article
	,Cou.Alias
	,sty.ProductType
	,ins.AddName

select
    FirstInspDate = t.FirstInspectionDate
    , t.Factory
    , POID = t.[PO#]
    , OrderID = t.[SP#]
    , t.Article
    , t.[Destination]
    , t.ProductType
    , QCName = t.AddName 
    --, t.TtlQty
    , RejectQty = t.RejectAndFixedQty
    , WFTRate = ROUND( iif(t.TtlQty = 0, 0, ROUND( (t.RejectAndFixedQty *1.0) / (t.TtlQty *1.0) *100,3)) ,2)
from #tmp_summy_first t

drop table #tmp_summy_first

";

            var result = ExecuteList<Production_EndlineWFT>(CommandType.Text, sqlcmd, objParameter) ?? new List<Production_EndlineWFT>();
            return result.ToList();
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public decimal Get_Production_MDPassRate(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @StyleUkey as bigint = (select Ukey from Style s where s.BrandID=@BrandID AND s.SeasonID = @SeasonID and s.ID= @StyleID )

SELECT distinct OrderID=o.ID ,o.BrandID ,oq.Seq ,PackingList = p.ID
INTO #Base
FROM Order_QtyShip oq
INNER JOIN Orders o on o.id=oq.id
INNER JOIN PackingList_Detail pd WITH (NOLOCK) on  pd.OrderID = o.ID AND pd.OrderShipmodeSeq = oq.Seq
INNER JOIN PackingList p WITH (NOLOCK)  ON p.ID = pd.ID AND p.Type NOT IN ('S', 'F')
WHERE 1=1
and o.StyleUkey =@StyleUkey
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}

--統計訂單以及 Scan 的數量
select pd.OrderID ,b.BrandID, ShipQty = sum(pd.ShipQty), ScanQty = sum(iif(pd.ScanQty > pd.ShipQty, pd.ShipQty, pd.ScanQty))
into #AllQty
from PackingList_Detail pd with(nolock) 
Inner join #Base b on b.PackingList=pd.ID AND b.OrderID=pd.OrderID and b.Seq=pd.OrderShipmodeSeq
group by pd.OrderID,b.BrandID

--找出第一次MD Fail紀錄
select pt.PackingListID, pt.OrderID ,pt.CTNStartNo ,　MinAddDate = MIN(pt.AddDate)
into #tmp
from PackErrTransfer pt with(nolock) 
Inner join PackingList_Detail pd with(nolock) on pt.PackingListID=pd.ID  and pt.CTNStartNo=pd.CTNStartNo
Inner join #Base b on b.PackingList=pd.ID AND b.OrderID=pd.OrderID and b.Seq=pd.OrderShipmodeSeq
where pt.PackingErrorID='00006'
group by  pt.PackingListID, pt.OrderID ,pt.CTNStartNo

--計算 MD Fail 的數量
select pt.OrderID,ErrQty=SUM(ErrQty)
into #ErrorQty
from PackErrTransfer pt
inner join #tmp t on pt.PackingListID = t.PackingListID AND pt.OrderID=t.OrderID and pt.CTNStartNo=t.CTNStartNo and pt.AddDate=t.MinAddDate
where ErrQty > 0
group by pt.OrderID

--計算Pass Qty
select MDPassRate = ROUND( IIF(SUM( a.ScanQty)=0,0, 1.0* ( SUM(  a.ScanQty  - ISNULL(e.ErrQty,0))) /SUM( a.ScanQty)) ,2)
from #AllQty a
left join #ErrorQty e on a.OrderID=e.OrderID
group by a.BrandID

drop table #Base,#AllQty,#tmp,#ErrorQty
";

            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToDecimal(result == null ? 0 : result);
        }

        /// <summary>
        /// 正式機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<Production_MDPassRate> Get_Production_MDPassRate_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @Article as varchar(30)='015157'
--DECLARE @StyleUkey as bigint = 89408


SELECT distinct OrderID=o.ID ,o.BrandID ,oq.Seq ,PackingList = p.ID
INTO #Base
FROM Order_QtyShip oq
INNER JOIN Orders o on o.id=oq.id
INNER JOIN PackingList_Detail pd WITH (NOLOCK) on  pd.OrderID = o.ID AND pd.OrderShipmodeSeq = oq.Seq
INNER JOIN PackingList p WITH (NOLOCK)  ON p.ID = pd.ID AND p.Type NOT IN ('S', 'F')
WHERE 1=1
 and o.StyleUkey =@StyleUkey
{(string.IsNullOrEmpty(Req.Article) ? string.Empty : "and Article = @Article")}

--統計訂單以及 Scan 的數量
select pd.OrderID ,b.BrandID, b.Seq, ShipQty = sum(pd.ShipQty), ScanQty = sum(iif(pd.ScanQty > pd.ShipQty, pd.ShipQty, pd.ScanQty))
into #AllQty
from PackingList_Detail pd with(nolock) 
Inner join #Base b on b.PackingList=pd.ID AND b.OrderID=pd.OrderID and b.Seq=pd.OrderShipmodeSeq
group by pd.OrderID,b.BrandID, b.Seq

--找出第一次MD Fail紀錄
select pt.PackingListID, pt.OrderID , b.Seq,pt.CTNStartNo ,　MinAddDate = MIN(pt.AddDate)
into #tmp
from PackErrTransfer pt with(nolock) 
Inner join PackingList_Detail pd with(nolock) on pt.PackingListID=pd.ID  and pt.CTNStartNo=pd.CTNStartNo
Inner join #Base b on b.PackingList=pd.ID AND b.OrderID=pd.OrderID and b.Seq=pd.OrderShipmodeSeq
where pt.PackingErrorID='00006'
group by  pt.PackingListID, pt.OrderID , b.Seq,pt.CTNStartNo

--計算 MD Fail 的數量
select pt.OrderID,ErrQty=SUM(ErrQty)
into #ErrorQty
from PackErrTransfer pt
inner join #tmp t on pt.PackingListID = t.PackingListID AND pt.OrderID=t.OrderID and pt.CTNStartNo=t.CTNStartNo and pt.AddDate=t.MinAddDate
where ErrQty > 0
group by pt.OrderID

--計算Pass Qty
select OrderID = o.ID
    ,Article.Articles   
    ,DeliveryDate=o.BuyerDelivery
    ,MDFailQty = SUM( ISNULL(e.ErrQty,0) )
from #AllQty a
inner join Order_QtyShip oq on oq.ID = a.OrderID and oq.Seq = a.Seq
inner join Orders o on a.OrderID = o.ID
left join #ErrorQty e on a.OrderID=e.OrderID
outer apply( SELECT
        Articles = STUFF((SELECT DISTINCT
                CONCAT(',', oqd.Article)
            FROM Order_QtyShip_Detail oqd WITH (NOLOCK)
            WHERE oqd.Id = o.ID and oqd.Seq = a.Seq 
            FOR XML PATH (''))
        , 1, 1, '')
)Article
--where  a.ScanQty > 0 
--and e.ErrQty <> 0  
group by o.ID
    ,a.Seq
    ,o.StyleID
    ,o.SeasonID
    ,o.BuyerDelivery
    ,Article.Articles   
    ,o.SewLine 
    ,o.BrandID
    ,oq.Qty

drop table #Base,#AllQty,#tmp,#ErrorQty

";

            var result = ExecuteList<Production_MDPassRate>(CommandType.Text, sqlcmd, objParameter) ?? new List<Production_MDPassRate>();
            return result.ToList();
        }

        /// <summary>
        /// 備機執行
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public List<FinalInspection> Get_FinalInspection_Detail(EtoEFlowChart_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@StyleUkey", DbType.Int64, Req.StyleUkey } ,
            };
            string sqlcmd = $@"
--DECLARE @StyleID as varchar(30)='ARWCS22114A'
--DECLARE @BrandID as varchar(30)='REEBOK'
--DECLARE @SeasonID as varchar(30)='22SS'
--DECLARE @Article as varchar(30)='015157'
--DECLARE @StyleUkey as bigint = 92336

select PassRate = ROUND( (
    select IIF(SUM( fi.SampleSize)=0,0,1.0 * SUM( fi.PassQty) / SUM( fi.SampleSize))
    from Finalinspection fi with (nolock)
    where 1=1
    and fi.InspectionStage ='Final'
    and fi.InspectionResult != 'On-going'
    and fi.SubmitDate is not null
    AND exists(
        select 1 
        from FinalInspection_Order a
        inner join SciProduction_Orders b on a.OrderID=b.ID
        where a.ID = fi.ID AND b.StyleUkey = @StyleUkey
    )
) ,2)
, SQR = ROUND( (
    select IIF(SUM( fi.SampleSize)=0,0,1.0 * SUM( fi.RejectQty) / SUM( fi.SampleSize))
    from Finalinspection fi with (nolock)
    where 1=1
    and fi.InspectionStage ='Final'
    and fi.InspectionResult != 'Fail'
    and fi.SubmitDate is not null
    AND exists(
        select 1 
        from FinalInspection_Order a
        inner join SciProduction_Orders b on a.OrderID=b.ID
        where a.ID = fi.ID AND b.StyleUkey = @StyleUkey
    )
) ,2)
, ChinaPassRate = ROUND( (
    select ISNULL( IIF(SUM( fi.SampleSize)=0,0,1.0 * SUM( fi.PassQty) / SUM( fi.SampleSize)) ,0)
    from Finalinspection fi with (nolock)
    where 1=1
    and fi.InspectionStage ='Final'
    and fi.InspectionResult != 'On-going'
    and fi.SubmitDate is not null
    AND exists(
        select 1 
        from FinalInspection_Order a
        inner join SciProduction_Orders b on a.OrderID=b.ID
        where a.ID = fi.ID AND b.StyleUkey = @StyleUkey
        and b.Dest = 'CN'

    )
) ,2)
,JapanPassRate = ROUND( (
    select ISNULL( IIF(SUM( fi.SampleSize)=0,0,1.0 * SUM( fi.PassQty) / SUM( fi.SampleSize)),0)
    from Finalinspection fi with (nolock)
    where 1=1
    and fi.InspectionStage ='Final'
    and fi.InspectionResult != 'On-going'
    and fi.SubmitDate is not null
    AND exists(
        select 1 
        from FinalInspection_Order a
        inner join SciProduction_Orders b on a.OrderID=b.ID
        where a.ID = fi.ID AND b.StyleUkey = @StyleUkey
        and b.Dest = 'JP'

    )
) ,2)

";

            var result = ExecuteList<FinalInspection>(CommandType.Text, sqlcmd, objParameter) ?? new List<FinalInspection>();
            return result.ToList();
        }

    }
}

