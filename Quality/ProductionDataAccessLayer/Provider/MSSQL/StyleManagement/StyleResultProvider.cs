using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.SampleRFT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Provider.MSSQL.StyleManagement
{
    public class StyleResultProvider : SQLDAL
    {
        #region 底層連線
        public StyleResultProvider(string ConString) : base(ConString) { }
        public StyleResultProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion
        public IList<SelectListItem> GetBrands()
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
SELECT DISTINCT [Text] = ID, [Value]= ID
FROM Production.dbo.Brand with (nolock)
WHERE Junk=0
");

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public IList<SelectListItem> GetSeasons(string brandID = "")
        {
            StringBuilder SbSql = new StringBuilder();

            string where = string.IsNullOrEmpty(brandID) ? string.Empty : $" and BrandID = '{brandID}'";

            SbSql.Append($@"
SELECT distinct [Text] = ID, [Value]= ID
FROM Production.dbo.Season with (nolock)
WHERE Junk=0 {where}
");

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public IList<StyleResult_Request> GetStyle(StyleResult_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection listPar = new SQLParameterCollection();

            SbSql.Append($@"
SELECT StyleID=ID,StyleUkey=Cast(Ukey AS VARCHAR),BrandID,SeasonID
FROM Production.dbo.Style with (nolock)
WHERE Junk=0
");
            if (!string.IsNullOrWhiteSpace(Req.StyleUkey))
            {
                SbSql.Append($@"and Ukey = {Req.StyleUkey} ");
            }

            if (!string.IsNullOrWhiteSpace(Req.StyleID))
            {
                SbSql.Append($@"and ID = '{Req.StyleUkey}' ");
            }
            if (!string.IsNullOrWhiteSpace(Req.BrandID))
            {
                SbSql.Append($@"and BrandID = '{Req.BrandID}' ");
            }
            if (!string.IsNullOrWhiteSpace(Req.SeasonID))
            {
                SbSql.Append($@"and SeasonID = '{Req.SeasonID}' ");
            }

            return ExecuteList<StyleResult_Request>(CommandType.Text, SbSql.ToString(), listPar);
        }
        public IList<SelectListItem> GetStyles(StyleResult_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection listPar = new SQLParameterCollection();

            SbSql.Append($@"
SELECT  distinct [Text] = ID, [Value]= ID
FROM Production.dbo.Style with (nolock)
WHERE Junk=0
");
            if (!string.IsNullOrWhiteSpace(Req.StyleUkey))
            {
                SbSql.Append($@"and Ukey = {Req.StyleUkey} ");
            }

            if (!string.IsNullOrWhiteSpace(Req.StyleID))
            {
                SbSql.Append($@"and ID = '{Req.StyleUkey}' ");
            }
            if (!string.IsNullOrWhiteSpace(Req.BrandID))
            {
                SbSql.Append($@"and BrandID = '{Req.BrandID}' ");
            }
            if (!string.IsNullOrWhiteSpace(Req.SeasonID))
            {
                SbSql.Append($@"and SeasonID = '{Req.SeasonID}' ");
            }

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), listPar);
        }

        public IList<StyleResult_ViewModel> Get_StyleInfo(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;

            if (!string.IsNullOrEmpty(styleResult_Request.BrandID))
            {
                sqlWhere += " and s.BrandID = @BrandID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.SeasonID))
            {
                sqlWhere += " and s.SeasonID = @SeasonID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.StyleID))
            {
                sqlWhere += " and s.ID = @StyleID";
            }

            listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
            listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));

            switch (styleResult_Request.CallType)
            {
                case StyleResult_Request.EnumCallType.PrintBarcode:
                    sqlCol = @"
        [StyleUkey] = cast(s.Ukey as varchar),
        [StyleID] = s.ID,
        s.BrandID,
        s.SeasonID,
        s.ProgramID,
        s.Description,
        [ProductType] = (   select  TOP 1 Name
							from Reason  WITH(NOLOCK)
							where ReasonTypeID = 'Style_Apparel_Type' and ID = s.ApparelType
                        ),
        s.StyleName";
                    break;
                case StyleResult_Request.EnumCallType.StyleResult:
                    sqlCol = @"
        [StyleUkey] = cast(s.Ukey as varchar),
        [StyleID] = s.ID,
        s.BrandID,
        s.SeasonID,
        s.ProgramID,
        s.Description,
        [ProductType] = (   select  TOP 1 Name
							from Reason  WITH(NOLOCK)
							where ReasonTypeID = 'Style_Apparel_Type' and ID = s.ApparelType
                        ),
        [Article] = (   Stuff((
                                Select concat( ',',Article)
                                From Style_Article with (nolock)
                                Where StyleUkey = s.Ukey
                                Order by Seq FOR XML PATH('')
                            ),1,1,'') 
                    ),
        s.StyleName,
        [SpecialMark] = (select Name 
                        from Reason WITH (NOLOCK) 
                        where   ReasonTypeID = 'Style_SpecialMark' and
                        	    ID = s.SpecialMark
                        ),
        [SMR] = (select Concat (ID, ' ', Name)
                    from   TPEPass1  with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkSMR, s.SampleSMR)
                ),
        Handle = (select Concat (ID, ' ', Name)
                    from   TPEPass1  with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkMRHandle, s.SampleMRHandle)
                ),
        [RFT] = LEFT( CAST( RFT.Val as varchar),5) ";
                    break;
                default:
                    break;
            }

            string sqlGet_StyleResult_Browse = $@"
select  {sqlCol}
    ,StyleRRLRPath = (select StyleRRLRPath from System WITH(NOLOCK))
from    Style s with (nolock)
OUTER APPLY(
	select Val = ROUND( SUM(IIF(Status = 'Pass',1,0)) * 1.0  / COUNT(1) *1.0  *100 , 2)
    FROM [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection  rft WITH(NOLOCK)
	WHERE rft.StyleUkey = s.Ukey
)RFT
where   1 = 1 {sqlWhere}
";
            return ExecuteList<StyleResult_ViewModel>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_SampleRFT> Get_StyleResult_SampleRFT(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;

            if (!string.IsNullOrEmpty(styleResult_Request.BrandID))
            {
                sqlWhere += " and s.BrandID = @BrandID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.SeasonID))
            {
                sqlWhere += " and s.SeasonID = @SeasonID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.StyleID))
            {
                sqlWhere += " and s.ID = @StyleID";
            }

            listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
            listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));

            string sqlGet_StyleResult_Browse = $@"
select SP = o.ID
	,SampleStage = o.OrderTypeID
	,Factory = o.FactoryID
	,Delivery = o.BuyerDelivery
	,o.SCIDelivery
	,InspectedQty = Inspected.val
	,RFT = Cast( Cast( IIF(Inspected.val = 0 , 0 , ROUND( ( RFT.val * 1.0 / Inspected.val ) * 100 ,2) ) as numeric(5,2)) as varchar )
	,BAProduct = BAProduct.val
	,BAAuditCriteria  = Cast( Cast( IIF(Inspected.val = 0, 0, ROUND(BAProduct.val * 1.0 / Inspected.val * 5 ,1) )as numeric(2,1)) as varchar )
from Orders o WITH(NOLOCK)
inner join Style s WITH(NOLOCK) on s.ID = o.StyleID AND o.BrandID = s.BrandID AND o.SeasonID = s.SeasonID
outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID
)Inspected

outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID and r.Status='Pass'
)RFT

outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID 
	AND (
		NOT EXISTS(--沒有 RFT_Inspeciton_Detail
			select 1
			from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection_Detail rd WITH(NOLOCK) where r.ID = rd.ID	
		)
		or NOT EXISTS( --RFT_Inspection_Detail 所有資料 PMS_RFTACriterialID 皆為空
			select 1
			from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection_Detail rd WITH(NOLOCK) where r.ID = rd.ID AND rd.PMS_RFTBACriteriaID != ''
		)
	)
)BAProduct
where o.Category = 'S'
{sqlWhere}
";
            return ExecuteList<StyleResult_SampleRFT>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }


        public DataTable Get_StyleResult_SampleRFT_DataTable(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;
            if (!string.IsNullOrEmpty(styleResult_Request.StyleUkey))
            {
                sqlWhere += " and s.Ukey = @StyleUkey";

                int Ukey = 0;
                if (int.TryParse(styleResult_Request.StyleUkey, out Ukey))
                {
                    listPar.Add("@StyleUkey", DbType.Int64, Ukey);
                }
                else
                {
                    listPar.Add("@StyleUkey", DbType.Int64, 0);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(styleResult_Request.BrandID))
                {
                    sqlWhere += " and s.BrandID = @BrandID";
                }

                if (!string.IsNullOrEmpty(styleResult_Request.SeasonID))
                {
                    sqlWhere += " and s.SeasonID = @SeasonID";
                }

                if (!string.IsNullOrEmpty(styleResult_Request.StyleID))
                {
                    sqlWhere += " and s.ID = @StyleID";
                }

                listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
                listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
                listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));
            }

            string SbSql = $@"
select SP = o.ID
	,SampleStage = o.OrderTypeID
	,Factory = o.FactoryID
	,Delivery = o.BuyerDelivery
	,o.SCIDelivery
	,InspectedQty = Inspected.val
	,RFT = Cast( Cast( IIF(Inspected.val = 0 , 0 , ROUND( ( RFT.val * 1.0 / Inspected.val ) * 100 ,2) ) as numeric(5,2)) as varchar )
	,BAProduct = BAProduct.val
	,BAAuditCriteria  = Cast( Cast( IIF(Inspected.val = 0, 0, ROUND(BAProduct.val * 1.0 / Inspected.val * 5 ,1) )as numeric(2,1)) as varchar )
from Orders o WITH(NOLOCK)
inner join Style s WITH(NOLOCK) on s.ID = o.StyleID AND o.BrandID = s.BrandID AND o.SeasonID = s.SeasonID
outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID
)Inspected

outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID and r.Status='Pass'
)RFT

outer apply(
	select val = COUNT(r.ID)
	from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection r WITH(NOLOCK)
	where r.OrderID = o.ID 
	AND (
		NOT EXISTS(--沒有 RFT_Inspeciton_Detail
			select 1
			from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection_Detail rd WITH(NOLOCK) where r.ID = rd.ID	
		)
		or NOT EXISTS( --RFT_Inspection_Detail 所有資料 PMS_RFTACriterialID 皆為空
			select 1
			from [ExtendServer].ManufacturingExecution.dbo.RFT_Inspection_Detail rd WITH(NOLOCK) where r.ID = rd.ID AND rd.PMS_RFTBACriteriaID != ''
		)
	)
)BAProduct
where 1=1
{sqlWhere}
";
            return ExecuteDataTable(CommandType.Text, SbSql.ToString(), listPar);
        }

        public IList<StyleResult_FTYDisclaimer> Get_StyleResult_FTYDisclaimer(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;
            if (!string.IsNullOrEmpty(styleResult_Request.BrandID))
            {
                sqlWhere += " and s.BrandID = @BrandID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.SeasonID))
            {
                sqlWhere += " and s.SeasonID = @SeasonID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.StyleID))
            {
                sqlWhere += " and s.ID = @StyleID";
            }

            listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
            listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));

            string sqlGet_StyleResult_Browse = $@"
select  ExpectionFormStatus = d.Name
,s.ExpectionFormDate
,s.ExpectionFormRemark
,sa.Article
,sa.Description
,FDFilePath = IIF(sa.SourceFile = null OR sa.SourceFile = '' 
				, '' 
				,(select StyleFDFilePath +  sa.SourceFile from System WITH(NOLOCK))
			)
,FDFileName = ISNULL(sa.SourceFile ,'')
from Style s WITH(NOLOCK)
left join DropDownList d WITH(NOLOCK) on d.Type = 'FactoryDisclaimer' AND s.ExpectionFormStatus = d.ID
left join Style_Article sa WITH(NOLOCK) on s.Ukey = sa.StyleUkey
where 1=1
{sqlWhere}
";
            return ExecuteList<StyleResult_FTYDisclaimer>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_RRLR> Get_StyleResult_RRLR(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;
            if (!string.IsNullOrEmpty(styleResult_Request.BrandID))
            {
                sqlWhere += " and s.BrandID = @BrandID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.SeasonID))
            {
                sqlWhere += " and s.SeasonID = @SeasonID";
            }

            if (!string.IsNullOrEmpty(styleResult_Request.StyleID))
            {
                sqlWhere += " and s.ID = @StyleID";
            }

            listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
            listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));

            string sqlGet_StyleResult_Browse = $@"
select sr.Refno
	,Supplier = sr.SuppID + '-' +su.AbbEN
	,sr.ColorID
	,sr.RR
	,Remark = sr.RRRemark
	,sr.LR
from Style s WITH(NOLOCK)
inner join Style_RRLR_Report sr WITH(NOLOCK) on s.Ukey = sr.StyleUkey
left join Supp su WITH(NOLOCK) ON sr.SuppID = su.ID
where 1=1
{sqlWhere}
";
            return ExecuteList<StyleResult_RRLR>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_BulkFGT> Get_StyleResult_BulkFGT(StyleResult_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlWhere = string.Empty;
            string sqlCol = string.Empty;
            listPar.Add(new SqlParameter("@StyleID", styleResult_Request.StyleID));
            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));
            listPar.Add(new SqlParameter("@SeasonID", styleResult_Request.SeasonID));
            listPar.Add(new SqlParameter("@MDivisionID", styleResult_Request.MDivisionID));

            string sqlGet_StyleResult_Browse = $@"
--每個Articl都會要有的Type，先準備好
SELECT [Type] = '451'
INTO #Type
UNION
SELECT [Type] = IIF( EXISTS(
	select SpecialMark,r.Name
	from Style s WITH(NOLOCK)
	inner join Reason r WITH(NOLOCK) on s.SpecialMark = r.ID AND r.ReasonTypeID= 'Style_SpecialMark'
	where s.ID = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
	AND r.Name IN (
		'MATCH TEAMWEAR',
		'BASEBALL ON FIELD',
		'SOFTBALL ON FIELD',
		'TRAINING TEAMWEAR',
		'LACROSSE ONFIELD',
		'AMERIC. FOOT. ON-FIELD',
		'TIRO',
		'BASEBALL OFF FIELD',
		'NCAA ON ICE',
		'ON-COURT',
		'BBALL PERFORMANCE',
		'BRANDED BLANKS',
		'SLD ON-FIELD',
		'NHL ON ICE',
		'SLD ON-COURT'
	)
),'710','701')


--Type 450一個Style只會出現一次
select Article='', Type='450', TestName = 'Seam Breakage'
,LastResult=(
	select CASE WHEN g.SeamBreakageResult = 'P' THEN 'Pass'
				WHEN g.SeamBreakageResult = 'F' THEN 'Fail'
				ELSE ''
			END
	from GarmentTest g WITH(NOLOCK)
	WHERE g.StyleID = @StyleID
		AND g.BrandID = @BrandID
		AND g.SeasonID = @SeasonID
		AND g.MDivisionid = @MDivisionID
		AND g.SeamBreakageLastTestDate = (		
			select MAX(SeamBreakageLastTestDate)
			from GarmentTest gg WITH(NOLOCK)
			where gg.StyleID = g.StyleID
				AND gg.BrandID = g.BrandID
				AND gg.SeasonID = g.SeasonID
				AND gg.MDivisionid= g.MDivisionID
		)
)
,LastTestDate=(
	select g.SeamBreakageLastTestDate 
	from GarmentTest g WITH(NOLOCK)
	WHERE g.StyleID = @StyleID
		AND g.BrandID = @BrandID
		AND g.SeasonID = @SeasonID
		AND g.MDivisionid = @MDivisionID
		AND g.SeamBreakageLastTestDate = (		
			select MAX(SeamBreakageLastTestDate)
			from GarmentTest gg WITH(NOLOCK)
			where gg.StyleID = g.StyleID
				AND gg.BrandID = g.BrandID
				AND gg.SeasonID = g.SeasonID
				AND gg.MDivisionid= g.MDivisionID
		)
)
UNION
select sa.Article, t.Type
,TestName = CASE    WHEN t.Type = '451' THEN 'Odour'
					WHEN  t.Type = '701' THEN 'Garment Wash'
					WHEN  t.Type = '710' THEN 'Team Wear Wash Test'
					ELSE ''
			END
,LastResult = CASE  WHEN t.Type = '451' THEN (
	                                                SELECT CASE WHEN Type451.OdourResult = 'P' THEN 'Pass'
				                                                WHEN Type451.OdourResult = 'F' THEN 'Fail'
				                                                ELSE ''
			                                                END
                                                )
					WHEN  t.Type = '701' OR t.Type = '710' THEN (
	                                                                SELECT CASE WHEN Type701_710.WashResult = 'P' THEN 'Pass'
				                                                                WHEN Type701_710.WashResult = 'F' THEN 'Fail'
				                                                                ELSE ''
			                                                                END
                                                                )
					ELSE ''
			END
,LastTestDate= CASE   WHEN t.Type = '451' THEN Type451.Date
					  WHEN  t.Type = '701' OR t.Type = '710' THEN Type701_710.Date
					  ELSE ''
			   END
from Style_Article sa WITH(NOLOCK)
inner join Style s WITH(NOLOCK) ON s.Ukey = sa.StyleUkey
OUTER APPLY(
	select * from #Type
)t
OUTER APPLY(
	select g.OdourResult ,g.Date
	from GarmentTest g WITH(NOLOCK)
	WHERE g.StyleID = s.ID
		AND g.BrandID = s.BrandID
		AND g.SeasonID = s.SeasonID
		AND g.Article = sa.Article
		AND g.MDivisionid = @MDivisionID
		AND g.Date = (		
			select MAX(Date)
			from GarmentTest gg WITH(NOLOCK)
			where gg.StyleID = s.ID
				AND gg.BrandID = s.BrandID
				AND gg.SeasonID = s.SeasonID
				AND gg.Article = sa.Article
				AND gg.MDivisionid = g.MDivisionID
		)
)Type451
OUTER APPLY(
	select g.WashResult ,g.Date
	from GarmentTest g WITH(NOLOCK)
	WHERE g.StyleID = s.ID
		AND g.BrandID = s.BrandID
		AND g.SeasonID = s.SeasonID
		AND g.Article = sa.Article
		AND g.MDivisionid = @MDivisionID
		AND g.Date = (		
			select MAX(Date)
			from GarmentTest gg WITH(NOLOCK)
			where gg.StyleID = s.ID
				AND gg.BrandID = s.BrandID
				AND gg.SeasonID = s.SeasonID
				AND gg.Article = sa.Article
				AND gg.MDivisionid = g.MDivisionID
		)
)Type701_710
where s.ID = @StyleID
AND s.BrandID = @BrandID
AND s.SeasonID = @SeasonID
";
            return ExecuteList<StyleResult_BulkFGT>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }
    }
}
