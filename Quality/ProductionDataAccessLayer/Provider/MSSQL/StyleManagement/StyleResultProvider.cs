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
            if (!string.IsNullOrEmpty(styleResult_Request.StyleUkey))
            {
                sqlWhere += " and s.Ukey = @StyleUkey";
                //listPar.Add(new SqlParameter("@StyleUkey", styleResult_Request.StyleUkey));

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
							from Production.dbo.Reason 
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
							from Production.dbo.Reason 
							where ReasonTypeID = 'Style_Apparel_Type' and ID = s.ApparelType
                        ),
        [Article] = (   Stuff((
                                Select concat( ',',Article)
                                From Production.dbo.Style_Article with (nolock)
                                Where StyleUkey = s.Ukey
                                Order by Seq FOR XML PATH('')
                            ),1,1,'') 
                    ),
        s.StyleName,
        [SpecialMark] = (select Name 
                        from Production.dbo.Reason WITH (NOLOCK) 
                        where   ReasonTypeID = 'Style_SpecialMark' and
                        	    ID = s.SpecialMark
                        ),
        [SMR] = (select Concat (ID, ' ', Name)
                    from   Production.dbo.pass1 with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkSMR, s.SampleSMR)
                ),
        Handle = (select Concat (ID, ' ', Name)
                    from   Production.dbo.pass1 with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkMRHandle, s.SampleMRHandle)
                ),
        [RFT] = ''";
                    break;
                default:
                    break;
            }

            string sqlGet_StyleResult_Browse = $@"
select  {sqlCol}
from    Production.dbo.Style s with (nolock)
where   1 = 1 {sqlWhere}
";
            return ExecuteList<StyleResult_ViewModel>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_SampleRFT> Get_StyleResult_SampleRFT(StyleResult_Request styleResult_Request)
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

            string sqlGet_StyleResult_Browse = $@"
select SP = o.ID
	,SampleStage = o.OrderTypeID
	,Factory = o.FactoryID
	,Delivery = o.BuyerDelivery
	,o.SCIDelivery
	,InspectedQty = Inspected.val
	,RFT = Cast( Cast( IIF(Inspected.val = 0 , 0 , ROUND( ( RFT.val * 1.0 / Inspected.val ) * 100 ,2) ) as numeric(5,2)) as varchar )
	,BAProduct = BAProduct.val
	,BACriteria  = Cast( Cast( IIF(Inspected.val = 0, 0, ROUND(BAProduct.val * 1.0 / Inspected.val * 5 ,1) )as numeric(2,1)) as varchar )
from Orders o
inner join Style s on s.ID = o.StyleID
outer apply(
	select val = COUNT(r.ID)
	from ManufacturingExecution.dbo.RFT_Inspection r
	where r.OrderID = o.ID
)Inspected

outer apply(
	select val = COUNT(r.ID)
	from ManufacturingExecution.dbo.RFT_Inspection r
	where r.OrderID = o.ID and r.Status='Pass'
)RFT

outer apply(
	select val = COUNT(r.ID)
	from ManufacturingExecution.dbo.RFT_Inspection r
	where r.OrderID = o.ID 
	AND (
		NOT EXISTS(--沒有 RFT_Inspeciton_Detail
			select 1
			from ManufacturingExecution.dbo.RFT_Inspection_Detail rd where r.ID = rd.ID	
		)
		or NOT EXISTS( --RFT_Inspection_Detail 所有資料 PMS_RFTACriterialID 皆為空
			select 1
			from ManufacturingExecution.dbo.RFT_Inspection_Detail rd where r.ID = rd.ID AND rd.PMS_RFTBACriteriaID != ''
		)
	)
)BAProduct
where 1=1
{sqlWhere}
";
            return ExecuteList<StyleResult_SampleRFT>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_FTYDisclamier> Get_StyleResult_FTYDisclamier(StyleResult_Request styleResult_Request)
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

            string sqlGet_StyleResult_Browse = $@"
select  ExpectionFormStatus = d.Name
,s.ExpectionFormDate
,s.ExpectionFormRemark
,sa.Article
,sa.Description
,FDFile = IIF(sa.SourceFile = null OR sa.SourceFile = '' 
				, '' 
				,(select StyleFDFilePath +  sa.SourceFile from System)
			)
from Style s
inner join DropDownList d on d.Type = 'FactoryDisclaimer' AND s.ExpectionFormStatus = d.ID
inner join Style_Article sa on s.Ukey = sa.StyleUkey
where 1=1
{sqlWhere}
";
            return ExecuteList<StyleResult_FTYDisclamier>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }

        public IList<StyleResult_RRLR> Get_StyleResult_RRLR(StyleResult_Request styleResult_Request)
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

            string sqlGet_StyleResult_Browse = $@"
select sr.Refno
	,Supplier = sr.SuppID + '-' +su.AbbEN
	,sr.ColorID
	,sr.RR
	,Remark = sr.RRRemark
	,sr.LR
from Style s
inner join Style_RRLRReport sr on s.Ukey = sr.StyleUkey
left join Supp su ON sr.SuppID = su.ID
where 1=1
{sqlWhere}
";
            return ExecuteList<StyleResult_RRLR>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }
    }
}
