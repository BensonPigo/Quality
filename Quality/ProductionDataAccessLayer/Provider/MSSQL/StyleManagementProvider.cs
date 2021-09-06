using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class StyleManagementProvider : SQLDAL, IStyleManagementProvider
    {
        #region 底層連線
        public StyleManagementProvider(string ConString) : base(ConString) { }
        public StyleManagementProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SelectListItem> GetBrands()
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
SELECT DISTINCT [Text] = ID, [Value]= ID
FROM Brand with (nolock)
WHERE Junk=0
");

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public IList<SelectListItem> GetSeasons(string brandID)
        {
            StringBuilder SbSql = new StringBuilder();

            string where = string.IsNullOrEmpty(brandID) ? string.Empty : $" and BrandID = '{brandID}'";

            SbSql.Append($@"
SELECT distinct [Text] = ID, [Value]= ID
FROM Season with (nolock)
WHERE Junk=0 {where}
");

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

        public IList<SelectListItem> GetStyles(StyleManagement_Request Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection listPar = new SQLParameterCollection();

            SbSql.Append($@"
SELECT  distinct [Text] = ID, [Value]= ID
FROM Style with (nolock)
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

        public IList<StyleResult_ViewModel> Get_StyleInfo(StyleManagement_Request styleResult_Request)
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
                case StyleManagement_Request.EnumCallType.PrintBarcode:
                    sqlCol = @"
        [StyleUkey] = cast(s.Ukey as varchar),
        [StyleID] = s.ID,
        s.BrandID,
        s.SeasonID,
        s.ProgramID,
        s.Description,
        [ProductType] = (   select  TOP 1 Name
							from Reason 
							where ReasonTypeID = 'Style_Apparel_Type' and ID = s.ApparelType
                        ),
        s.StyleName";
                    break;
                case StyleManagement_Request.EnumCallType.StyleResult:
                    sqlCol = @"
        [StyleUkey] = cast(s.Ukey as varchar),
        [StyleID] = s.ID,
        s.BrandID,
        s.SeasonID,
        s.ProgramID,
        s.Description,
        [ProductType] = (   select  TOP 1 Name
							from Reason 
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
                    from   pass1 with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkSMR, s.SampleSMR)
                ),
        Handle = (select Concat (ID, ' ', Name)
                    from    pass1 with (nolock)
                    where   ID = iif(s.Phase = 'Bulk', s.BulkMRHandle, s.SampleMRHandle)
                ),
        [RFT] = ''";
                    break;
                default:
                    break;
            }

            string sqlGet_StyleResult_Browse = $@"
select  {sqlCol}
from    Style s with (nolock)
where   1 = 1 {sqlWhere}
";
            return ExecuteList<StyleResult_ViewModel>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
        }
    }
}
