using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using ProductionDataAccessLayer.Interface;
using Sci.Win.Tools;
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
    public class PrintBarcodeProvider : SQLDAL
    {
        #region 底層連線
        public PrintBarcodeProvider(string ConString) : base(ConString) { }
        public PrintBarcodeProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<PrintBarcode_Detail> Get_StyleInfo(StyleManagement_Request styleResult_Request)
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
                    listPar.Add("@StyleUkey",  Ukey);
                }
                else
                {
                    listPar.Add("@StyleUkey",  0);
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
							from Reason  WITH(NOLOCK)
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
            var tmp = ExecuteList<PrintBarcode_Detail>(CommandType.Text, sqlGet_StyleResult_Browse, listPar);
            return tmp.Any() ? tmp.ToList() : new List<PrintBarcode_Detail>();
        }

        public List<SelectListItem> Get_SampleStage(StyleManagement_Request styleResult_Request)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            string sqlCol = $@"
select DISTINCT Text=ID, Value = Cast( SerialKey as varchar(10) )
from OrderType
where junk =0 
and Category IN( 'S')
and BrandID = @BrandID
";

            listPar.Add(new SqlParameter("@BrandID", styleResult_Request.BrandID));

            var tmp = ExecuteList<SelectListItem>(CommandType.Text, sqlCol, listPar);
            return tmp.Any() ? tmp.ToList() : new List<SelectListItem>();
        }
    }
}
