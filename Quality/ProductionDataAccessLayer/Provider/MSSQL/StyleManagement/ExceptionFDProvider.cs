using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ViewModel.StyleManagement;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Provider.MSSQL.StyleManagement
{
    public class ExceptionFDProvider : SQLDAL
    {
        #region 底層連線
        public ExceptionFDProvider(string ConString) : base(ConString) { }
        public ExceptionFDProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public IList<ExceptionFD_ViewModel> GetData()
        {
            StringBuilder SbSql = new StringBuilder();

            SbSql.Append($@"
select  
	StyleID = s.ID
	,BrandID
	,SeasonID
	,Article = Article.Val
	,ExpectionFormStatus = d.Name
	,s.ExpectionFormDate
	,s.ExpectionFormRemark
from Style s
left join DropDownList d ON d.Type = 'FactoryDisclaimer' AND s.ExpectionFormStatus = d.ID
outer apply(
	select Val = STUFF((
		select DISTINCt ',' + Article
		from Style_Article sa
		where sa.StyleUkey = s.Ukey
		for xml path('')
		),1,1,'')
)Article
where s.ExpectionFormDate >= DATEADD(Year,-2, GETDATE())
");

            return ExecuteList<ExceptionFD_ViewModel>(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
        }

		public DataTable GetDataTable(string brandID = "")
		{
			StringBuilder SbSql = new StringBuilder();

			SbSql.Append($@"
select  
	StyleID = s.ID
	,BrandID
	,SeasonID
	,Article = Article.Val
	,ExpectionFormStatus = d.Name
	,s.ExpectionFormDate
	,s.ExpectionFormRemark
from Style s
left join DropDownList d ON d.Type = 'FactoryDisclaimer' AND s.ExpectionFormStatus = d.ID
outer apply(
	select Val = STUFF((
		select DISTINCt ',' + Article
		from Style_Article sa
		where sa.StyleUkey = s.Ukey
		for xml path('')
		),1,1,'')
)Article
where s.ExpectionFormDate >= DATEADD(Year,-2, GETDATE())
");

			return ExecuteDataTable(CommandType.Text, SbSql.ToString(), new SQLParameterCollection());
		}

	}
}
