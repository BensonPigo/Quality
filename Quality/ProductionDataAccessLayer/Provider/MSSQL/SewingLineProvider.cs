using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class SewingLineProvider : SQLDAL, ISewingLineProvider
    {
        #region 底層連線
        public SewingLineProvider(string ConString) : base(ConString) { }
        public SewingLineProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SewingLine> GetSewinglineID()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            SbSql.Append(
                @"
select distinct ID
from Sewingline WITH(NOLOCK)
where Junk = 0
" + Environment.NewLine);

            return ExecuteList<SewingLine>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
