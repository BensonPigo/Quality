using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using DatabaseObject.ProductionDB;
using ADOHelper.Utility;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class FactoryProvider : SQLDAL, IFactoryProvider
    {
        #region 底層連線
        public FactoryProvider(string ConString) : base(ConString) { }
        public FactoryProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<Factory> GetFtyGroup()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            SbSql.Append(
                @"
select distinct FTYGroup
from Factory WITH(NOLOCK)
where Junk = 0
and IsProduceFty = 1
" + Environment.NewLine);

            return ExecuteList<Factory>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<Factory> GetMDivisionID(string fty)
        {
            string sqlcmd = $@"
select MDivisionID from Factory  WITH(NOLOCK)
where ID = '{fty}'
";

            return ExecuteList<Factory>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }
    }
}
