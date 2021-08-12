using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using ManufacturingExecutionDataAccessLayer.Interface;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class MailToProvider : SQLDAL, IMailToProvider
    {
        #region 底層連線
        public MailToProvider(string ConString) : base(ConString) { }
        public MailToProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<MailTo> Get(MailTo Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ID", DbType.String, Item.ID }
            };

            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,Description" + Environment.NewLine);
            SbSql.Append("        ,ToAddress" + Environment.NewLine);
            SbSql.Append("        ,CcAddress" + Environment.NewLine);
            SbSql.Append("        ,Subject" + Environment.NewLine);
            SbSql.Append("        ,Content" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("FROM [MailTo]" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ID)) { SbSql.Append(" and ID = @ID" + Environment.NewLine); }

            return ExecuteList<MailTo>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
