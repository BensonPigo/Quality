using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class QualityBrandTestCodeProvider : SQLDAL
    {
        #region 底層連線
        public QualityBrandTestCodeProvider(string conString) : base(conString) { }
        public QualityBrandTestCodeProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<QualityBrandTestCode> Get(string brandId ,string functuionName)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            SbSql.Append($@"
SELECT * FROM QualityBrandTestCode
WHERE  BrandID = @BrandID AND FunctionName = @FunctionName
" + Environment.NewLine);

            objParameter.Add("@BrandID", DbType.String, brandId ?? string.Empty);
            objParameter.Add("@FunctionName", DbType.String, functuionName ?? string.Empty);

            return ExecuteList<QualityBrandTestCode>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        #endregion
    }
}
