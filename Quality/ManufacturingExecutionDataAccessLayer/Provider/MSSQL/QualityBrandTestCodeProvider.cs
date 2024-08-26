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
WHERE  1=1 
" + Environment.NewLine);

            if (!string.IsNullOrEmpty(brandId))
            {
                SbSql.Append("AND BrandID = @BrandID" + Environment.NewLine);
                objParameter.Add("@BrandID", DbType.String, brandId);
            }

            if (!string.IsNullOrEmpty(functuionName))
            {
                SbSql.Append("AND FunctionName = @FunctionName" + Environment.NewLine);
                objParameter.Add("@FunctionName", DbType.String, functuionName);
            }

            return ExecuteList<QualityBrandTestCode>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        #endregion
    }
}
