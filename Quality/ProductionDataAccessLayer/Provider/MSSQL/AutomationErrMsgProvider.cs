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
    public class AutomationErrMsgProvider : SQLDAL, IAutomationErrMsgProvider
    {
        #region 底層連線
        public AutomationErrMsgProvider(string ConString) : base(ConString) { }
        public AutomationErrMsgProvider(SQLDataTransaction tra) : base(tra) { }

        public void Insert(AutomationErrMsg automationErrMsg)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@SuppID", automationErrMsg.suppID);
            objParameter.Add("@ModuleName", automationErrMsg.moduleName);
            objParameter.Add("@APIThread", automationErrMsg.apiThread);
            objParameter.Add("@SuppAPIThread", automationErrMsg.suppAPIThread);
            objParameter.Add("@ErrorMsg", automationErrMsg.errorMsg);
            objParameter.Add("@JSON", automationErrMsg.json);
            objParameter.Add("@AddName", automationErrMsg.addName);

            string sqlInsert = @"
insert into AutomationErrMsg(SuppID, ModuleName, APIThread, SuppAPIThread, ErrorMsg, JSON, AddName, AddDate)
            values(@SuppID, @ModuleName, @APIThread, @SuppAPIThread, @ErrorMsg, @JSON, @AddName, getdate())
";

            ExecuteNonQuery(CommandType.Text, sqlInsert, objParameter);
        }
        #endregion

    }
}
