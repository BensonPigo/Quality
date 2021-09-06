using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using DatabaseObject.ResultModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class BulkFGTMailGroupProvider : SQLDAL, IBulkFGTMailGroupProvider
    {
        #region 底層連線
        public BulkFGTMailGroupProvider(string conString) : base(conString) { }
        public BulkFGTMailGroupProvider(SQLDataTransaction tra) : base(tra) { }

        public IList<Quality_MailGroup> MailGroupGet(Quality_MailGroup quality_Mail)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Type ", DbType.String, quality_Mail.Type);
            objParameter.Add("@FactoryID ", DbType.String, quality_Mail.FactoryID);
            objParameter.Add("@GroupName ", DbType.String, quality_Mail.GroupName);
            StringBuilder SbSql = new StringBuilder();
            SbSql.Append(@"
select q.FactoryID, q.Type, q.GroupName, q.ToAddress, q.CcAddress 
from Quality_MailGroup q 
where q.Type = @Type 
and q.FactoryID = @FactoryID 
");

            if (!string.IsNullOrEmpty(quality_Mail.GroupName))
            {
                SbSql.Append(" and q.GroupName = @GroupName ");
            }

            return ExecuteList<Quality_MailGroup>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        // type 0:新增/更新, 1:刪除
        public int MailGroupSave(Quality_MailGroup quality_Mail, int type)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID ", DbType.String, quality_Mail.FactoryID },
                { "@Type ", DbType.String, quality_Mail.Type },
                { "@GroupName ", DbType.String, quality_Mail.GroupName },
                { "@ToAddress ", DbType.String, quality_Mail.ToAddress },
                { "@CcAddress ", DbType.String, quality_Mail.CcAddress },
            };

            StringBuilder SbSql = new StringBuilder();
            if (type == 0)
            {
                SbSql.Append(@"
if not exists(
    select 1 from Quality_MailGroup 
    where FactoryID = @FactoryID
    and Type = @Type
    and GroupName = @GroupName )
Begin
INSERT INTO [dbo].[Quality_MailGroup]
           ([FactoryID]
           ,[Type]
           ,[GroupName]
           ,[ToAddress]
           ,[CcAddress])
     VALUES
           (@FactoryID
           ,@Type
           ,@GroupName
           ,@ToAddress
           ,@CcAddress)
End
else
Begin
    update Quality_MailGroup
    Set
        ToAddress = @ToAddress,
        CcAddress = @CcAddress
    where FactoryID = @FactoryID
    and Type = @Type
    and GroupName = @GroupName
End
");
            }
            else if (type == 1)
            {
                SbSql.Append(@"
delete Quality_MailGroup 
where FactoryID = @FactoryID
and Type = @Type
and GroupName = @GroupName  
");
            }

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

    }
}
