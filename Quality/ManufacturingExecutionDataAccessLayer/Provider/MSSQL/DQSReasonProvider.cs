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
    public class DQSReasonProvider : SQLDAL, IDQSReasonProvider
    {
        #region 底層連線
        public DQSReasonProvider(string conString) : base(conString) { }
        public DQSReasonProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<DQSReason> Get(DQSReason Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@Type", DbType.String, Item.Type},
                { "@ID", DbType.String, Item.ID},
                { "@Description", DbType.String, Item.Description},
                { "@LocalDescription", DbType.String, Item.LocalDescription},
            };
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         Type"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,LocalDescription"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("FROM [DQSReason]"+ Environment.NewLine);
            SbSql.Append("where 1=1" + Environment.NewLine);
            SbSql.Append("and Junk = 0" + Environment.NewLine);
            if (Item.Type != null) { SbSql.Append(" and Type=@Type" + Environment.NewLine); }
            if (Item.ID != null) { SbSql.Append(" and ID=@ID" + Environment.NewLine); }
            if (Item.Description != null) { SbSql.Append("and Description=@Description" + Environment.NewLine); }
            if (Item.LocalDescription != null) { SbSql.Append("and LocalDescription=@LocalDescription" + Environment.NewLine); }



            return ExecuteList<DQSReason>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(DQSReason Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [DQSReason]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Type"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,LocalDescription"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@LocalDescription"); objParameter.Add("@LocalDescription", DbType.String, Item.LocalDescription);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(DQSReason Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [DQSReason]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Type != null) { SbSql.Append("Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.ID != null) { SbSql.Append(",ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.LocalDescription != null) { SbSql.Append(",LocalDescription=@LocalDescription"+ Environment.NewLine); objParameter.Add("@LocalDescription", DbType.String, Item.LocalDescription);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(DQSReason Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [DQSReason]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
