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
    /*(QualityPass1Provider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/07/30  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/07/30  1.00    Admin        Create
    /// </history>
    public class QualityPass1Provider : SQLDAL, IQualityPass1Provider
    {
        #region 底層連線
        public QualityPass1Provider(string conString) : base(conString) { }
        public QualityPass1Provider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public IList<Quality_Pass1> Get(Quality_Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, Item.ID } ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Position"+ Environment.NewLine);
            SbSql.Append("        ,BulkFGT_Brand" + Environment.NewLine);
            SbSql.Append("FROM [Quality_Pass1]"+ Environment.NewLine);
            SbSql.Append("Where ID=@ID" + Environment.NewLine);

            return ExecuteList<Quality_Pass1>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public int Create(Quality_Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Quality_Pass1]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Position"+ Environment.NewLine);
            SbSql.Append("        ,BulkFGT_Brand" + Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Position"); objParameter.Add("@Position", DbType.String, Item.Position);
            SbSql.Append("        ,@BulkFGT_Brand"); objParameter.Add("@BulkFGT_Brand", DbType.String, Item.BulkFGT_Brand);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public int Update(Quality_Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Quality_Pass1]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Position != null) { SbSql.Append(",Position=@Position"+ Environment.NewLine); objParameter.Add("@Position", DbType.String, Item.Position);}
            if (Item.BulkFGT_Brand != null) { SbSql.Append(",BulkFGT_Brand=@BulkFGT_Brand" + Environment.NewLine); objParameter.Add("@BulkFGT_Brand", DbType.String, Item.BulkFGT_Brand);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public int Delete(Quality_Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Quality_Pass1]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
