using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;

namespace ManufacturingExecutionDataAccessLayer.DataAccessLayer.Provider.MSSQL
{
    /*(QualityPass2Provider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/07/30  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/07/30  1.00    Admin        Create
    /// </history>
    public class QualityPass2Provider : SQLDAL, IQualityPass2Provider
    {
        #region 底層連線
        public QualityPass2Provider(string conString) : base(conString) { }
        public QualityPass2Provider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<Quality_Pass2> Get(Quality_Pass2 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         PositionID"+ Environment.NewLine);
            SbSql.Append("        ,MenuID"+ Environment.NewLine);
            SbSql.Append("        ,Used"+ Environment.NewLine);
            SbSql.Append("FROM [Quality_Pass2]"+ Environment.NewLine);



            return ExecuteList<Quality_Pass2>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Quality_Pass2 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Quality_Pass2]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         PositionID"+ Environment.NewLine);
            SbSql.Append("        ,MenuID"+ Environment.NewLine);
            SbSql.Append("        ,Used"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @PositionID"); objParameter.Add("@PositionID", DbType.String, Item.PositionID);
            SbSql.Append("        ,@MenuID"); objParameter.Add("@MenuID", DbType.String, Item.MenuID);
            SbSql.Append("        ,@Used"); objParameter.Add("@Used", DbType.String, Item.Used);
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
        public int Update(Quality_Pass2 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Quality_Pass2]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.PositionID != null) { SbSql.Append("PositionID=@PositionID"+ Environment.NewLine); objParameter.Add("@PositionID", DbType.String, Item.PositionID);}
            if (Item.MenuID > 0) { SbSql.Append(",MenuID=@MenuID"+ Environment.NewLine); objParameter.Add("@MenuID", DbType.String, Item.MenuID);}
            if (Item.Used) { SbSql.Append(",Used=@Used"+ Environment.NewLine); objParameter.Add("@Used", DbType.String, Item.Used);}
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
        public int Delete(Quality_Pass2 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Quality_Pass2]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
