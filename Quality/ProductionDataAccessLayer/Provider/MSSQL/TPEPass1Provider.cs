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
    public class TPEPass1Provider : SQLDAL, ITPEPass1Provider
    {
        #region 底層連線
        public TPEPass1Provider(string ConString) : base(ConString) { }
        public TPEPass1Provider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Taipei Pass1(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Taipei Pass1
        /// </summary>
        /// <param name="Item">Taipei Pass1成員</param>
        /// <returns>回傳Taipei Pass1</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<TPEPass1> Get(TPEPass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Name"+ Environment.NewLine);
            SbSql.Append("        ,ExtNo"+ Environment.NewLine);
            SbSql.Append("        ,EMail"+ Environment.NewLine);
            SbSql.Append("FROM [TPEPass1] WITH(NOLOCK)" + Environment.NewLine);



            return ExecuteList<TPEPass1>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Taipei Pass1(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Taipei Pass1
        /// </summary>
        /// <param name="Item">Taipei Pass1成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(TPEPass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [TPEPass1]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Name"+ Environment.NewLine);
            SbSql.Append("        ,ExtNo"+ Environment.NewLine);
            SbSql.Append("        ,EMail"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Name"); objParameter.Add("@Name", DbType.String, Item.Name);
            SbSql.Append("        ,@ExtNo"); objParameter.Add("@ExtNo", DbType.String, Item.ExtNo);
            SbSql.Append("        ,@EMail"); objParameter.Add("@EMail", DbType.String, Item.EMail);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Taipei Pass1(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Taipei Pass1
        /// </summary>
        /// <param name="Item">Taipei Pass1成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(TPEPass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [TPEPass1]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Name != null) { SbSql.Append(",Name=@Name"+ Environment.NewLine); objParameter.Add("@Name", DbType.String, Item.Name);}
            if (Item.ExtNo != null) { SbSql.Append(",ExtNo=@ExtNo"+ Environment.NewLine); objParameter.Add("@ExtNo", DbType.String, Item.ExtNo);}
            if (Item.EMail != null) { SbSql.Append(",EMail=@EMail"+ Environment.NewLine); objParameter.Add("@EMail", DbType.String, Item.EMail);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Taipei Pass1(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Taipei Pass1
        /// </summary>
        /// <param name="Item">Taipei Pass1成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(TPEPass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [TPEPass1]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
