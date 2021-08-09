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
    public class RFTPicDuringDummyFittingProvider : SQLDAL, IRFTPicDuringDummyFittingProvider
    {
        #region 底層連線
        public RFTPicDuringDummyFittingProvider(string conString) : base(conString) { }
        public RFTPicDuringDummyFittingProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<RFT_PicDuringDummyFitting> Get(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Front"+ Environment.NewLine);
            SbSql.Append("        ,Side"+ Environment.NewLine);
            SbSql.Append("        ,Back"+ Environment.NewLine);
            SbSql.Append("FROM [RFT_PicDuringDummyFitting]"+ Environment.NewLine);



            return ExecuteList<RFT_PicDuringDummyFitting>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_PicDuringDummyFitting]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Front"+ Environment.NewLine);
            SbSql.Append("        ,Side"+ Environment.NewLine);
            SbSql.Append("        ,Back"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Size"); objParameter.Add("@Size", DbType.String, Item.Size);
            SbSql.Append("        ,@Front"); objParameter.Add("@Front", DbType.String, Item.Front);
            SbSql.Append("        ,@Side"); objParameter.Add("@Side", DbType.String, Item.Side);
            SbSql.Append("        ,@Back"); objParameter.Add("@Back", DbType.String, Item.Back);
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
        public int Update(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_PicDuringDummyFitting]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.OrderID != null) { SbSql.Append("OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Size != null) { SbSql.Append(",Size=@Size"+ Environment.NewLine); objParameter.Add("@Size", DbType.String, Item.Size);}
            if (Item.Front != null) { SbSql.Append(",Front=@Front"+ Environment.NewLine); objParameter.Add("@Front", DbType.String, Item.Front);}
            if (Item.Side != null) { SbSql.Append(",Side=@Side"+ Environment.NewLine); objParameter.Add("@Side", DbType.String, Item.Side);}
            if (Item.Back != null) { SbSql.Append(",Back=@Back"+ Environment.NewLine); objParameter.Add("@Back", DbType.String, Item.Back);}
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
        public int Delete(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_PicDuringDummyFitting]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
