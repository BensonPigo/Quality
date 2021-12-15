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
    public class OrderQtyProvider : SQLDAL, IOrderQtyProvider
    {
        #region 底層連線
        public OrderQtyProvider(string ConString) : base(ConString) { }
        public OrderQtyProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Qty Breakdown(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Qty Breakdown
        /// </summary>
        /// <param name="Item">Qty Breakdown成員</param>
        /// <returns>回傳Qty Breakdown</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<Order_Qty> Get(Order_Qty Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,Qty"+ Environment.NewLine);
            SbSql.Append("        ,OriQty"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("FROM [Order_Qty] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ID.ToString())) { SbSql.Append("And ID = @ID" + Environment.NewLine); }
            objParameter.Add("@ID", DbType.String, Item.ID);



            return ExecuteList<Order_Qty>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Qty Breakdown(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Qty Breakdown
        /// </summary>
        /// <param name="Item">Qty Breakdown成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(Order_Qty Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Order_Qty]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,Qty"+ Environment.NewLine);
            SbSql.Append("        ,OriQty"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@Qty"); objParameter.Add("@Qty", DbType.Int32, Item.Qty);
            SbSql.Append("        ,@OriQty"); objParameter.Add("@OriQty", DbType.Int32, Item.OriQty);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Qty Breakdown(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Qty Breakdown
        /// </summary>
        /// <param name="Item">Qty Breakdown成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(Order_Qty Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Order_Qty]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode"+ Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);}
            if (Item.Qty != null) { SbSql.Append(",Qty=@Qty"+ Environment.NewLine); objParameter.Add("@Qty", DbType.Int32, Item.Qty);}
            if (Item.OriQty != null) { SbSql.Append(",OriQty=@OriQty"+ Environment.NewLine); objParameter.Add("@OriQty", DbType.Int32, Item.OriQty);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Qty Breakdown(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Qty Breakdown
        /// </summary>
        /// <param name="Item">Qty Breakdown成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(Order_Qty Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Order_Qty]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        public IList<Order_Qty> GetDistinctArticle(Order_Qty Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT Distinct Article
From Order_Qty WITH(NOLOCK)
Where 1 = 1
");
            if (!string.IsNullOrEmpty(Item.ID.ToString())) { SbSql.Append("And ID = @ID" + Environment.NewLine); }
            objParameter.Add("@ID", DbType.String, Item.ID);

            return ExecuteList<Order_Qty>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
