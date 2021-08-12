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
    public class ReworkCardProvider : SQLDAL, IReworkCardProvider
    {
        #region 底層連線
        public ReworkCardProvider(string conString) : base(conString) { }
        public ReworkCardProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<ReworkCard> Get(ReworkCard Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@FactoryID", DbType.String, Item.FactoryID},
                { "@Line", DbType.String, Item.Line},
                { "@Type", DbType.String, Item.Type} ,
            };

            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         No" + Environment.NewLine);
            SbSql.Append("        ,Type" + Environment.NewLine);
            SbSql.Append("        ,FactoryID" + Environment.NewLine);
            SbSql.Append("        ,Line" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,Status" + Environment.NewLine);
            SbSql.Append("FROM [ReworkCard]" + Environment.NewLine);
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.FactoryID)) { SbSql.Append(" and FactoryID = @FactoryID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Line)) { SbSql.Append(" and Line = @Line" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Type)) { SbSql.Append(" and Type = @Type" + Environment.NewLine); }

            return ExecuteList<ReworkCard>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(ReworkCard Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [ReworkCard]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         No"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,FactoryID"+ Environment.NewLine);
            SbSql.Append("        ,Line"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @No"); objParameter.Add("@No", DbType.String, Item.No);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@FactoryID"); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);
            SbSql.Append("        ,@Line"); objParameter.Add("@Line", DbType.String, Item.Line);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
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
        public int Update(ReworkCard Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                  { "@No", DbType.String, Item.No},
                  { "@Type", DbType.String, Item.Type},
                  { "@FactoryID", DbType.String, Item.FactoryID},
                  { "@Line", DbType.String, Item.Line},
                  { "@Status", DbType.String, Item.Status},

                  { "@AddDate", DbType.DateTime, Item.AddDate},
                  { "@AddName", DbType.String, Item.AddName},
                  { "@EditDate", DbType.DateTime, Item.EditDate},
                  { "@EditName", DbType.String, Item.EditName},
            };

            SbSql.Append("UPDATE [ReworkCard]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.No))
            {
                SbSql.Append("No=@No"+ Environment.NewLine);
            }
            else
            {
                SbSql.Append("No=No" + Environment.NewLine);
            }

            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine);}
            if (Item.FactoryID != null) { SbSql.Append(",FactoryID=@FactoryID"+ Environment.NewLine);}
            if (Item.Line != null) { SbSql.Append(",Line=@Line"+ Environment.NewLine);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
            if (Item.No != null) { SbSql.Append("No=@No" + Environment.NewLine); }
            if (Item.Type != null) { SbSql.Append(",Type=@Type" + Environment.NewLine); }
            if (Item.FactoryID != null) { SbSql.Append(",FactoryID=@FactoryID" + Environment.NewLine); }
            if (Item.Line != null) { SbSql.Append(",Line=@Line" + Environment.NewLine); }

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
        public int Delete(ReworkCard Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [ReworkCard]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
