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
    /*各Brand的季別基本檔(SeasonProvider) 詳細敘述如下*/
    /// <summary>
    /// 各Brand的季別基本檔
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/19  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19  1.00    Admin        Create
    /// </history>
    public class SeasonProvider : SQLDAL, ISeasonProvider
    {
        #region 底層連線
        public SeasonProvider(string ConString) : base(ConString) { }
        public SeasonProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳各Brand的季別基本檔(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳各Brand的季別基本檔
        /// </summary>
        /// <param name="Item">各Brand的季別基本檔成員</param>
        /// <returns>回傳各Brand的季別基本檔</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public IList<Season> Get(Season Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,CostRatio"+ Environment.NewLine);
            SbSql.Append("        ,SeasonSCIID"+ Environment.NewLine);
            SbSql.Append("        ,Month"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("FROM [Season] WITH(NOLOCK)" + Environment.NewLine);



            return ExecuteList<Season>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立各Brand的季別基本檔(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立各Brand的季別基本檔
        /// </summary>
        /// <param name="Item">各Brand的季別基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Create(Season Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Season]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,CostRatio"+ Environment.NewLine);
            SbSql.Append("        ,SeasonSCIID"+ Environment.NewLine);
            SbSql.Append("        ,Month"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@CostRatio"); objParameter.Add("@CostRatio", DbType.String, Item.CostRatio);
            SbSql.Append("        ,@SeasonSCIID"); objParameter.Add("@SeasonSCIID", DbType.String, Item.SeasonSCIID);
            SbSql.Append("        ,@Month"); objParameter.Add("@Month", DbType.String, Item.Month);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新各Brand的季別基本檔(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新各Brand的季別基本檔
        /// </summary>
        /// <param name="Item">各Brand的季別基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Update(Season Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Season]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID"+ Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID);}
            if (Item.CostRatio != null) { SbSql.Append(",CostRatio=@CostRatio"+ Environment.NewLine); objParameter.Add("@CostRatio", DbType.String, Item.CostRatio);}
            if (Item.SeasonSCIID != null) { SbSql.Append(",SeasonSCIID=@SeasonSCIID"+ Environment.NewLine); objParameter.Add("@SeasonSCIID", DbType.String, Item.SeasonSCIID);}
            if (Item.Month != null) { SbSql.Append(",Month=@Month"+ Environment.NewLine); objParameter.Add("@Month", DbType.String, Item.Month);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除各Brand的季別基本檔(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除各Brand的季別基本檔
        /// </summary>
        /// <param name="Item">各Brand的季別基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Delete(Season Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Season]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
