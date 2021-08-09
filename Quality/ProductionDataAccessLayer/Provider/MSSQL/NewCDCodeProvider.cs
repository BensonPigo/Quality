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
    public class NewCDCodeProvider : SQLDAL, INewCDCodeProvider
    {
        #region 底層連線
        public NewCDCodeProvider(string ConString) : base(ConString) { }
        public NewCDCodeProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<NewCDCode> Get(NewCDCode Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         Classifty"+ Environment.NewLine);
            SbSql.Append("        ,TypeName"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Placket"+ Environment.NewLine);
            SbSql.Append("        ,Definition"+ Environment.NewLine);
            SbSql.Append("        ,CPU"+ Environment.NewLine);
            SbSql.Append("        ,ComboPcs"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("FROM [NewCDCode]"+ Environment.NewLine);



            return ExecuteList<NewCDCode>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(NewCDCode Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [NewCDCode]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Classifty"+ Environment.NewLine);
            SbSql.Append("        ,TypeName"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Placket"+ Environment.NewLine);
            SbSql.Append("        ,Definition"+ Environment.NewLine);
            SbSql.Append("        ,CPU"+ Environment.NewLine);
            SbSql.Append("        ,ComboPcs"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Classifty"); objParameter.Add("@Classifty", DbType.String, Item.Classifty);
            SbSql.Append("        ,@TypeName"); objParameter.Add("@TypeName", DbType.String, Item.TypeName);
            SbSql.Append("        ,@ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Placket"); objParameter.Add("@Placket", DbType.String, Item.Placket);
            SbSql.Append("        ,@Definition"); objParameter.Add("@Definition", DbType.String, Item.Definition);
            SbSql.Append("        ,@CPU"); objParameter.Add("@CPU", DbType.String, Item.CPU);
            SbSql.Append("        ,@ComboPcs"); objParameter.Add("@ComboPcs", DbType.String, Item.ComboPcs);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
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
        public int Update(NewCDCode Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [NewCDCode]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Classifty != null) { SbSql.Append("Classifty=@Classifty"+ Environment.NewLine); objParameter.Add("@Classifty", DbType.String, Item.Classifty);}
            if (Item.TypeName != null) { SbSql.Append(",TypeName=@TypeName"+ Environment.NewLine); objParameter.Add("@TypeName", DbType.String, Item.TypeName);}
            if (Item.ID != null) { SbSql.Append(",ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Placket != null) { SbSql.Append(",Placket=@Placket"+ Environment.NewLine); objParameter.Add("@Placket", DbType.String, Item.Placket);}
            if (Item.Definition != null) { SbSql.Append(",Definition=@Definition"+ Environment.NewLine); objParameter.Add("@Definition", DbType.String, Item.Definition);}
            if (Item.CPU != null) { SbSql.Append(",CPU=@CPU"+ Environment.NewLine); objParameter.Add("@CPU", DbType.String, Item.CPU);}
            if (Item.ComboPcs != null) { SbSql.Append(",ComboPcs=@ComboPcs"+ Environment.NewLine); objParameter.Add("@ComboPcs", DbType.String, Item.ComboPcs);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
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
        public int Delete(NewCDCode Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [NewCDCode]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
