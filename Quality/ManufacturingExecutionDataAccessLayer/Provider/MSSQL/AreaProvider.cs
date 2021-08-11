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
    public class AreaProvider : SQLDAL, IAreaProvider
    {
        #region 底層連線
        public AreaProvider(string conString) : base(conString) { }
        public AreaProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<Area> Get(Area Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@Type", DbType.String, Item.Type } ,
                { "@T", DbType.String, Item.T } ,
                { "@B", DbType.String, Item.B } ,
                { "@I", DbType.String, Item.I } ,
                { "@O", DbType.String, Item.O } ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         Code"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,T"+ Environment.NewLine);
            SbSql.Append("        ,B"+ Environment.NewLine);
            SbSql.Append("        ,I"+ Environment.NewLine);
            SbSql.Append("        ,O"+ Environment.NewLine);
            SbSql.Append("        ,LocalCode"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("FROM [Area]"+ Environment.NewLine);
            SbSql.Append("WHERE Junk = 0" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.Type)) { SbSql.Append("and Type = @Type" + Environment.NewLine); }
            if (Item.T) { SbSql.Append("and T = @T" + Environment.NewLine); }
            if (Item.B) { SbSql.Append("and B = @B" + Environment.NewLine); }
            if (Item.I) { SbSql.Append("and I = @I" + Environment.NewLine); }
            if (Item.O) { SbSql.Append("and O = @O" + Environment.NewLine); }

            SbSql.Append("order by Seq" + Environment.NewLine);


            return ExecuteList<Area>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Area Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Area]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Code"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,T"+ Environment.NewLine);
            SbSql.Append("        ,B"+ Environment.NewLine);
            SbSql.Append("        ,I"+ Environment.NewLine);
            SbSql.Append("        ,O"+ Environment.NewLine);
            SbSql.Append("        ,LocalCode"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@T"); objParameter.Add("@T", DbType.String, Item.T);
            SbSql.Append("        ,@B"); objParameter.Add("@B", DbType.String, Item.B);
            SbSql.Append("        ,@I"); objParameter.Add("@I", DbType.String, Item.I);
            SbSql.Append("        ,@O"); objParameter.Add("@O", DbType.String, Item.O);
            SbSql.Append("        ,@LocalCode"); objParameter.Add("@LocalCode", DbType.String, Item.LocalCode);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.String, Item.Seq);
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
        public int Update(Area Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Area]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Code != null) { SbSql.Append("Code=@Code"+ Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.T != null) { SbSql.Append(",T=@T"+ Environment.NewLine); objParameter.Add("@T", DbType.String, Item.T);}
            if (Item.B != null) { SbSql.Append(",B=@B"+ Environment.NewLine); objParameter.Add("@B", DbType.String, Item.B);}
            if (Item.I != null) { SbSql.Append(",I=@I"+ Environment.NewLine); objParameter.Add("@I", DbType.String, Item.I);}
            if (Item.O != null) { SbSql.Append(",O=@O"+ Environment.NewLine); objParameter.Add("@O", DbType.String, Item.O);}
            if (Item.LocalCode != null) { SbSql.Append(",LocalCode=@LocalCode"+ Environment.NewLine); objParameter.Add("@LocalCode", DbType.String, Item.LocalCode);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.String, Item.Seq);}
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
        public int Delete(Area Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Area]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
