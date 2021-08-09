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
    public class MeasurementProvider : SQLDAL, IMeasurementProvider
    {
        #region 底層連線
        public MeasurementProvider(string conString) : base(conString) { }
        public MeasurementProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<Measurement> Get(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Tol1"+ Environment.NewLine);
            SbSql.Append("        ,Tol2"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,SizeGroup"+ Environment.NewLine);
            SbSql.Append("        ,MeasurementTranslateUkey"+ Environment.NewLine);
            SbSql.Append("FROM [Measurement]"+ Environment.NewLine);



            return ExecuteList<Measurement>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Measurement]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Tol1"+ Environment.NewLine);
            SbSql.Append("        ,Tol2"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,SizeGroup"+ Environment.NewLine);
            SbSql.Append("        ,MeasurementTranslateUkey"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@Tol1"); objParameter.Add("@Tol1", DbType.String, Item.Tol1);
            SbSql.Append("        ,@Tol2"); objParameter.Add("@Tol2", DbType.String, Item.Tol2);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@SizeGroup"); objParameter.Add("@SizeGroup", DbType.String, Item.SizeGroup);
            SbSql.Append("        ,@MeasurementTranslateUkey"); objParameter.Add("@MeasurementTranslateUkey", DbType.String, Item.MeasurementTranslateUkey);
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
        public int Update(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Measurement]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.StyleUkey != null) { SbSql.Append("StyleUkey=@StyleUkey"+ Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);}
            if (Item.Tol1 != null) { SbSql.Append(",Tol1=@Tol1"+ Environment.NewLine); objParameter.Add("@Tol1", DbType.String, Item.Tol1);}
            if (Item.Tol2 != null) { SbSql.Append(",Tol2=@Tol2"+ Environment.NewLine); objParameter.Add("@Tol2", DbType.String, Item.Tol2);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.Code != null) { SbSql.Append(",Code=@Code"+ Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code);}
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode"+ Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);}
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec"+ Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);}
            if (Item.Ukey != null) { SbSql.Append(",Ukey=@Ukey"+ Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.SizeGroup != null) { SbSql.Append(",SizeGroup=@SizeGroup"+ Environment.NewLine); objParameter.Add("@SizeGroup", DbType.String, Item.SizeGroup);}
            if (Item.MeasurementTranslateUkey != null) { SbSql.Append(",MeasurementTranslateUkey=@MeasurementTranslateUkey"+ Environment.NewLine); objParameter.Add("@MeasurementTranslateUkey", DbType.String, Item.MeasurementTranslateUkey);}
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
        public int Delete(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Measurement]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
