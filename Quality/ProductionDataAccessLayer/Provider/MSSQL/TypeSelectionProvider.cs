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
    public class TypeSelectionProvider : SQLDAL,  ITypeSelectionProvider
    {
        #region 底層連線
        public TypeSelectionProvider(string ConString) : base(ConString) { }
        public TypeSelectionProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public IList<TypeSelection> Get(TypeSelection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         VersionID"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("FROM [TypeSelection]"+ Environment.NewLine);



            return ExecuteList<TypeSelection>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Create(TypeSelection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [TypeSelection]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         VersionID"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @VersionID"); objParameter.Add("@VersionID", DbType.Int32, Item.VersionID);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
            SbSql.Append("        ,@Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Update(TypeSelection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [TypeSelection]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.VersionID != null) { SbSql.Append("VersionID=@VersionID"+ Environment.NewLine); objParameter.Add("@VersionID", DbType.Int32, Item.VersionID);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
            if (Item.Code != null) { SbSql.Append(",Code=@Code"+ Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Delete(TypeSelection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [TypeSelection]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
