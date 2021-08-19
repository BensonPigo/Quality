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
    /*(QualityMenuProvider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/07/30  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/07/30  1.00    Admin        Create
    /// </history>
    public class QualityMenuProvider : SQLDAL, IQualityMenuProvider
    {
        #region 底層連線
        public QualityMenuProvider(string conString) : base(conString) { }
        public QualityMenuProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<Quality_Menu> Get(Quality_Pass1 pass1)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@PositionID", DbType.String, pass1.Position} ,
                { "@Brand", DbType.String, pass1.BulkFGT_Brand} ,
            };

            SbSql.Append(@"
select m.[ID]
	, m.[ModuleName]
	, m.[ModuleSeq]
	, [FunctionName] = isnull(md.[FunctionName] ,m.[FunctionName])
	, m.[FunctionSeq]
	, m.[Junk]
	, m.[Url]
from Quality_Position p
inner join Quality_Pass2 p2 on p.ID = PositionID
inner join Quality_Menu m on m.ID = p2.MenuID
left join Quality_Menu_Detail md on md.ID = m.ID and md.Type = @Brand
where p.ID = @PositionID
and p2.Used = iif(p.IsAdmin = 1, p2.Used, 1)
and m.Junk = 0
order by ModuleSeq, FunctionSeq" + Environment.NewLine);

            return ExecuteList<Quality_Menu>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Quality_Menu Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Quality_Menu]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,ModuleName"+ Environment.NewLine);
            SbSql.Append("        ,ModuleSeq"+ Environment.NewLine);
            SbSql.Append("        ,FunctionName"+ Environment.NewLine);
            SbSql.Append("        ,FunctionSeq"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@ModuleName"); objParameter.Add("@ModuleName", DbType.String, Item.ModuleName);
            SbSql.Append("        ,@ModuleSeq"); objParameter.Add("@ModuleSeq", DbType.Int32, Item.ModuleSeq);
            SbSql.Append("        ,@FunctionName"); objParameter.Add("@FunctionName", DbType.String, Item.FunctionName);
            SbSql.Append("        ,@FunctionSeq"); objParameter.Add("@FunctionSeq", DbType.Int32, Item.FunctionSeq);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
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
        public int Update(Quality_Menu Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Quality_Menu]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID > 0) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.ModuleName != null) { SbSql.Append(",ModuleName=@ModuleName"+ Environment.NewLine); objParameter.Add("@ModuleName", DbType.String, Item.ModuleName);}
            if (Item.ModuleSeq != null) { SbSql.Append(",ModuleSeq=@ModuleSeq"+ Environment.NewLine); objParameter.Add("@ModuleSeq", DbType.Int32, Item.ModuleSeq);}
            if (Item.FunctionName != null) { SbSql.Append(",FunctionName=@FunctionName"+ Environment.NewLine); objParameter.Add("@FunctionName", DbType.String, Item.FunctionName);}
            if (Item.FunctionSeq != null) { SbSql.Append(",FunctionSeq=@FunctionSeq"+ Environment.NewLine); objParameter.Add("@FunctionSeq", DbType.Int32, Item.FunctionSeq);}
            if (Item.Junk) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
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
        public int Delete(Quality_Menu Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Quality_Menu]"+ Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
