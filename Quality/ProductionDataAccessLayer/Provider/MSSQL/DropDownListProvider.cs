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
    public class DropDownListProvider : SQLDAL, IDropDownListProvider
    {
        #region 底層連線
        public DropDownListProvider(string ConString) : base(ConString) { }
        public DropDownListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳將getReason.prg 轉Table , ID 和Name長度需一樣(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳將getReason.prg 轉Table , ID 和Name長度需一樣
        /// </summary>
        /// <param name="Item">將getReason.prg 轉Table , ID 和Name長度需一樣成員</param>
        /// <returns>回傳將getReason.prg 轉Table , ID 和Name長度需一樣</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<DropDownList> Get(DropDownList Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@Type", DbType.String, Item.Type } ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         Type"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Name"+ Environment.NewLine);
            SbSql.Append("        ,RealLength"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("FROM [DropDownList] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("WHERE Type = @Type" + Environment.NewLine);
            SbSql.Append("Order by Seq" + Environment.NewLine);


            return ExecuteList<DropDownList>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立將getReason.prg 轉Table , ID 和Name長度需一樣(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立將getReason.prg 轉Table , ID 和Name長度需一樣
        /// </summary>
        /// <param name="Item">將getReason.prg 轉Table , ID 和Name長度需一樣成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(DropDownList Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [DropDownList]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Type"+ Environment.NewLine);
            SbSql.Append("        ,ID"+ Environment.NewLine);
            SbSql.Append("        ,Name"+ Environment.NewLine);
            SbSql.Append("        ,RealLength"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Name"); objParameter.Add("@Name", DbType.String, Item.Name);
            SbSql.Append("        ,@RealLength"); objParameter.Add("@RealLength", DbType.String, Item.RealLength);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新將getReason.prg 轉Table , ID 和Name長度需一樣(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新將getReason.prg 轉Table , ID 和Name長度需一樣
        /// </summary>
        /// <param name="Item">將getReason.prg 轉Table , ID 和Name長度需一樣成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(DropDownList Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [DropDownList]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Type != null) { SbSql.Append("Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.ID != null) { SbSql.Append(",ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Name != null) { SbSql.Append(",Name=@Name"+ Environment.NewLine); objParameter.Add("@Name", DbType.String, Item.Name);}
            if (Item.RealLength != null) { SbSql.Append(",RealLength=@RealLength"+ Environment.NewLine); objParameter.Add("@RealLength", DbType.String, Item.RealLength);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除將getReason.prg 轉Table , ID 和Name長度需一樣(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除將getReason.prg 轉Table , ID 和Name長度需一樣
        /// </summary>
        /// <param name="Item">將getReason.prg 轉Table , ID 和Name長度需一樣成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(DropDownList Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [DropDownList]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
