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
    public class GarmentTestDetailApperanceProvider : SQLDAL, IGarmentTestDetailApperanceProvider
    {
        #region 底層連線
        public GarmentTestDetailApperanceProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailApperanceProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<GarmentTest_Detail_Apperance> Get(GarmentTest_Detail_Apperance Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,Wash1"+ Environment.NewLine);
            SbSql.Append("        ,Wash2"+ Environment.NewLine);
            SbSql.Append("        ,Wash3"+ Environment.NewLine);
            SbSql.Append("        ,Comment"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Wash4"+ Environment.NewLine);
            SbSql.Append("        ,Wash5"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest_Detail_Apperance]"+ Environment.NewLine);



            return ExecuteList<GarmentTest_Detail_Apperance>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(GarmentTest_Detail_Apperance Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest_Detail_Apperance]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,Wash1"+ Environment.NewLine);
            SbSql.Append("        ,Wash2"+ Environment.NewLine);
            SbSql.Append("        ,Wash3"+ Environment.NewLine);
            SbSql.Append("        ,Comment"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Wash4"+ Environment.NewLine);
            SbSql.Append("        ,Wash5"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@Wash1"); objParameter.Add("@Wash1", DbType.String, Item.Wash1);
            SbSql.Append("        ,@Wash2"); objParameter.Add("@Wash2", DbType.String, Item.Wash2);
            SbSql.Append("        ,@Wash3"); objParameter.Add("@Wash3", DbType.String, Item.Wash3);
            SbSql.Append("        ,@Comment"); objParameter.Add("@Comment", DbType.String, Item.Comment);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
            SbSql.Append("        ,@Wash4"); objParameter.Add("@Wash4", DbType.String, Item.Wash4);
            SbSql.Append("        ,@Wash5"); objParameter.Add("@Wash5", DbType.String, Item.Wash5);
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
        public int Update(GarmentTest_Detail_Apperance Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest_Detail_Apperance]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.Wash1 != null) { SbSql.Append(",Wash1=@Wash1"+ Environment.NewLine); objParameter.Add("@Wash1", DbType.String, Item.Wash1);}
            if (Item.Wash2 != null) { SbSql.Append(",Wash2=@Wash2"+ Environment.NewLine); objParameter.Add("@Wash2", DbType.String, Item.Wash2);}
            if (Item.Wash3 != null) { SbSql.Append(",Wash3=@Wash3"+ Environment.NewLine); objParameter.Add("@Wash3", DbType.String, Item.Wash3);}
            if (Item.Comment != null) { SbSql.Append(",Comment=@Comment"+ Environment.NewLine); objParameter.Add("@Comment", DbType.String, Item.Comment);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
            if (Item.Wash4 != null) { SbSql.Append(",Wash4=@Wash4"+ Environment.NewLine); objParameter.Add("@Wash4", DbType.String, Item.Wash4);}
            if (Item.Wash5 != null) { SbSql.Append(",Wash5=@Wash5"+ Environment.NewLine); objParameter.Add("@Wash5", DbType.String, Item.Wash5);}
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
        public int Delete(GarmentTest_Detail_Apperance Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest_Detail_Apperance]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
