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
    public class GarmentTestDetailFGWTProvider : SQLDAL, IGarmentTestDetailFGWTProvider
    {
        #region 底層連線
        public GarmentTestDetailFGWTProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailFGWTProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<GarmentTest_Detail_FGWT> Get(GarmentTest_Detail_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,BeforeWash"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,AfterWash"+ Environment.NewLine);
            SbSql.Append("        ,Shrinkage"+ Environment.NewLine);
            SbSql.Append("        ,Scale"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,Criteria2"+ Environment.NewLine);
            SbSql.Append("        ,SystemType"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest_Detail_FGWT]"+ Environment.NewLine);



            return ExecuteList<GarmentTest_Detail_FGWT>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(GarmentTest_Detail_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest_Detail_FGWT]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,BeforeWash"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,AfterWash"+ Environment.NewLine);
            SbSql.Append("        ,Shrinkage"+ Environment.NewLine);
            SbSql.Append("        ,Scale"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,Criteria2"+ Environment.NewLine);
            SbSql.Append("        ,SystemType"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@TestDetail"); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);
            SbSql.Append("        ,@BeforeWash"); objParameter.Add("@BeforeWash", DbType.String, Item.BeforeWash);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@AfterWash"); objParameter.Add("@AfterWash", DbType.String, Item.AfterWash);
            SbSql.Append("        ,@Shrinkage"); objParameter.Add("@Shrinkage", DbType.String, Item.Shrinkage);
            SbSql.Append("        ,@Scale"); objParameter.Add("@Scale", DbType.String, Item.Scale);
            SbSql.Append("        ,@Criteria"); objParameter.Add("@Criteria", DbType.String, Item.Criteria);
            SbSql.Append("        ,@Criteria2"); objParameter.Add("@Criteria2", DbType.String, Item.Criteria2);
            SbSql.Append("        ,@SystemType"); objParameter.Add("@SystemType", DbType.String, Item.SystemType);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
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
        public int Update(GarmentTest_Detail_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest_Detail_FGWT]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.TestDetail != null) { SbSql.Append(",TestDetail=@TestDetail"+ Environment.NewLine); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);}
            if (Item.BeforeWash != null) { SbSql.Append(",BeforeWash=@BeforeWash"+ Environment.NewLine); objParameter.Add("@BeforeWash", DbType.String, Item.BeforeWash);}
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec"+ Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);}
            if (Item.AfterWash != null) { SbSql.Append(",AfterWash=@AfterWash"+ Environment.NewLine); objParameter.Add("@AfterWash", DbType.String, Item.AfterWash);}
            if (Item.Shrinkage != null) { SbSql.Append(",Shrinkage=@Shrinkage"+ Environment.NewLine); objParameter.Add("@Shrinkage", DbType.String, Item.Shrinkage);}
            if (Item.Scale != null) { SbSql.Append(",Scale=@Scale"+ Environment.NewLine); objParameter.Add("@Scale", DbType.String, Item.Scale);}
            if (Item.Criteria != null) { SbSql.Append(",Criteria=@Criteria"+ Environment.NewLine); objParameter.Add("@Criteria", DbType.String, Item.Criteria);}
            if (Item.Criteria2 != null) { SbSql.Append(",Criteria2=@Criteria2"+ Environment.NewLine); objParameter.Add("@Criteria2", DbType.String, Item.Criteria2);}
            if (Item.SystemType != null) { SbSql.Append(",SystemType=@SystemType"+ Environment.NewLine); objParameter.Add("@SystemType", DbType.String, Item.SystemType);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
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
        public int Delete(GarmentTest_Detail_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest_Detail_FGWT]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
