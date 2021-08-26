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
    public class OvenProvider : SQLDAL, IOvenProvider
    {
        #region 底層連線
        public OvenProvider(string ConString) : base(ConString) { }
        public OvenProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Fabric Oven Test(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Fabric Oven Test
        /// </summary>
        /// <param name="Item">Fabric Oven Test成員</param>
        /// <returns>回傳Fabric Oven Test</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public IList<Oven> Get(Oven Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,TestNo"+ Environment.NewLine);
            SbSql.Append("        ,InspDate"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,addName"+ Environment.NewLine);
            SbSql.Append("        ,addDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,Time"+ Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append("FROM [Oven]"+ Environment.NewLine);



            return ExecuteList<Oven>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Fabric Oven Test(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Fabric Oven Test
        /// </summary>
        /// <param name="Item">Fabric Oven Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Create(Oven Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Oven]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,TestNo"+ Environment.NewLine);
            SbSql.Append("        ,InspDate"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,addName"+ Environment.NewLine);
            SbSql.Append("        ,addDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,Time"+ Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@TestNo"); objParameter.Add("@TestNo", DbType.String, Item.TestNo);
            SbSql.Append("        ,@InspDate"); objParameter.Add("@InspDate", DbType.String, Item.InspDate);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
            SbSql.Append("        ,@Inspector"); objParameter.Add("@Inspector", DbType.String, Item.Inspector);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@addName"); objParameter.Add("@addName", DbType.String, Item.addName);
            SbSql.Append("        ,@addDate"); objParameter.Add("@addDate", DbType.DateTime, Item.addDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@Temperature"); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);
            SbSql.Append("        ,@Time"); objParameter.Add("@Time", DbType.Int32, Item.Time);
            SbSql.Append("        ,@TestBeforePicture"); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);
            SbSql.Append("        ,@TestAfterPicture"); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Fabric Oven Test(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Fabric Oven Test
        /// </summary>
        /// <param name="Item">Fabric Oven Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Update(Oven Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Oven]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.POID != null) { SbSql.Append(",POID=@POID"+ Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID);}
            if (Item.TestNo != null) { SbSql.Append(",TestNo=@TestNo"+ Environment.NewLine); objParameter.Add("@TestNo", DbType.String, Item.TestNo);}
            if (Item.InspDate != null) { SbSql.Append(",InspDate=@InspDate"+ Environment.NewLine); objParameter.Add("@InspDate", DbType.String, Item.InspDate);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine); objParameter.Add("@Status", DbType.String, Item.Status);}
            if (Item.Inspector != null) { SbSql.Append(",Inspector=@Inspector"+ Environment.NewLine); objParameter.Add("@Inspector", DbType.String, Item.Inspector);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.addName != null) { SbSql.Append(",addName=@addName"+ Environment.NewLine); objParameter.Add("@addName", DbType.String, Item.addName);}
            if (Item.addDate != null) { SbSql.Append(",addDate=@addDate"+ Environment.NewLine); objParameter.Add("@addDate", DbType.DateTime, Item.addDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.Temperature != null) { SbSql.Append(",Temperature=@Temperature"+ Environment.NewLine); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);}
            if (Item.Time != null) { SbSql.Append(",Time=@Time"+ Environment.NewLine); objParameter.Add("@Time", DbType.Int32, Item.Time);}
            if (Item.TestBeforePicture != null) { SbSql.Append(",TestBeforePicture=@TestBeforePicture"+ Environment.NewLine); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);}
            if (Item.TestAfterPicture != null) { SbSql.Append(",TestAfterPicture=@TestAfterPicture"+ Environment.NewLine); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Fabric Oven Test(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Fabric Oven Test
        /// </summary>
        /// <param name="Item">Fabric Oven Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Delete(Oven Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Oven]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
