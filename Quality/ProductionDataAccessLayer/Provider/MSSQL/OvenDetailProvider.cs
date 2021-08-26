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
    public class OvenDetailProvider : SQLDAL, IOvenDetailProvider
    {
        #region 底層連線
        public OvenDetailProvider(string ConString) : base(ConString) { }
        public OvenDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Fabric Oven Test Detail(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Fabric Oven Test Detail
        /// </summary>
        /// <param name="Item">Fabric Oven Test Detail成員</param>
        /// <returns>回傳Fabric Oven Test Detail</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public IList<Oven_Detail> Get(Oven_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,OvenGroup"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,changeScale"+ Environment.NewLine);
            SbSql.Append("        ,StainingScale"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,ResultChange"+ Environment.NewLine);
            SbSql.Append("        ,ResultStain"+ Environment.NewLine);
            SbSql.Append("        ,SubmitDate"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,Time"+ Environment.NewLine);
            SbSql.Append("FROM [Oven_Detail]"+ Environment.NewLine);



            return ExecuteList<Oven_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Fabric Oven Test Detail(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Fabric Oven Test Detail
        /// </summary>
        /// <param name="Item">Fabric Oven Test Detail成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Create(Oven_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Oven_Detail]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,OvenGroup"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,changeScale"+ Environment.NewLine);
            SbSql.Append("        ,StainingScale"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,ResultChange"+ Environment.NewLine);
            SbSql.Append("        ,ResultStain"+ Environment.NewLine);
            SbSql.Append("        ,SubmitDate"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,Time"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@OvenGroup"); objParameter.Add("@OvenGroup", DbType.String, Item.OvenGroup);
            SbSql.Append("        ,@SEQ1"); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);
            SbSql.Append("        ,@SEQ2"); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);
            SbSql.Append("        ,@Roll"); objParameter.Add("@Roll", DbType.String, Item.Roll);
            SbSql.Append("        ,@Dyelot"); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@changeScale"); objParameter.Add("@changeScale", DbType.String, Item.changeScale);
            SbSql.Append("        ,@StainingScale"); objParameter.Add("@StainingScale", DbType.String, Item.StainingScale);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@ResultChange"); objParameter.Add("@ResultChange", DbType.String, Item.ResultChange);
            SbSql.Append("        ,@ResultStain"); objParameter.Add("@ResultStain", DbType.String, Item.ResultStain);
            SbSql.Append("        ,@SubmitDate"); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);
            SbSql.Append("        ,@Temperature"); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);
            SbSql.Append("        ,@Time"); objParameter.Add("@Time", DbType.Int32, Item.Time);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Fabric Oven Test Detail(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Fabric Oven Test Detail
        /// </summary>
        /// <param name="Item">Fabric Oven Test Detail成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Update(Oven_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Oven_Detail]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.OvenGroup != null) { SbSql.Append(",OvenGroup=@OvenGroup"+ Environment.NewLine); objParameter.Add("@OvenGroup", DbType.String, Item.OvenGroup);}
            if (Item.SEQ1 != null) { SbSql.Append(",SEQ1=@SEQ1"+ Environment.NewLine); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);}
            if (Item.SEQ2 != null) { SbSql.Append(",SEQ2=@SEQ2"+ Environment.NewLine); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);}
            if (Item.Roll != null) { SbSql.Append(",Roll=@Roll"+ Environment.NewLine); objParameter.Add("@Roll", DbType.String, Item.Roll);}
            if (Item.Dyelot != null) { SbSql.Append(",Dyelot=@Dyelot"+ Environment.NewLine); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.changeScale != null) { SbSql.Append(",changeScale=@changeScale"+ Environment.NewLine); objParameter.Add("@changeScale", DbType.String, Item.changeScale);}
            if (Item.StainingScale != null) { SbSql.Append(",StainingScale=@StainingScale"+ Environment.NewLine); objParameter.Add("@StainingScale", DbType.String, Item.StainingScale);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.ResultChange != null) { SbSql.Append(",ResultChange=@ResultChange"+ Environment.NewLine); objParameter.Add("@ResultChange", DbType.String, Item.ResultChange);}
            if (Item.ResultStain != null) { SbSql.Append(",ResultStain=@ResultStain"+ Environment.NewLine); objParameter.Add("@ResultStain", DbType.String, Item.ResultStain);}
            if (Item.SubmitDate != null) { SbSql.Append(",SubmitDate=@SubmitDate"+ Environment.NewLine); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);}
            if (Item.Temperature != null) { SbSql.Append(",Temperature=@Temperature"+ Environment.NewLine); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);}
            if (Item.Time != null) { SbSql.Append(",Time=@Time"+ Environment.NewLine); objParameter.Add("@Time", DbType.Int32, Item.Time);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Fabric Oven Test Detail(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Fabric Oven Test Detail
        /// </summary>
        /// <param name="Item">Fabric Oven Test Detail成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/26  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/26  1.00    Admin        Create
        /// </history>
        public int Delete(Oven_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Oven_Detail]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
