using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using DatabaseObject.ProductionDB;
using ADOHelper.Utility;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class FIRLaboratoryHeatProvider : SQLDAL, IFIRLaboratoryHeatProvider
    {
        #region 底層連線
        public FIRLaboratoryHeatProvider(string ConString) : base(ConString) { }
        public FIRLaboratoryHeatProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Laboratory - Heat Test(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Laboratory - Heat Test
        /// </summary>
        /// <param name="Item">Laboratory - Heat Test成員</param>
        /// <returns>回傳Laboratory - Heat Test</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public IList<FIR_Laboratory_Heat> Get(FIR_Laboratory_Heat Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,Inspdate"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalRate"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalOriginal"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest1"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest2"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest3"+ Environment.NewLine);
            SbSql.Append("        ,VerticalRate"+ Environment.NewLine);
            SbSql.Append("        ,VerticalOriginal"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest1"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest2"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest3"+ Environment.NewLine);
            SbSql.Append("FROM [FIR_Laboratory_Heat]"+ Environment.NewLine);



            return ExecuteList<FIR_Laboratory_Heat>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Laboratory - Heat Test(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Laboratory - Heat Test
        /// </summary>
        /// <param name="Item">Laboratory - Heat Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Create(FIR_Laboratory_Heat Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [FIR_Laboratory_Heat]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,Inspdate"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalRate"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalOriginal"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest1"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest2"+ Environment.NewLine);
            SbSql.Append("        ,HorizontalTest3"+ Environment.NewLine);
            SbSql.Append("        ,VerticalRate"+ Environment.NewLine);
            SbSql.Append("        ,VerticalOriginal"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest1"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest2"+ Environment.NewLine);
            SbSql.Append("        ,VerticalTest3"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Roll"); objParameter.Add("@Roll", DbType.String, Item.Roll);
            SbSql.Append("        ,@Dyelot"); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);
            SbSql.Append("        ,@Inspdate"); objParameter.Add("@Inspdate", DbType.String, Item.Inspdate);
            SbSql.Append("        ,@Inspector"); objParameter.Add("@Inspector", DbType.String, Item.Inspector);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@HorizontalRate"); objParameter.Add("@HorizontalRate", DbType.String, Item.HorizontalRate);
            SbSql.Append("        ,@HorizontalOriginal"); objParameter.Add("@HorizontalOriginal", DbType.String, Item.HorizontalOriginal);
            SbSql.Append("        ,@HorizontalTest1"); objParameter.Add("@HorizontalTest1", DbType.String, Item.HorizontalTest1);
            SbSql.Append("        ,@HorizontalTest2"); objParameter.Add("@HorizontalTest2", DbType.String, Item.HorizontalTest2);
            SbSql.Append("        ,@HorizontalTest3"); objParameter.Add("@HorizontalTest3", DbType.String, Item.HorizontalTest3);
            SbSql.Append("        ,@VerticalRate"); objParameter.Add("@VerticalRate", DbType.String, Item.VerticalRate);
            SbSql.Append("        ,@VerticalOriginal"); objParameter.Add("@VerticalOriginal", DbType.String, Item.VerticalOriginal);
            SbSql.Append("        ,@VerticalTest1"); objParameter.Add("@VerticalTest1", DbType.String, Item.VerticalTest1);
            SbSql.Append("        ,@VerticalTest2"); objParameter.Add("@VerticalTest2", DbType.String, Item.VerticalTest2);
            SbSql.Append("        ,@VerticalTest3"); objParameter.Add("@VerticalTest3", DbType.String, Item.VerticalTest3);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Laboratory - Heat Test(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Laboratory - Heat Test
        /// </summary>
        /// <param name="Item">Laboratory - Heat Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Update(FIR_Laboratory_Heat Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [FIR_Laboratory_Heat]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Roll != null) { SbSql.Append(",Roll=@Roll"+ Environment.NewLine); objParameter.Add("@Roll", DbType.String, Item.Roll);}
            if (Item.Dyelot != null) { SbSql.Append(",Dyelot=@Dyelot"+ Environment.NewLine); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);}
            if (Item.Inspdate != null) { SbSql.Append(",Inspdate=@Inspdate"+ Environment.NewLine); objParameter.Add("@Inspdate", DbType.String, Item.Inspdate);}
            if (Item.Inspector != null) { SbSql.Append(",Inspector=@Inspector"+ Environment.NewLine); objParameter.Add("@Inspector", DbType.String, Item.Inspector);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.HorizontalRate != null) { SbSql.Append(",HorizontalRate=@HorizontalRate"+ Environment.NewLine); objParameter.Add("@HorizontalRate", DbType.String, Item.HorizontalRate);}
            if (Item.HorizontalOriginal != null) { SbSql.Append(",HorizontalOriginal=@HorizontalOriginal"+ Environment.NewLine); objParameter.Add("@HorizontalOriginal", DbType.String, Item.HorizontalOriginal);}
            if (Item.HorizontalTest1 != null) { SbSql.Append(",HorizontalTest1=@HorizontalTest1"+ Environment.NewLine); objParameter.Add("@HorizontalTest1", DbType.String, Item.HorizontalTest1);}
            if (Item.HorizontalTest2 != null) { SbSql.Append(",HorizontalTest2=@HorizontalTest2"+ Environment.NewLine); objParameter.Add("@HorizontalTest2", DbType.String, Item.HorizontalTest2);}
            if (Item.HorizontalTest3 != null) { SbSql.Append(",HorizontalTest3=@HorizontalTest3"+ Environment.NewLine); objParameter.Add("@HorizontalTest3", DbType.String, Item.HorizontalTest3);}
            if (Item.VerticalRate != null) { SbSql.Append(",VerticalRate=@VerticalRate"+ Environment.NewLine); objParameter.Add("@VerticalRate", DbType.String, Item.VerticalRate);}
            if (Item.VerticalOriginal != null) { SbSql.Append(",VerticalOriginal=@VerticalOriginal"+ Environment.NewLine); objParameter.Add("@VerticalOriginal", DbType.String, Item.VerticalOriginal);}
            if (Item.VerticalTest1 != null) { SbSql.Append(",VerticalTest1=@VerticalTest1"+ Environment.NewLine); objParameter.Add("@VerticalTest1", DbType.String, Item.VerticalTest1);}
            if (Item.VerticalTest2 != null) { SbSql.Append(",VerticalTest2=@VerticalTest2"+ Environment.NewLine); objParameter.Add("@VerticalTest2", DbType.String, Item.VerticalTest2);}
            if (Item.VerticalTest3 != null) { SbSql.Append(",VerticalTest3=@VerticalTest3"+ Environment.NewLine); objParameter.Add("@VerticalTest3", DbType.String, Item.VerticalTest3);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Laboratory - Heat Test(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Laboratory - Heat Test
        /// </summary>
        /// <param name="Item">Laboratory - Heat Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Delete(FIR_Laboratory_Heat Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [FIR_Laboratory_Heat]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
