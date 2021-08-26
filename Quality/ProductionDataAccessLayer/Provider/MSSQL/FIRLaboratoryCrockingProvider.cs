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
    public class FIRLaboratoryCrockingProvider : SQLDAL, IFIRLaboratoryCrockingProvider
    {
        #region 底層連線
        public FIRLaboratoryCrockingProvider(string ConString) : base(ConString) { }
        public FIRLaboratoryCrockingProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Laboratory - Shade bone Inspection(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Laboratory - Shade bone Inspection
        /// </summary>
        /// <param name="Item">Laboratory - Shade bone Inspection成員</param>
        /// <returns>回傳Laboratory - Shade bone Inspection</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public IList<FIR_Laboratory_Crocking> Get(FIR_Laboratory_Crocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,DryScale"+ Environment.NewLine);
            SbSql.Append("        ,WetScale"+ Environment.NewLine);
            SbSql.Append("        ,Inspdate"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,ResultDry"+ Environment.NewLine);
            SbSql.Append("        ,ResultWet"+ Environment.NewLine);
            SbSql.Append("        ,DryScale_Weft"+ Environment.NewLine);
            SbSql.Append("        ,WetScale_Weft"+ Environment.NewLine);
            SbSql.Append("        ,ResultDry_Weft"+ Environment.NewLine);
            SbSql.Append("        ,ResultWet_Weft"+ Environment.NewLine);
            SbSql.Append("FROM [FIR_Laboratory_Crocking]"+ Environment.NewLine);



            return ExecuteList<FIR_Laboratory_Crocking>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Laboratory - Shade bone Inspection(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Laboratory - Shade bone Inspection
        /// </summary>
        /// <param name="Item">Laboratory - Shade bone Inspection成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Create(FIR_Laboratory_Crocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [FIR_Laboratory_Crocking]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,DryScale"+ Environment.NewLine);
            SbSql.Append("        ,WetScale"+ Environment.NewLine);
            SbSql.Append("        ,Inspdate"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,ResultDry"+ Environment.NewLine);
            SbSql.Append("        ,ResultWet"+ Environment.NewLine);
            SbSql.Append("        ,DryScale_Weft"+ Environment.NewLine);
            SbSql.Append("        ,WetScale_Weft"+ Environment.NewLine);
            SbSql.Append("        ,ResultDry_Weft"+ Environment.NewLine);
            SbSql.Append("        ,ResultWet_Weft"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Roll"); objParameter.Add("@Roll", DbType.String, Item.Roll);
            SbSql.Append("        ,@Dyelot"); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);
            SbSql.Append("        ,@DryScale"); objParameter.Add("@DryScale", DbType.String, Item.DryScale);
            SbSql.Append("        ,@WetScale"); objParameter.Add("@WetScale", DbType.String, Item.WetScale);
            SbSql.Append("        ,@Inspdate"); objParameter.Add("@Inspdate", DbType.String, Item.Inspdate);
            SbSql.Append("        ,@Inspector"); objParameter.Add("@Inspector", DbType.String, Item.Inspector);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@ResultDry"); objParameter.Add("@ResultDry", DbType.String, Item.ResultDry);
            SbSql.Append("        ,@ResultWet"); objParameter.Add("@ResultWet", DbType.String, Item.ResultWet);
            SbSql.Append("        ,@DryScale_Weft"); objParameter.Add("@DryScale_Weft", DbType.String, Item.DryScale_Weft);
            SbSql.Append("        ,@WetScale_Weft"); objParameter.Add("@WetScale_Weft", DbType.String, Item.WetScale_Weft);
            SbSql.Append("        ,@ResultDry_Weft"); objParameter.Add("@ResultDry_Weft", DbType.String, Item.ResultDry_Weft);
            SbSql.Append("        ,@ResultWet_Weft"); objParameter.Add("@ResultWet_Weft", DbType.String, Item.ResultWet_Weft);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Laboratory - Shade bone Inspection(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Laboratory - Shade bone Inspection
        /// </summary>
        /// <param name="Item">Laboratory - Shade bone Inspection成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Update(FIR_Laboratory_Crocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [FIR_Laboratory_Crocking]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Roll != null) { SbSql.Append(",Roll=@Roll"+ Environment.NewLine); objParameter.Add("@Roll", DbType.String, Item.Roll);}
            if (Item.Dyelot != null) { SbSql.Append(",Dyelot=@Dyelot"+ Environment.NewLine); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);}
            if (Item.DryScale != null) { SbSql.Append(",DryScale=@DryScale"+ Environment.NewLine); objParameter.Add("@DryScale", DbType.String, Item.DryScale);}
            if (Item.WetScale != null) { SbSql.Append(",WetScale=@WetScale"+ Environment.NewLine); objParameter.Add("@WetScale", DbType.String, Item.WetScale);}
            if (Item.Inspdate != null) { SbSql.Append(",Inspdate=@Inspdate"+ Environment.NewLine); objParameter.Add("@Inspdate", DbType.String, Item.Inspdate);}
            if (Item.Inspector != null) { SbSql.Append(",Inspector=@Inspector"+ Environment.NewLine); objParameter.Add("@Inspector", DbType.String, Item.Inspector);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.ResultDry != null) { SbSql.Append(",ResultDry=@ResultDry"+ Environment.NewLine); objParameter.Add("@ResultDry", DbType.String, Item.ResultDry);}
            if (Item.ResultWet != null) { SbSql.Append(",ResultWet=@ResultWet"+ Environment.NewLine); objParameter.Add("@ResultWet", DbType.String, Item.ResultWet);}
            if (Item.DryScale_Weft != null) { SbSql.Append(",DryScale_Weft=@DryScale_Weft"+ Environment.NewLine); objParameter.Add("@DryScale_Weft", DbType.String, Item.DryScale_Weft);}
            if (Item.WetScale_Weft != null) { SbSql.Append(",WetScale_Weft=@WetScale_Weft"+ Environment.NewLine); objParameter.Add("@WetScale_Weft", DbType.String, Item.WetScale_Weft);}
            if (Item.ResultDry_Weft != null) { SbSql.Append(",ResultDry_Weft=@ResultDry_Weft"+ Environment.NewLine); objParameter.Add("@ResultDry_Weft", DbType.String, Item.ResultDry_Weft);}
            if (Item.ResultWet_Weft != null) { SbSql.Append(",ResultWet_Weft=@ResultWet_Weft"+ Environment.NewLine); objParameter.Add("@ResultWet_Weft", DbType.String, Item.ResultWet_Weft);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Laboratory - Shade bone Inspection(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Laboratory - Shade bone Inspection
        /// </summary>
        /// <param name="Item">Laboratory - Shade bone Inspection成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Delete(FIR_Laboratory_Crocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [FIR_Laboratory_Crocking]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
