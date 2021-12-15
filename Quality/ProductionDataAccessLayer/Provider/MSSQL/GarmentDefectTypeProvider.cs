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
    public class GarmentDefectTypeProvider : SQLDAL, IGarmentDefectTypeProvider
    {
        #region 底層連線
        public GarmentDefectTypeProvider(string ConString) : base(ConString) { }
        public GarmentDefectTypeProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Defect type for REF/CFA(Garment)(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Defect type for REF/CFA(Garment)
        /// </summary>
        /// <param name="Item">Defect type for REF/CFA(Garment)成員</param>
        /// <returns>回傳Defect type for REF/CFA(Garment)</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<GarmentDefectType> Get()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,LocalDescription"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentDefectType] g WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("WHERE Junk = 0" + Environment.NewLine);
            SbSql.Append("AND EXISTS (SELECT 1 FROM GarmentDefectCode WITH(NOLOCK) WHERE g.ID = GarmentDefectTypeID AND Junk = 0 AND IsCFA  = 0)" + Environment.NewLine);
            SbSql.Append("ORDER BY Seq" + Environment.NewLine);


            return ExecuteList<GarmentDefectType>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Defect type for REF/CFA(Garment)(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Defect type for REF/CFA(Garment)
        /// </summary>
        /// <param name="Item">Defect type for REF/CFA(Garment)成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(GarmentDefectType Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentDefectType]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,LocalDescription"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@LocalDescription"); objParameter.Add("@LocalDescription", DbType.String, Item.LocalDescription);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.String, Item.Seq);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Defect type for REF/CFA(Garment)(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Defect type for REF/CFA(Garment)
        /// </summary>
        /// <param name="Item">Defect type for REF/CFA(Garment)成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(GarmentDefectType Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentDefectType]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.LocalDescription != null) { SbSql.Append(",LocalDescription=@LocalDescription"+ Environment.NewLine); objParameter.Add("@LocalDescription", DbType.String, Item.LocalDescription);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.String, Item.Seq);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Defect type for REF/CFA(Garment)(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Defect type for REF/CFA(Garment)
        /// </summary>
        /// <param name="Item">Defect type for REF/CFA(Garment)成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(GarmentDefectType Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentDefectType]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
