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
    public class GarmentDetailSpiralityProvider : SQLDAL, IGarmentDetailSpiralityProvider
    {
        #region 底層連線
        public GarmentDetailSpiralityProvider(string ConString) : base(ConString) { }
        public GarmentDetailSpiralityProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<Garment_Detail_Spirality> Get(Garment_Detail_Spirality Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,MethodA_AAPrime"+ Environment.NewLine);
            SbSql.Append("        ,MethodA_APrimeB"+ Environment.NewLine);
            SbSql.Append("        ,MethodB_AAPrime"+ Environment.NewLine);
            SbSql.Append("        ,MethodB_AB"+ Environment.NewLine);
            SbSql.Append("        ,CM"+ Environment.NewLine);
            SbSql.Append("        ,MethodA"+ Environment.NewLine);
            SbSql.Append("        ,MethodB"+ Environment.NewLine);
            SbSql.Append("FROM [Garment_Detail_Spirality]"+ Environment.NewLine);



            return ExecuteList<Garment_Detail_Spirality>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Garment_Detail_Spirality Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Garment_Detail_Spirality]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,MethodA_AAPrime"+ Environment.NewLine);
            SbSql.Append("        ,MethodA_APrimeB"+ Environment.NewLine);
            SbSql.Append("        ,MethodB_AAPrime"+ Environment.NewLine);
            SbSql.Append("        ,MethodB_AB"+ Environment.NewLine);
            SbSql.Append("        ,CM"+ Environment.NewLine);
            SbSql.Append("        ,MethodA"+ Environment.NewLine);
            SbSql.Append("        ,MethodB"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@MethodA_AAPrime"); objParameter.Add("@MethodA_AAPrime", DbType.String, Item.MethodA_AAPrime);
            SbSql.Append("        ,@MethodA_APrimeB"); objParameter.Add("@MethodA_APrimeB", DbType.String, Item.MethodA_APrimeB);
            SbSql.Append("        ,@MethodB_AAPrime"); objParameter.Add("@MethodB_AAPrime", DbType.String, Item.MethodB_AAPrime);
            SbSql.Append("        ,@MethodB_AB"); objParameter.Add("@MethodB_AB", DbType.String, Item.MethodB_AB);
            SbSql.Append("        ,@CM"); objParameter.Add("@CM", DbType.String, Item.CM);
            SbSql.Append("        ,@MethodA"); objParameter.Add("@MethodA", DbType.String, Item.MethodA);
            SbSql.Append("        ,@MethodB"); objParameter.Add("@MethodB", DbType.String, Item.MethodB);
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
        public int Update(Garment_Detail_Spirality Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Garment_Detail_Spirality]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.MethodA_AAPrime != null) { SbSql.Append(",MethodA_AAPrime=@MethodA_AAPrime"+ Environment.NewLine); objParameter.Add("@MethodA_AAPrime", DbType.String, Item.MethodA_AAPrime);}
            if (Item.MethodA_APrimeB != null) { SbSql.Append(",MethodA_APrimeB=@MethodA_APrimeB"+ Environment.NewLine); objParameter.Add("@MethodA_APrimeB", DbType.String, Item.MethodA_APrimeB);}
            if (Item.MethodB_AAPrime != null) { SbSql.Append(",MethodB_AAPrime=@MethodB_AAPrime"+ Environment.NewLine); objParameter.Add("@MethodB_AAPrime", DbType.String, Item.MethodB_AAPrime);}
            if (Item.MethodB_AB != null) { SbSql.Append(",MethodB_AB=@MethodB_AB"+ Environment.NewLine); objParameter.Add("@MethodB_AB", DbType.String, Item.MethodB_AB);}
            if (Item.CM != null) { SbSql.Append(",CM=@CM"+ Environment.NewLine); objParameter.Add("@CM", DbType.String, Item.CM);}
            if (Item.MethodA != null) { SbSql.Append(",MethodA=@MethodA"+ Environment.NewLine); objParameter.Add("@MethodA", DbType.String, Item.MethodA);}
            if (Item.MethodB != null) { SbSql.Append(",MethodB=@MethodB"+ Environment.NewLine); objParameter.Add("@MethodB", DbType.String, Item.MethodB);}
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
        public int Delete(Garment_Detail_Spirality Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Garment_Detail_Spirality]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
