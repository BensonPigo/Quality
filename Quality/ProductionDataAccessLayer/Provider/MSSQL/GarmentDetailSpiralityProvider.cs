using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using DatabaseObject.ProductionDB;
using ADOHelper.Utility;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentDetailSpiralityProvider : SQLDAL, IGarmentDetailSpiralityProvider
    {
        #region 底層連線
        public GarmentDetailSpiralityProvider(string ConString) : base(ConString) { }
        public GarmentDetailSpiralityProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<Garment_Detail_Spirality> Get_Garment_Detail_Spirality(Int64 ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.Int64, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"
select 
[ID]
,[No]
,[Location]
,[MethodA_AAPrime]
,[MethodA_APrimeB]
,[MethodB_AAPrime]
,[MethodB_AB]
,[CM]
,[MethodA]
,[MethodB]
from Garment_Detail_Spirality
where ID = @ID
and No = @No
";
            return ExecuteList<Garment_Detail_Spirality>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Update_Spirality(List<Garment_Detail_Spirality> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            int idx = 0;
            string sqlcmd = string.Empty;

            foreach (var item in source)
            {
                // Key
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@No{idx}", item.No));
                objParameter.Add(new SqlParameter($"@Location{idx}", item.Location));

                objParameter.Add(new SqlParameter($"@MethodA_AAPrime{idx}", item.MethodA_AAPrime));
                objParameter.Add(new SqlParameter($"@MethodA_APrimeB{idx}", item.MethodA_APrimeB));
                objParameter.Add(new SqlParameter($"@MethodB_AAPrime{idx}", item.MethodB_AAPrime));
                objParameter.Add(new SqlParameter($"@MethodB_AB{idx}", item.MethodB_AB));
                objParameter.Add(new SqlParameter($"@CM{idx}", item.CM));
                objParameter.Add(new SqlParameter($"@MethodA{idx}", item.MethodA));
                objParameter.Add(new SqlParameter($"@MethodB{idx}", item.MethodB));

                sqlcmd += $@"
UPDATE [dbo].[Garment_Detail_Spirality]
   SET [MethodA_AAPrime] = @MethodA_AAPrime{idx}
      ,[MethodA_APrimeB] = @MethodA_APrimeB{idx}
      ,[MethodB_AAPrime] = @MethodB_AAPrime{idx}
      ,[MethodB_AB] = @MethodB_AB{idx}
      ,[CM] = @CM{idx}
      ,[MethodA] = @MethodA{idx}
      ,[MethodB] = @MethodB{idx}
WHERE id = @ID{idx} and No = @No{idx} and Location = @Location{idx}
" + Environment.NewLine;
            }

            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        /*回傳(Get) 詳細敘述如下*/
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
