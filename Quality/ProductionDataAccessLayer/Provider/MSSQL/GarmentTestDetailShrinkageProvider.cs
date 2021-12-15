using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailShrinkageProvider : SQLDAL, IGarmentTestDetailShrinkageProvider
    {
        #region 底層連線
        public GarmentTestDetailShrinkageProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailShrinkageProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base

        public IList<GarmentTest_Detail_Shrinkage> Get_GarmentTest_Detail_Shrinkage(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"

select 
ID,No
,[Location] = case [Location] 
	when 'T' then 'Top' 
	when 'I' then 'INNER'
	when 'O' then 'OUTER'
	when 'B' then 'BOTTOM'
	ELSE '' end
,[Type]
,[BeforeWash]
,[SizeSpec]
,[AfterWash1]
,[Shrinkage1]
,[AfterWash2]
,[Shrinkage2]
,[AfterWash3]
,[Shrinkage3]
,[Seq]
from GarmentTest_Detail_Shrinkage WITH(NOLOCK)
where ID = @ID
and No = @No
";

            return ExecuteList<GarmentTest_Detail_Shrinkage>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Update_GarmentTestShrinkage(List<GarmentTest_Detail_Shrinkage> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            int idx = 0;
            string sqlcmd = string.Empty;


            foreach (var item in source)
            {
                // 畫面資料 轉換成DB資料
                if (item.Location.ToUpper() == "BOTTOM") { item.Location = "B"; }
                if (item.Location.ToUpper() == "TOP") { item.Location = "T"; }
                if (item.Location.ToUpper() == "INNER") { item.Location = "I"; }
                if (item.Location.ToUpper() == "OUTER") { item.Location = "O"; }

                // Key
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@No{idx}", item.No));
                objParameter.Add(new SqlParameter($"@Location{idx}", item.Location));
                objParameter.Add(new SqlParameter($"@Type{idx}", item.Type));

                objParameter.Add(new SqlParameter($"@BeforeWash{idx}", item.BeforeWash));
                objParameter.Add(new SqlParameter($"@SizeSpec{idx}", item.SizeSpec));

                objParameter.Add(new SqlParameter($"@AfterWash1{idx}", item.AfterWash1));
                objParameter.Add(new SqlParameter($"@Shrinkage1{idx}", item.Shrinkage1));

                objParameter.Add(new SqlParameter($"@AfterWash2{idx}", item.AfterWash2));
                objParameter.Add(new SqlParameter($"@Shrinkage2{idx}", item.Shrinkage2));

                objParameter.Add(new SqlParameter($"@AfterWash3{idx}", item.AfterWash3));
                objParameter.Add(new SqlParameter($"@Shrinkage3{idx}", item.Shrinkage3));


                sqlcmd += $@"
update GarmentTest_Detail_Shrinkage set
    [BeforeWash]    = @BeforeWash{idx},
    [SizeSpec]	    = @SizeSpec{idx},
    [AfterWash1]	= @AfterWash1{idx},
    [Shrinkage1]	= @Shrinkage1{idx},
    [AfterWash2]	= @AfterWash2{idx},
    [Shrinkage2]	= @Shrinkage2{idx},
    [AfterWash3]	= @AfterWash3{idx},
    [Shrinkage3]	= @Shrinkage3{idx}
where ID = @ID{idx} and No = @No{idx} and Type = @Type{idx} and Location = @Location{idx}
" + Environment.NewLine;
                idx++;
            }

            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public DataTable Get_dt_Shrinkage(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"

select [ID],[No]
,[Location]
,[Type],[BeforeWash],[SizeSpec],[AfterWash1],[Shrinkage1],[AfterWash2],[Shrinkage2],[AfterWash3],[Shrinkage3]
from GarmentTest_Detail_Shrinkage WITH(NOLOCK)
where ID = @ID
and No = @No
order by Location desc, seq
";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
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
        public int Create(GarmentTest_Detail_Shrinkage Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest_Detail_Shrinkage]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,BeforeWash"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,AfterWash1"+ Environment.NewLine);
            SbSql.Append("        ,Shrinkage1"+ Environment.NewLine);
            SbSql.Append("        ,AfterWash2"+ Environment.NewLine);
            SbSql.Append("        ,Shrinkage2"+ Environment.NewLine);
            SbSql.Append("        ,AfterWash3"+ Environment.NewLine);
            SbSql.Append("        ,Shrinkage3"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@BeforeWash"); objParameter.Add("@BeforeWash", DbType.String, Item.BeforeWash);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@AfterWash1"); objParameter.Add("@AfterWash1", DbType.String, Item.AfterWash1);
            SbSql.Append("        ,@Shrinkage1"); objParameter.Add("@Shrinkage1", DbType.String, Item.Shrinkage1);
            SbSql.Append("        ,@AfterWash2"); objParameter.Add("@AfterWash2", DbType.String, Item.AfterWash2);
            SbSql.Append("        ,@Shrinkage2"); objParameter.Add("@Shrinkage2", DbType.String, Item.Shrinkage2);
            SbSql.Append("        ,@AfterWash3"); objParameter.Add("@AfterWash3", DbType.String, Item.AfterWash3);
            SbSql.Append("        ,@Shrinkage3"); objParameter.Add("@Shrinkage3", DbType.String, Item.Shrinkage3);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.String, Item.Seq);
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
        public int Update(GarmentTest_Detail_Shrinkage Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest_Detail_Shrinkage]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.BeforeWash != null) { SbSql.Append(",BeforeWash=@BeforeWash"+ Environment.NewLine); objParameter.Add("@BeforeWash", DbType.String, Item.BeforeWash);}
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec"+ Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);}
            if (Item.AfterWash1 != null) { SbSql.Append(",AfterWash1=@AfterWash1"+ Environment.NewLine); objParameter.Add("@AfterWash1", DbType.String, Item.AfterWash1);}
            if (Item.Shrinkage1 != null) { SbSql.Append(",Shrinkage1=@Shrinkage1"+ Environment.NewLine); objParameter.Add("@Shrinkage1", DbType.String, Item.Shrinkage1);}
            if (Item.AfterWash2 != null) { SbSql.Append(",AfterWash2=@AfterWash2"+ Environment.NewLine); objParameter.Add("@AfterWash2", DbType.String, Item.AfterWash2);}
            if (Item.Shrinkage2 != null) { SbSql.Append(",Shrinkage2=@Shrinkage2"+ Environment.NewLine); objParameter.Add("@Shrinkage2", DbType.String, Item.Shrinkage2);}
            if (Item.AfterWash3 != null) { SbSql.Append(",AfterWash3=@AfterWash3"+ Environment.NewLine); objParameter.Add("@AfterWash3", DbType.String, Item.AfterWash3);}
            if (Item.Shrinkage3 != null) { SbSql.Append(",Shrinkage3=@Shrinkage3"+ Environment.NewLine); objParameter.Add("@Shrinkage3", DbType.String, Item.Shrinkage3);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.String, Item.Seq);}
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
        public int Delete(GarmentTest_Detail_Shrinkage Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest_Detail_Shrinkage]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
