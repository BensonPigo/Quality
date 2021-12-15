using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailApperanceProvider : SQLDAL, IGarmentTestDetailApperanceProvider
    {
        #region 底層連線
        public GarmentTestDetailApperanceProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailApperanceProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<GarmentTest_Detail_Apperance_ViewModel> Get_GarmentTest_Detail_Apperance(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"
select  
ga.[ID]
,[No]
,[Type]
,[Wash1]
,[Wash2]
,[Wash3]
,[Comment]
,[Seq]
,[Wash4]
,[Wash5] 
,[WashName2] = IIF(WashName.Value is null,'3','10')
,[WashName3] = IIF(WashName.Value is null,'','15')
from GarmentTest_Detail_Apperance ga  WITH(NOLOCK)
inner join GarmentTest g on ga.ID = g.ID
outer apply(
	select Value =  r.Name 
	from Style s WITH(NOLOCK)
	inner join Reason r WITH(NOLOCK) on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
	and r.Name in ('MATCH TEAMWEAR','BASEBALL ON FIELD','SOFTBALL ON FIELD','TRAINING TEAMWEAR','LACROSSE ONFIELD','AMERIC. FOOT. ON-FIELD','TIRO','BASEBALL OFF FIELD','NCAA ON ICE','ON-COURT','BBALL PERFORMANCE','BRANDED BLANKS','SLD ON-FIELD','NHL ON ICE','SLD ON-COURT')
)WashName
where ga.ID = @ID
and ga.No = @No
";
            return ExecuteList<GarmentTest_Detail_Apperance_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Update_Apperance(List<GarmentTest_Detail_Apperance_ViewModel> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            int idx = 0;
            string sqlcmd = string.Empty;


            foreach (var item in source)
            {
                // Key
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@No{idx}", item.No));
                objParameter.Add(new SqlParameter($"@Seq{idx}", item.Seq));

                objParameter.Add(new SqlParameter($"@Type{idx}", string.IsNullOrEmpty(item.Type) ? string.Empty : item.Type));
                objParameter.Add(new SqlParameter($"@Wash1{idx}", string.IsNullOrEmpty(item.Wash1) ? string.Empty : item.Wash1));
                objParameter.Add(new SqlParameter($"@Wash2{idx}", string.IsNullOrEmpty(item.Wash2) ? string.Empty : item.Wash2));
                objParameter.Add(new SqlParameter($"@Wash3{idx}", string.IsNullOrEmpty(item.Wash3) ? string.Empty : item.Wash3));
                objParameter.Add(new SqlParameter($"@Wash4{idx}", string.IsNullOrEmpty(item.Wash4) ? string.Empty : item.Wash4));
                objParameter.Add(new SqlParameter($"@Wash5{idx}", string.IsNullOrEmpty(item.Wash5) ? string.Empty : item.Wash5));
                objParameter.Add(new SqlParameter($"@Comment{idx}", string.IsNullOrEmpty(item.Comment) ? string.Empty : item.Comment));

                sqlcmd += $@"
update GarmentTest_Detail_Apperance set
    [Type]  = @Type{idx},
	[Wash1] = @Wash1{idx},
    [Wash2]	= @Wash2{idx},
    [Wash3]	= @Wash3{idx},
    [Wash4]	= @Wash4{idx},
    [Wash5]	= @Wash5{idx},
    [Comment]	= @Comment{idx}
where ID = @ID{idx} and No = @No{idx} and Seq = @Seq{idx}
" + Environment.NewLine;
                idx++;
            }

            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

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
            SbSql.Append("FROM [GarmentTest_Detail_Apperance] WITH(NOLOCK)" + Environment.NewLine);



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
