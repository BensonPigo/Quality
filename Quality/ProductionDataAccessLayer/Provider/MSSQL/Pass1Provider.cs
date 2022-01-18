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
    public class Pass1Provider : SQLDAL, IPass1Provider
    {
        #region 底層連線
        public Pass1Provider(string ConString) : base(ConString) { }
        public Pass1Provider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public IList<Pass1> Get(Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, Item.ID } ,
                { "@Password", DbType.String, Item.Password } ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Name"+ Environment.NewLine);
            SbSql.Append("        ,Password"+ Environment.NewLine);
            SbSql.Append("        ,Position"+ Environment.NewLine);
            SbSql.Append("        ,FKPass0"+ Environment.NewLine);
            SbSql.Append("        ,IsAdmin"+ Environment.NewLine);
            SbSql.Append("        ,IsMIS"+ Environment.NewLine);
            SbSql.Append("        ,OrderGroup"+ Environment.NewLine);
            SbSql.Append("        ,EMail"+ Environment.NewLine);
            SbSql.Append("        ,ExtNo"+ Environment.NewLine);
            SbSql.Append("        ,OnBoard"+ Environment.NewLine);
            SbSql.Append("        ,Resign"+ Environment.NewLine);
            SbSql.Append("        ,Supervisor"+ Environment.NewLine);
            SbSql.Append("        ,Manager"+ Environment.NewLine);
            SbSql.Append("        ,Deputy"+ Environment.NewLine);
            SbSql.Append("        ,Factory"+ Environment.NewLine);
            SbSql.Append("        ,CodePage"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,LastLoginTime"+ Environment.NewLine);
            SbSql.Append("        ,ESignature"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("FROM [Pass1] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("Where ID=@ID" + Environment.NewLine);
            if (Item.Password != null) { SbSql.Append("and Password = @Password" + Environment.NewLine); }


            return ExecuteList<Pass1>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public int Update(Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Pass1]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Name != null) { SbSql.Append(",Name=@Name"+ Environment.NewLine); objParameter.Add("@Name", DbType.String, Item.Name);}
            if (Item.Password != null) { SbSql.Append(",Password=@Password"+ Environment.NewLine); objParameter.Add("@Password", DbType.String, Item.Password);}
            if (Item.Position != null) { SbSql.Append(",Position=@Position"+ Environment.NewLine); objParameter.Add("@Position", DbType.String, Item.Position);}
            if (Item.FKPass0 > 0) { SbSql.Append(",FKPass0=@FKPass0"+ Environment.NewLine); objParameter.Add("@FKPass0", DbType.String, Item.FKPass0);}
            if (Item.IsAdmin) { SbSql.Append(",IsAdmin=@IsAdmin"+ Environment.NewLine); objParameter.Add("@IsAdmin", DbType.String, Item.IsAdmin);}
            if (Item.IsMIS) { SbSql.Append(",IsMIS=@IsMIS"+ Environment.NewLine); objParameter.Add("@IsMIS", DbType.String, Item.IsMIS);}
            if (Item.OrderGroup != null) { SbSql.Append(",OrderGroup=@OrderGroup"+ Environment.NewLine); objParameter.Add("@OrderGroup", DbType.String, Item.OrderGroup);}
            if (Item.EMail != null) { SbSql.Append(",EMail=@EMail"+ Environment.NewLine); objParameter.Add("@EMail", DbType.String, Item.EMail);}
            if (Item.ExtNo != null) { SbSql.Append(",ExtNo=@ExtNo"+ Environment.NewLine); objParameter.Add("@ExtNo", DbType.String, Item.ExtNo);}
            if (Item.OnBoard != null) { SbSql.Append(",OnBoard=@OnBoard"+ Environment.NewLine); objParameter.Add("@OnBoard", DbType.DateTime, Item.OnBoard);}
            if (Item.Resign != null) { SbSql.Append(",Resign=@Resign"+ Environment.NewLine); objParameter.Add("@Resign", DbType.DateTime, Item.Resign);}
            if (Item.Supervisor != null) { SbSql.Append(",Supervisor=@Supervisor"+ Environment.NewLine); objParameter.Add("@Supervisor", DbType.String, Item.Supervisor);}
            if (Item.Manager != null) { SbSql.Append(",Manager=@Manager"+ Environment.NewLine); objParameter.Add("@Manager", DbType.String, Item.Manager);}
            if (Item.Deputy != null) { SbSql.Append(",Deputy=@Deputy"+ Environment.NewLine); objParameter.Add("@Deputy", DbType.String, Item.Deputy);}
            if (Item.Factory != null) { SbSql.Append(",Factory=@Factory"+ Environment.NewLine); objParameter.Add("@Factory", DbType.String, Item.Factory);}
            if (Item.CodePage != null) { SbSql.Append(",CodePage=@CodePage"+ Environment.NewLine); objParameter.Add("@CodePage", DbType.String, Item.CodePage);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.LastLoginTime != null) { SbSql.Append(",LastLoginTime=@LastLoginTime"+ Environment.NewLine); objParameter.Add("@LastLoginTime", DbType.DateTime, Item.LastLoginTime);}
            if (Item.ESignature != null) { SbSql.Append(",ESignature=@ESignature"+ Environment.NewLine); objParameter.Add("@ESignature", DbType.String, Item.ESignature);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark ?? "");}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/07/30  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/07/30  1.00    Admin        Create
        /// </history>
        public int Delete(Pass1 Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Pass1]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
