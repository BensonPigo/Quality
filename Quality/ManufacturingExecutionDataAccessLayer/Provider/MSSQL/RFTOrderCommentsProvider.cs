using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class RFTOrderCommentsProvider : SQLDAL, IRFTOrderCommentsProvider
    {
        #region 底層連線
        public RFTOrderCommentsProvider(string conString) : base(conString) { }
        public RFTOrderCommentsProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<RFT_OrderComments_ViewModel> Get(RFT_OrderComments Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, Item.OrderID } ,
            };
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         oc.OrderID"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTCommentsID = dd.ID" + Environment.NewLine);
            SbSql.Append("        ,oc.Comnments"+ Environment.NewLine);
            SbSql.Append("        ,[PMS_RFTCommentsDescription] = dd.Description" + Environment.NewLine);
            SbSql.Append("from Production..DropdownList dd" + Environment.NewLine);
            SbSql.Append("left join RFT_OrderComments oc on dd.ID = oc.PMS_RFTCommentsID" + Environment.NewLine);
            SbSql.Append("where dd.Type='PMS_RFTComments'" + Environment.NewLine);
            SbSql.Append("and oc.OrderID = @OrderID" + Environment.NewLine);

            return ExecuteList<RFT_OrderComments_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(RFT_OrderComments Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_OrderComments]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         OrderID"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTCommentsID"+ Environment.NewLine);
            SbSql.Append("        ,Comnments"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@PMS_RFTCommentsID"); objParameter.Add("@PMS_RFTCommentsID", DbType.String, Item.PMS_RFTCommentsID);
            SbSql.Append("        ,@Comnments"); objParameter.Add("@Comnments", DbType.String, Item.Comnments);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(List<RFT_OrderComments> Items)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            int rowSeq = 1;
            foreach (var item in Items)
            {
                SbSql.Append("UPDATE [RFT_OrderComments]" + Environment.NewLine);
                SbSql.Append("SET" + Environment.NewLine);
                if (item.Comnments != null)
                {
                    SbSql.Append($"Comnments=@Comnments{rowSeq}" + Environment.NewLine);
                    objParameter.Add($"@Comnments{rowSeq}", DbType.String, item.Comnments);
                }

                SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
                if (!string.IsNullOrEmpty(item.OrderID))
                {
                    SbSql.Append($" and OrderID=@OrderID{rowSeq}" + Environment.NewLine);
                    objParameter.Add($"@OrderID{rowSeq}", DbType.String, item.OrderID);
                }
                if (!string.IsNullOrEmpty(item.PMS_RFTCommentsID))
                {
                    SbSql.Append($"and PMS_RFTCommentsID=@PMS_RFTCommentsID{rowSeq}" + Environment.NewLine);
                    objParameter.Add($"@PMS_RFTCommentsID{rowSeq}", DbType.String, item.PMS_RFTCommentsID);
                }

                rowSeq++;
            }

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(RFT_OrderComments Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_OrderComments]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Save_upd_ins(List<RFT_OrderComments> Items)
        {
            string sqlcmd = string.Empty;
            SQLParameterCollection objParameter = new SQLParameterCollection();

            int rowSeq = 1;
            foreach (var item in Items)
            {
                objParameter.Add($"@Comnments{rowSeq}", DbType.String, item.Comnments);
                objParameter.Add($"@OrderID{rowSeq}", DbType.String, item.OrderID);
                objParameter.Add($"@PMS_RFTCommentsID{rowSeq}", DbType.String, item.PMS_RFTCommentsID);

                sqlcmd += $@"
if exists(select 1 from RFT_OrderComments where OrderID = @OrderID{rowSeq} and PMS_RFTCommentsID = @PMS_RFTCommentsID{rowSeq})
begin
	UPDATE [RFT_OrderComments]
	set Comnments = @Comnments{rowSeq}
	where OrderID = @OrderID{rowSeq} and PMS_RFTCommentsID = @PMS_RFTCommentsID{rowSeq}
end
else
begin
	insert into [RFT_OrderComments](OrderID,PMS_RFTCommentsID,Comnments)
	values(@OrderID{rowSeq}, @PMS_RFTCommentsID{rowSeq}, @Comnments{rowSeq})
end
";
                rowSeq++;
            }

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }
        #endregion
    }
}
