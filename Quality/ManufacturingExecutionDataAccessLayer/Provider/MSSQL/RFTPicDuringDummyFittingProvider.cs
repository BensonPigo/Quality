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
    public class RFTPicDuringDummyFittingProvider : SQLDAL, IRFTPicDuringDummyFittingProvider
    {
        #region 底層連線
        public RFTPicDuringDummyFittingProvider(string conString) : base(conString) { }
        public RFTPicDuringDummyFittingProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<RFT_PicDuringDummyFitting> Get(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@OrderID", DbType.String, Item.OrderID},
                { "@Article", DbType.String, Item.Article},
                { "@Size", DbType.String, Item.Size} ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Front"+ Environment.NewLine);
            SbSql.Append("        ,Side"+ Environment.NewLine);
            SbSql.Append("        ,Back"+ Environment.NewLine);
            SbSql.Append("FROM [ExtendServer].PMSFile.dbo.[RFT_PicDuringDummyFitting]  WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.OrderID)) { SbSql.Append(" and OrderID = @OrderID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Article)) { SbSql.Append(" and Article = @Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Size)) { SbSql.Append(" and Size = @Size" + Environment.NewLine); }

            return ExecuteList<RFT_PicDuringDummyFitting>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_PicDuringDummyFitting]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Front"+ Environment.NewLine);
            SbSql.Append("        ,Side"+ Environment.NewLine);
            SbSql.Append("        ,Back"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Size"); objParameter.Add("@Size", DbType.String, Item.Size);
            SbSql.Append("        ,@Front"); objParameter.Add("@Front", DbType.String, Item.Front);
            SbSql.Append("        ,@Side"); objParameter.Add("@Side", DbType.String, Item.Side);
            SbSql.Append("        ,@Back"); objParameter.Add("@Back", DbType.String, Item.Back);
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
        public int Update(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_PicDuringDummyFitting]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.OrderID != null) { SbSql.Append("OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Size != null) { SbSql.Append(",Size=@Size"+ Environment.NewLine); objParameter.Add("@Size", DbType.String, Item.Size);}
            if (Item.Front != null) { SbSql.Append(",Front=@Front"+ Environment.NewLine); objParameter.Add("@Front", DbType.String, Item.Front);}
            if (Item.Side != null) { SbSql.Append(",Side=@Side"+ Environment.NewLine); objParameter.Add("@Side", DbType.String, Item.Side);}
            if (Item.Back != null) { SbSql.Append(",Back=@Back"+ Environment.NewLine); objParameter.Add("@Back", DbType.String, Item.Back);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




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
        public int Delete(RFT_PicDuringDummyFitting Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_PicDuringDummyFitting]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Save_Upd_Ins(RFT_PicDuringDummyFitting Item)
        {
            string sqlcmd = string.Empty;
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@OrderID", DbType.String, Item.OrderID},
                { "@Article", DbType.String, Item.Article},
                { "@Size", DbType.String, Item.Size} ,
            };

            if (Item.Front != null){objParameter.Add("@Front", Item.Front);}
            else{objParameter.Add("@Front", System.Data.SqlTypes.SqlBinary.Null);}

            if (Item.Side != null){objParameter.Add("@Side", Item.Side);}
            else{objParameter.Add("@Side", System.Data.SqlTypes.SqlBinary.Null);}

            if (Item.Back != null){objParameter.Add("@Back", Item.Back);}
            else{objParameter.Add("@Back", System.Data.SqlTypes.SqlBinary.Null);}

            sqlcmd += $@"

SET XACT_ABORT ON
if exists(select 1 from RFT_PicDuringDummyFitting WITH(NOLOCK) where OrderID = @OrderID and Article = @Article and Size = @Size)
begin
	UPDATE [RFT_PicDuringDummyFitting]
	set Front = @Front
    ,Side = @Side
    ,Back = @Back
	where OrderID = @OrderID and Article = @Article and Size = @Size

	UPDATE [ExtendServer].PMSFile.dbo.[RFT_PicDuringDummyFitting]
	set Front = @Front
    ,Side = @Side
    ,Back = @Back
	where OrderID = @OrderID and Article = @Article and Size = @Size
end
else
begin
	insert into RFT_PicDuringDummyFitting(OrderID,Article,Size,Front,Side,Back)
	values(@OrderID, @Article,@Size,@Front,@Side,@Back)

	insert into [ExtendServer].PMSFile.dbo.RFT_PicDuringDummyFitting(OrderID,Article,Size,Front,Side,Back)
	values(@OrderID, @Article,@Size,@Front,@Side,@Back)
end
";


            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }
        #endregion


        public IList<RFT_PicDuringDummyFitting> Get_PicDuringDummy_Result(RFT_PicDuringDummyFitting_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            SbSql.Append($@"
SELECT oq.Article
	,Size = oq.SizeCode
	,p.Front
	,p.Side
	,p.Back
from SciProduction_Order_Qty oq WITH(NOLOCK) 
left join [ExtendServer].PMSFile.dbo.RFT_PicDuringDummyFitting p WITH(NOLOCK) ON oq.ID = p.OrderID AND oq.Article = p.Article AND oq.SizeCode=p.Size
WHERE 1=1
");
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql.Append(" AND oq.ID = @OrderID" + Environment.NewLine);
                objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            }

            return ExecuteList<RFT_PicDuringDummyFitting>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<RFT_PicDuringDummyFitting_ViewModel> Check_OrderID_Exists(string OrderID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, OrderID);

            SbSql.Append($@"
SELECT OrderID = ID, StyleID ,OrderTypeID ,SeasonID
FROM SciProduction_Orders o WITH(NOLOCK)
where o.Junk=0
and o.Category='S'
--and o.OnSiteSample!=1
and o.ID = @OrderID
");

            return ExecuteList<RFT_PicDuringDummyFitting_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
