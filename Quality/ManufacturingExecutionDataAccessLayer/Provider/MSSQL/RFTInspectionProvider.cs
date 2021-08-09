using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class RFTInspectionProvider : SQLDAL, IRFTInspectionProvider
    {
        #region 底層連線
        public RFTInspectionProvider(string conString) : base(conString) { }
        public RFTInspectionProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base 

        public IList<RFT_Inspection> Get(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID", DbType.String, Item.FactoryID } ,
                { "@Line", DbType.String, Item.Line } ,
                { "@InspectionDate", DbType.DateTime, Item.InspectionDate } ,
            };

            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Line"+ Environment.NewLine);
            SbSql.Append("        ,FactoryID"+ Environment.NewLine);
            SbSql.Append("        ,StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,FixType"+ Environment.NewLine);
            SbSql.Append("        ,ReworkCardNo"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,ReworkCardType"+ Environment.NewLine);
            SbSql.Append("        ,InspectionDate"+ Environment.NewLine);
            SbSql.Append("        ,DisposeReason"+ Environment.NewLine);
            SbSql.Append("FROM [RFT_Inspection] r"+ Environment.NewLine);
            SbSql.Append("Where r.FactoryID = @FactoryID" + Environment.NewLine);

            if (string.IsNullOrEmpty(Item.Line)) { SbSql.Append("And r.Line = @Line" + Environment.NewLine); }

            if (Item.InspectionDate.HasValue) { 
                SbSql.Append("And ((r.AddDate >= @InspectionDate and r.AddDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate))) " + Environment.NewLine);
                SbSql.Append("  or (r.EditDate >= @InspectionDate and r.EditDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate)))) " + Environment.NewLine);
            }

            return ExecuteList<RFT_Inspection>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_Inspection]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Size"+ Environment.NewLine);
            SbSql.Append("        ,Line"+ Environment.NewLine);
            SbSql.Append("        ,FactoryID"+ Environment.NewLine);
            SbSql.Append("        ,StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,FixType"+ Environment.NewLine);
            SbSql.Append("        ,ReworkCardNo"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,ReworkCardType"+ Environment.NewLine);
            SbSql.Append("        ,InspectionDate"+ Environment.NewLine);
            SbSql.Append("        ,DisposeReason"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Size"); objParameter.Add("@Size", DbType.String, Item.Size);
            SbSql.Append("        ,@Line"); objParameter.Add("@Line", DbType.String, Item.Line);
            SbSql.Append("        ,@FactoryID"); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);
            SbSql.Append("        ,@StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@FixType"); objParameter.Add("@FixType", DbType.String, Item.FixType);
            SbSql.Append("        ,@ReworkCardNo"); objParameter.Add("@ReworkCardNo", DbType.String, Item.ReworkCardNo);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@ReworkCardType"); objParameter.Add("@ReworkCardType", DbType.String, Item.ReworkCardType);
            SbSql.Append("        ,@InspectionDate"); objParameter.Add("@InspectionDate", DbType.DateTime, Item.InspectionDate);
            SbSql.Append("        ,@DisposeReason"); objParameter.Add("@DisposeReason", DbType.String, Item.DisposeReason);
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
        public int Update(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_Inspection]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.OrderID != null) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.Size != null) { SbSql.Append(",Size=@Size"+ Environment.NewLine); objParameter.Add("@Size", DbType.String, Item.Size);}
            if (Item.Line != null) { SbSql.Append(",Line=@Line"+ Environment.NewLine); objParameter.Add("@Line", DbType.String, Item.Line);}
            if (Item.FactoryID != null) { SbSql.Append(",FactoryID=@FactoryID"+ Environment.NewLine); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);}
            if (Item.StyleUkey != null) { SbSql.Append(",StyleUkey=@StyleUkey"+ Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);}
            if (Item.FixType != null) { SbSql.Append(",FixType=@FixType"+ Environment.NewLine); objParameter.Add("@FixType", DbType.String, Item.FixType);}
            if (Item.ReworkCardNo != null) { SbSql.Append(",ReworkCardNo=@ReworkCardNo"+ Environment.NewLine); objParameter.Add("@ReworkCardNo", DbType.String, Item.ReworkCardNo);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine); objParameter.Add("@Status", DbType.String, Item.Status);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.ReworkCardType != null) { SbSql.Append(",ReworkCardType=@ReworkCardType"+ Environment.NewLine); objParameter.Add("@ReworkCardType", DbType.String, Item.ReworkCardType);}
            if (Item.InspectionDate != null) { SbSql.Append(",InspectionDate=@InspectionDate"+ Environment.NewLine); objParameter.Add("@InspectionDate", DbType.DateTime, Item.InspectionDate);}
            if (Item.DisposeReason != null) { SbSql.Append(",DisposeReason=@DisposeReason"+ Environment.NewLine); objParameter.Add("@DisposeReason", DbType.String, Item.DisposeReason);}
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
        public int Delete(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_Inspection]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
