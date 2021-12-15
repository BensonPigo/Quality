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
                { "@ID", DbType.String, Item.ID },
                { "@OrderID", DbType.String, Item.OrderID },
                { "@StyleUkey", DbType.String, Item.StyleUkey },
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
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.FactoryID)) { SbSql.Append("And r.FactoryID = @FactoryID" + Environment.NewLine); }

            if (Item.ID != 0) { SbSql.Append("And r.ID = @ID" + Environment.NewLine); }

            if (!string.IsNullOrEmpty(Item.Line)) { SbSql.Append("And r.Line = @Line" + Environment.NewLine); }

            if (Item.InspectionDate.HasValue) { 
                SbSql.Append("And ((r.AddDate >= @InspectionDate and r.AddDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate))) " + Environment.NewLine);
                SbSql.Append("  or (r.EditDate >= @InspectionDate and r.EditDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate)))) " + Environment.NewLine);
            }

            if (!string.IsNullOrEmpty(Item.OrderID)) { SbSql.Append("And r.OrderID = @OrderID" + Environment.NewLine); }
            if (Item.StyleUkey > 0) { SbSql.Append("And r.StyleUkey = @StyleUkey" + Environment.NewLine); }

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
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ID", DbType.String, Item.ID } ,
                { "@OrderID", DbType.String, Item.OrderID } ,
                { "@Article", DbType.String, Item.Article } ,
                { "@Location", DbType.String, Item.Location },
                { "@Size", DbType.String, Item.Size },
                { "@Line", DbType.String, Item.Line },

                { "@FactoryID", DbType.String, Item.FactoryID} ,
                { "@StyleUkey", DbType.String, Item.StyleUkey } ,
                { "@FixType", DbType.String, Item.FixType } ,
                { "@ReworkCardNo", DbType.String, Item.ReworkCardNo },
                { "@Status", DbType.String, Item.Status },

                { "@AddDate", DbType.DateTime, Item.AddDate },
                { "@AddName", DbType.String, Item.AddName } ,
                { "@EditDate", DbType.DateTime, Item.EditDate } ,
                { "@EditName", DbType.String, Item.EditName } ,
                { "@ReworkCardType", DbType.String, Item.ReworkCardType },
                { "@InspectionDate", DbType.DateTime, Item.InspectionDate },
                { "@DisposeReason", DbType.String, Item.DisposeReason },
            };


            SbSql.Append("UPDATE [RFT_Inspection]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            { SbSql.Append("ID=ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.OrderID)) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.Article)) { SbSql.Append(",Article=@Article"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.Location)) { SbSql.Append(",Location=@Location"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.Size)) { SbSql.Append(",Size=@Size" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Line)) { SbSql.Append(",Line=@Line"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.FactoryID)) { SbSql.Append(",FactoryID=@FactoryID"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.StyleUkey.ToString())) { SbSql.Append(",StyleUkey=@StyleUkey"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.FixType)) { SbSql.Append(",FixType=@FixType"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.ReworkCardNo)) { SbSql.Append(",ReworkCardNo=@ReworkCardNo"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.Status)) { SbSql.Append(",Status=@Status"+ Environment.NewLine);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.AddName)) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.EditName)) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.ReworkCardType)) { SbSql.Append(",ReworkCardType=@ReworkCardType"+ Environment.NewLine);}
            if (Item.InspectionDate != null) { SbSql.Append(",InspectionDate=@InspectionDate"+ Environment.NewLine);}
            if (!string.IsNullOrEmpty(Item.DisposeReason)) { SbSql.Append(",DisposeReason=@DisposeReason"+ Environment.NewLine);}

            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ID.ToString())) { SbSql.Append(" and ID=@ID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.OrderID)) { SbSql.Append(" and OrderID=@OrderID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Article)) { SbSql.Append(" and Article=@Article" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Location)) { SbSql.Append(" and Location=@Location" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Size)) { SbSql.Append(" and Size=@Size" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Line)) { SbSql.Append(" and Line=@Line" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.FactoryID)) { SbSql.Append(" and FactoryID=@FactoryID" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.StyleUkey.ToString())) { SbSql.Append(" and StyleUkey=@StyleUkey" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.FixType)) { SbSql.Append(" and FixType=@FixType" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.ReworkCardNo)) { SbSql.Append(" and ReworkCardNo=@ReworkCardNo" + Environment.NewLine); }
            if (!string.IsNullOrEmpty(Item.Status)) { SbSql.Append(" and Status=@Status" + Environment.NewLine); }
            
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
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, Item.ID } ,
            };
            SbSql.Append("DELETE FROM [RFT_Inspection]" + Environment.NewLine);
            SbSql.Append("where 1=1" + Environment.NewLine);
            SbSql.Append("and id = @ID" + Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int SaveReworkListAction(List<RFT_Inspection> items, string statusType)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            // 調整Repl 無法加.的問題
            statusType = statusType.ToLower() == "repl" ? "Repl." : statusType;
            int rowcnt = 1;
            string sqlcmd = "";
            foreach (var item in items)
            {
                // sql Parameter
                objParameter.Add($"@ID{rowcnt}", string.IsNullOrEmpty(item.ID.ToString()) ? "" : item.ID.ToString());
                objParameter.Add($"@FixType{rowcnt}", string.IsNullOrEmpty(statusType) ? "" : statusType);
                objParameter.Add($"@EditName{rowcnt}", string.IsNullOrEmpty(item.EditName.ToString()) ? "" : item.EditName.ToString());
                objParameter.Add($"@inspectDate{rowcnt}", item.InspectionDate == null ? "NULL" : ((DateTime)item.InspectionDate).ToString("d"));
                objParameter.Add($"@DisposeReason{rowcnt}", string.IsNullOrEmpty(item.DisposeReason) ? "" : item.DisposeReason.ToString());

                if (statusType.ToLower() == "pass")
                {
                    sqlcmd += $@"
update RFT_Inspection
set	status = 'Fixed'
,InspectionDate = @inspectDate{rowcnt}
,EditName = @EditName{rowcnt} , EditDate = GETDATE()
where ID = @ID{rowcnt}

update rc
set Status = 'Fixed'
,EditName = @EditName{rowcnt}, EditDate = GETDATE()
from ReworkCard rc
inner join RFT_Inspection inp on inp.ReworkCardNo = rc.No
and inp.ReworkCardType = rc.Type and inp.Line = rc.Line
and inp.FactoryID = rc.FactoryID
where inp.ID = @ID{rowcnt}
";
                }
                else if(statusType.ToLower() == "wash" ||
                    statusType.ToLower() == "repl." ||
                    statusType.ToLower() == "print" ||
                    statusType.ToLower() == "shade")
                {
                    sqlcmd += $@"
update RFT_Inspection
set	FixType = @FixType{rowcnt}
,EditName = @EditName{rowcnt} , EditDate = GETDATE()
where ID = @ID{rowcnt}

update rc
set Status = 'Fixed'
,EditName = @EditName{rowcnt}, EditDate = GETDATE()
from ReworkCard rc
inner join RFT_Inspection inp on inp.ReworkCardNo = rc.No
and inp.ReworkCardType = rc.Type and inp.Line = rc.Line
and inp.FactoryID = rc.FactoryID
where inp.ID = @ID{rowcnt}
";
                }

                else if (statusType.ToLower() == "dispose")
                {
                    sqlcmd += $@"
update RFT_Inspection
set	Status = @FixType{rowcnt}
,EditName = @EditName{rowcnt} , EditDate = GETDATE()
,DisposeReason = @DisposeReason{rowcnt}
where ID = @ID{rowcnt}

update rc
set Status = 'Fixed'
,EditName = @EditName{rowcnt}, EditDate = GETDATE()
from ReworkCard rc
inner join RFT_Inspection inp on inp.ReworkCardNo = rc.No
and inp.ReworkCardType = rc.Type and inp.Line = rc.Line
and inp.FactoryID = rc.FactoryID
where inp.ID = @ID{rowcnt}
";
                }
               
                rowcnt++;
            }

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        public int SaveReworkListDelete(List<RFT_Inspection> items)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            int rowcnt = 1;
            string sqlcmd = "";
            foreach (var item in items)
            {
                // sql Parameter
                objParameter.Add($"@ID{rowcnt}", string.IsNullOrEmpty(item.ID.ToString()) ? "" : item.ID.ToString());
                objParameter.Add($"@EditName{rowcnt}", string.IsNullOrEmpty(item.EditName) ? "" : item.EditName);

                sqlcmd = $@"
SET XACT_ABORT ON
update rc
set Status = 'Fixed'
,EditName = @EditName{rowcnt}, EditDate = GETDATE()
from ReworkCard rc
inner join RFT_Inspection inp on inp.ReworkCardNo = rc.No
    and inp.ReworkCardType = rc.Type and inp.Line = rc.Line
    and inp.FactoryID = rc.FactoryID
where inp.ID = @ID{rowcnt}

delete from RFT_Inspection_Detail
where id = @id{rowcnt}

delete from [ExtendServer].PMSFile.dbo.RFT_Inspection_Detail
where id = @id{rowcnt}

delete from RFT_Inspection
where id = @id{rowcnt}
";
                rowcnt++;
            }

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }
        #endregion
    }
}
