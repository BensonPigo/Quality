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
    public class RFTInspectionDetailProvider : SQLDAL, IRFTInspectionDetailProvider
    {
        #region 底層連線
        public RFTInspectionDetailProvider(string conString) : base(conString) { }
        public RFTInspectionDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<RFT_Inspection_Detail> Top3Defects(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID", DbType.String, Item.FactoryID } ,
                { "@Line", DbType.String, Item.Line } ,
                { "@InspectionDate", DbType.DateTime, Item.InspectionDate } ,
            };

            SbSql.Append(
                @"
select rd.DefectCode, rd.AreaCode
into #tmp
from RFT_Inspection r
inner join RFT_Inspection_Detail rd on r.ID = rd.ID and rd.Junk = 0" + Environment.NewLine);


            SbSql.Append("Where r.FactoryID = @FactoryID" + Environment.NewLine);

            if (string.IsNullOrEmpty(Item.Line)) { SbSql.Append("And r.Line = @Line" + Environment.NewLine); }

            if (Item.InspectionDate.HasValue)
            {
                SbSql.Append("And ((r.AddDate >= @InspectionDate and r.AddDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate))) " + Environment.NewLine);
                SbSql.Append("  or (r.EditDate >= @InspectionDate and r.EditDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate)))) " + Environment.NewLine);
            }

            return ExecuteList<RFT_Inspection_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }

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
        public IList<RFT_Inspection_Detail> Get(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,DefectCode"+ Environment.NewLine);
            SbSql.Append("        ,AreaCode"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTBACriteriaID"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTRespID"+ Environment.NewLine);
            SbSql.Append("        ,GarmentDefectTypeID"+ Environment.NewLine);
            SbSql.Append("        ,GarmentDefectCodeID"+ Environment.NewLine);
            SbSql.Append("        ,DefectPicture"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("FROM [RFT_Inspection_Detail]"+ Environment.NewLine);



            return ExecuteList<RFT_Inspection_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_Inspection_Detail]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,DefectCode"+ Environment.NewLine);
            SbSql.Append("        ,AreaCode"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTBACriteriaID"+ Environment.NewLine);
            SbSql.Append("        ,PMS_RFTRespID"+ Environment.NewLine);
            SbSql.Append("        ,GarmentDefectTypeID"+ Environment.NewLine);
            SbSql.Append("        ,GarmentDefectCodeID"+ Environment.NewLine);
            SbSql.Append("        ,DefectPicture"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@DefectCode"); objParameter.Add("@DefectCode", DbType.String, Item.DefectCode);
            SbSql.Append("        ,@AreaCode"); objParameter.Add("@AreaCode", DbType.String, Item.AreaCode);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@PMS_RFTBACriteriaID"); objParameter.Add("@PMS_RFTBACriteriaID", DbType.String, Item.PMS_RFTBACriteriaID);
            SbSql.Append("        ,@PMS_RFTRespID"); objParameter.Add("@PMS_RFTRespID", DbType.String, Item.PMS_RFTRespID);
            SbSql.Append("        ,@GarmentDefectTypeID"); objParameter.Add("@GarmentDefectTypeID", DbType.String, Item.GarmentDefectTypeID);
            SbSql.Append("        ,@GarmentDefectCodeID"); objParameter.Add("@GarmentDefectCodeID", DbType.String, Item.GarmentDefectCodeID);
            SbSql.Append("        ,@DefectPicture"); objParameter.Add("@DefectPicture", DbType.String, Item.DefectPicture);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
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
        public int Update(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_Inspection_Detail]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Ukey != null) { SbSql.Append(",Ukey=@Ukey"+ Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey);}
            if (Item.DefectCode != null) { SbSql.Append(",DefectCode=@DefectCode"+ Environment.NewLine); objParameter.Add("@DefectCode", DbType.String, Item.DefectCode);}
            if (Item.AreaCode != null) { SbSql.Append(",AreaCode=@AreaCode"+ Environment.NewLine); objParameter.Add("@AreaCode", DbType.String, Item.AreaCode);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.PMS_RFTBACriteriaID != null) { SbSql.Append(",PMS_RFTBACriteriaID=@PMS_RFTBACriteriaID"+ Environment.NewLine); objParameter.Add("@PMS_RFTBACriteriaID", DbType.String, Item.PMS_RFTBACriteriaID);}
            if (Item.PMS_RFTRespID != null) { SbSql.Append(",PMS_RFTRespID=@PMS_RFTRespID"+ Environment.NewLine); objParameter.Add("@PMS_RFTRespID", DbType.String, Item.PMS_RFTRespID);}
            if (Item.GarmentDefectTypeID != null) { SbSql.Append(",GarmentDefectTypeID=@GarmentDefectTypeID"+ Environment.NewLine); objParameter.Add("@GarmentDefectTypeID", DbType.String, Item.GarmentDefectTypeID);}
            if (Item.GarmentDefectCodeID != null) { SbSql.Append(",GarmentDefectCodeID=@GarmentDefectCodeID"+ Environment.NewLine); objParameter.Add("@GarmentDefectCodeID", DbType.String, Item.GarmentDefectCodeID);}
            if (Item.DefectPicture != null) { SbSql.Append(",DefectPicture=@DefectPicture"+ Environment.NewLine); objParameter.Add("@DefectPicture", DbType.String, Item.DefectPicture);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
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
        public int Delete(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_Inspection_Detail]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
