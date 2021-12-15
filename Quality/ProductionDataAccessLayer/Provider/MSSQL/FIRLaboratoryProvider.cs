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
    public class FIRLaboratoryProvider : SQLDAL, IFIRLaboratoryProvider
    {
        #region 底層連線
        public FIRLaboratoryProvider(string ConString) : base(ConString) { }
        public FIRLaboratoryProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Laboratory Crocking & shrinkage Test(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Laboratory Crocking & shrinkage Test
        /// </summary>
        /// <param name="Item">Laboratory Crocking & shrinkage Test成員</param>
        /// <returns>回傳Laboratory Crocking & shrinkage Test</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public IList<FIR_Laboratory> Get(FIR_Laboratory Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,InspDeadline"+ Environment.NewLine);
            SbSql.Append("        ,Crocking"+ Environment.NewLine);
            SbSql.Append("        ,Heat"+ Environment.NewLine);
            SbSql.Append("        ,Wash"+ Environment.NewLine);
            SbSql.Append("        ,CrockingDate"+ Environment.NewLine);
            SbSql.Append("        ,HeatDate"+ Environment.NewLine);
            SbSql.Append("        ,WashDate"+ Environment.NewLine);
            SbSql.Append("        ,CrockingRemark"+ Environment.NewLine);
            SbSql.Append("        ,HeatRemark"+ Environment.NewLine);
            SbSql.Append("        ,WashRemark"+ Environment.NewLine);
            SbSql.Append("        ,ReceiveSampleDate"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,nonCrocking"+ Environment.NewLine);
            SbSql.Append("        ,nonHeat"+ Environment.NewLine);
            SbSql.Append("        ,nonWash"+ Environment.NewLine);
            SbSql.Append("        ,CrockingEncode"+ Environment.NewLine);
            SbSql.Append("        ,HeatEncode"+ Environment.NewLine);
            SbSql.Append("        ,WashEncode"+ Environment.NewLine);
            SbSql.Append("        ,SkewnessOptionID"+ Environment.NewLine);
            SbSql.Append("        ,CrockingInspector"+ Environment.NewLine);
            SbSql.Append("        ,HeatInspector"+ Environment.NewLine);
            SbSql.Append("        ,WashInspector"+ Environment.NewLine);
            SbSql.Append("        ,CrockingTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,CrockingTestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,HeatTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,HeatTestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,WashTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,WashTestAfterPicture"+ Environment.NewLine);
            SbSql.Append("FROM [FIR_Laboratory] WITH(NOLOCK)" + Environment.NewLine);



            return ExecuteList<FIR_Laboratory>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Laboratory Crocking & shrinkage Test(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Laboratory Crocking & shrinkage Test
        /// </summary>
        /// <param name="Item">Laboratory Crocking & shrinkage Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Create(FIR_Laboratory Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [FIR_Laboratory]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,InspDeadline"+ Environment.NewLine);
            SbSql.Append("        ,Crocking"+ Environment.NewLine);
            SbSql.Append("        ,Heat"+ Environment.NewLine);
            SbSql.Append("        ,Wash"+ Environment.NewLine);
            SbSql.Append("        ,CrockingDate"+ Environment.NewLine);
            SbSql.Append("        ,HeatDate"+ Environment.NewLine);
            SbSql.Append("        ,WashDate"+ Environment.NewLine);
            SbSql.Append("        ,CrockingRemark"+ Environment.NewLine);
            SbSql.Append("        ,HeatRemark"+ Environment.NewLine);
            SbSql.Append("        ,WashRemark"+ Environment.NewLine);
            SbSql.Append("        ,ReceiveSampleDate"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,nonCrocking"+ Environment.NewLine);
            SbSql.Append("        ,nonHeat"+ Environment.NewLine);
            SbSql.Append("        ,nonWash"+ Environment.NewLine);
            SbSql.Append("        ,CrockingEncode"+ Environment.NewLine);
            SbSql.Append("        ,HeatEncode"+ Environment.NewLine);
            SbSql.Append("        ,WashEncode"+ Environment.NewLine);
            SbSql.Append("        ,SkewnessOptionID"+ Environment.NewLine);
            SbSql.Append("        ,CrockingInspector"+ Environment.NewLine);
            SbSql.Append("        ,HeatInspector"+ Environment.NewLine);
            SbSql.Append("        ,WashInspector"+ Environment.NewLine);
            SbSql.Append("        ,CrockingTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,CrockingTestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,HeatTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,HeatTestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,WashTestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,WashTestAfterPicture"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@SEQ1"); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);
            SbSql.Append("        ,@SEQ2"); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);
            SbSql.Append("        ,@InspDeadline"); objParameter.Add("@InspDeadline", DbType.String, Item.InspDeadline);
            SbSql.Append("        ,@Crocking"); objParameter.Add("@Crocking", DbType.String, Item.Crocking);
            SbSql.Append("        ,@Heat"); objParameter.Add("@Heat", DbType.String, Item.Heat);
            SbSql.Append("        ,@Wash"); objParameter.Add("@Wash", DbType.String, Item.Wash);
            SbSql.Append("        ,@CrockingDate"); objParameter.Add("@CrockingDate", DbType.String, Item.CrockingDate);
            SbSql.Append("        ,@HeatDate"); objParameter.Add("@HeatDate", DbType.String, Item.HeatDate);
            SbSql.Append("        ,@WashDate"); objParameter.Add("@WashDate", DbType.String, Item.WashDate);
            SbSql.Append("        ,@CrockingRemark"); objParameter.Add("@CrockingRemark", DbType.String, Item.CrockingRemark);
            SbSql.Append("        ,@HeatRemark"); objParameter.Add("@HeatRemark", DbType.String, Item.HeatRemark);
            SbSql.Append("        ,@WashRemark"); objParameter.Add("@WashRemark", DbType.String, Item.WashRemark);
            SbSql.Append("        ,@ReceiveSampleDate"); objParameter.Add("@ReceiveSampleDate", DbType.String, Item.ReceiveSampleDate);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@nonCrocking"); objParameter.Add("@nonCrocking", DbType.String, Item.nonCrocking);
            SbSql.Append("        ,@nonHeat"); objParameter.Add("@nonHeat", DbType.String, Item.nonHeat);
            SbSql.Append("        ,@nonWash"); objParameter.Add("@nonWash", DbType.String, Item.nonWash);
            SbSql.Append("        ,@CrockingEncode"); objParameter.Add("@CrockingEncode", DbType.String, Item.CrockingEncode);
            SbSql.Append("        ,@HeatEncode"); objParameter.Add("@HeatEncode", DbType.String, Item.HeatEncode);
            SbSql.Append("        ,@WashEncode"); objParameter.Add("@WashEncode", DbType.String, Item.WashEncode);
            SbSql.Append("        ,@SkewnessOptionID"); objParameter.Add("@SkewnessOptionID", DbType.String, Item.SkewnessOptionID);
            SbSql.Append("        ,@CrockingInspector"); objParameter.Add("@CrockingInspector", DbType.String, Item.CrockingInspector);
            SbSql.Append("        ,@HeatInspector"); objParameter.Add("@HeatInspector", DbType.String, Item.HeatInspector);
            SbSql.Append("        ,@WashInspector"); objParameter.Add("@WashInspector", DbType.String, Item.WashInspector);
            SbSql.Append("        ,@CrockingTestBeforePicture"); objParameter.Add("@CrockingTestBeforePicture", DbType.String, Item.CrockingTestBeforePicture);
            SbSql.Append("        ,@CrockingTestAfterPicture"); objParameter.Add("@CrockingTestAfterPicture", DbType.String, Item.CrockingTestAfterPicture);
            SbSql.Append("        ,@HeatTestBeforePicture"); objParameter.Add("@HeatTestBeforePicture", DbType.String, Item.HeatTestBeforePicture);
            SbSql.Append("        ,@HeatTestAfterPicture"); objParameter.Add("@HeatTestAfterPicture", DbType.String, Item.HeatTestAfterPicture);
            SbSql.Append("        ,@WashTestBeforePicture"); objParameter.Add("@WashTestBeforePicture", DbType.String, Item.WashTestBeforePicture);
            SbSql.Append("        ,@WashTestAfterPicture"); objParameter.Add("@WashTestAfterPicture", DbType.String, Item.WashTestAfterPicture);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Laboratory Crocking & shrinkage Test(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Laboratory Crocking & shrinkage Test
        /// </summary>
        /// <param name="Item">Laboratory Crocking & shrinkage Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Update(FIR_Laboratory Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [FIR_Laboratory]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.POID != null) { SbSql.Append(",POID=@POID"+ Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID);}
            if (Item.SEQ1 != null) { SbSql.Append(",SEQ1=@SEQ1"+ Environment.NewLine); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);}
            if (Item.SEQ2 != null) { SbSql.Append(",SEQ2=@SEQ2"+ Environment.NewLine); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);}
            if (Item.InspDeadline != null) { SbSql.Append(",InspDeadline=@InspDeadline"+ Environment.NewLine); objParameter.Add("@InspDeadline", DbType.String, Item.InspDeadline);}
            if (Item.Crocking != null) { SbSql.Append(",Crocking=@Crocking"+ Environment.NewLine); objParameter.Add("@Crocking", DbType.String, Item.Crocking);}
            if (Item.Heat != null) { SbSql.Append(",Heat=@Heat"+ Environment.NewLine); objParameter.Add("@Heat", DbType.String, Item.Heat);}
            if (Item.Wash != null) { SbSql.Append(",Wash=@Wash"+ Environment.NewLine); objParameter.Add("@Wash", DbType.String, Item.Wash);}
            if (Item.CrockingDate != null) { SbSql.Append(",CrockingDate=@CrockingDate"+ Environment.NewLine); objParameter.Add("@CrockingDate", DbType.String, Item.CrockingDate);}
            if (Item.HeatDate != null) { SbSql.Append(",HeatDate=@HeatDate"+ Environment.NewLine); objParameter.Add("@HeatDate", DbType.String, Item.HeatDate);}
            if (Item.WashDate != null) { SbSql.Append(",WashDate=@WashDate"+ Environment.NewLine); objParameter.Add("@WashDate", DbType.String, Item.WashDate);}
            if (Item.CrockingRemark != null) { SbSql.Append(",CrockingRemark=@CrockingRemark"+ Environment.NewLine); objParameter.Add("@CrockingRemark", DbType.String, Item.CrockingRemark);}
            if (Item.HeatRemark != null) { SbSql.Append(",HeatRemark=@HeatRemark"+ Environment.NewLine); objParameter.Add("@HeatRemark", DbType.String, Item.HeatRemark);}
            if (Item.WashRemark != null) { SbSql.Append(",WashRemark=@WashRemark"+ Environment.NewLine); objParameter.Add("@WashRemark", DbType.String, Item.WashRemark);}
            if (Item.ReceiveSampleDate != null) { SbSql.Append(",ReceiveSampleDate=@ReceiveSampleDate"+ Environment.NewLine); objParameter.Add("@ReceiveSampleDate", DbType.String, Item.ReceiveSampleDate);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.nonCrocking != null) { SbSql.Append(",nonCrocking=@nonCrocking"+ Environment.NewLine); objParameter.Add("@nonCrocking", DbType.String, Item.nonCrocking);}
            if (Item.nonHeat != null) { SbSql.Append(",nonHeat=@nonHeat"+ Environment.NewLine); objParameter.Add("@nonHeat", DbType.String, Item.nonHeat);}
            if (Item.nonWash != null) { SbSql.Append(",nonWash=@nonWash"+ Environment.NewLine); objParameter.Add("@nonWash", DbType.String, Item.nonWash);}
            if (Item.CrockingEncode != null) { SbSql.Append(",CrockingEncode=@CrockingEncode"+ Environment.NewLine); objParameter.Add("@CrockingEncode", DbType.String, Item.CrockingEncode);}
            if (Item.HeatEncode != null) { SbSql.Append(",HeatEncode=@HeatEncode"+ Environment.NewLine); objParameter.Add("@HeatEncode", DbType.String, Item.HeatEncode);}
            if (Item.WashEncode != null) { SbSql.Append(",WashEncode=@WashEncode"+ Environment.NewLine); objParameter.Add("@WashEncode", DbType.String, Item.WashEncode);}
            if (Item.SkewnessOptionID != null) { SbSql.Append(",SkewnessOptionID=@SkewnessOptionID"+ Environment.NewLine); objParameter.Add("@SkewnessOptionID", DbType.String, Item.SkewnessOptionID);}
            if (Item.CrockingInspector != null) { SbSql.Append(",CrockingInspector=@CrockingInspector"+ Environment.NewLine); objParameter.Add("@CrockingInspector", DbType.String, Item.CrockingInspector);}
            if (Item.HeatInspector != null) { SbSql.Append(",HeatInspector=@HeatInspector"+ Environment.NewLine); objParameter.Add("@HeatInspector", DbType.String, Item.HeatInspector);}
            if (Item.WashInspector != null) { SbSql.Append(",WashInspector=@WashInspector"+ Environment.NewLine); objParameter.Add("@WashInspector", DbType.String, Item.WashInspector);}
            if (Item.CrockingTestBeforePicture != null) { SbSql.Append(",CrockingTestBeforePicture=@CrockingTestBeforePicture"+ Environment.NewLine); objParameter.Add("@CrockingTestBeforePicture", DbType.String, Item.CrockingTestBeforePicture);}
            if (Item.CrockingTestAfterPicture != null) { SbSql.Append(",CrockingTestAfterPicture=@CrockingTestAfterPicture"+ Environment.NewLine); objParameter.Add("@CrockingTestAfterPicture", DbType.String, Item.CrockingTestAfterPicture);}
            if (Item.HeatTestBeforePicture != null) { SbSql.Append(",HeatTestBeforePicture=@HeatTestBeforePicture"+ Environment.NewLine); objParameter.Add("@HeatTestBeforePicture", DbType.String, Item.HeatTestBeforePicture);}
            if (Item.HeatTestAfterPicture != null) { SbSql.Append(",HeatTestAfterPicture=@HeatTestAfterPicture"+ Environment.NewLine); objParameter.Add("@HeatTestAfterPicture", DbType.String, Item.HeatTestAfterPicture);}
            if (Item.WashTestBeforePicture != null) { SbSql.Append(",WashTestBeforePicture=@WashTestBeforePicture"+ Environment.NewLine); objParameter.Add("@WashTestBeforePicture", DbType.String, Item.WashTestBeforePicture);}
            if (Item.WashTestAfterPicture != null) { SbSql.Append(",WashTestAfterPicture=@WashTestAfterPicture"+ Environment.NewLine); objParameter.Add("@WashTestAfterPicture", DbType.String, Item.WashTestAfterPicture);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Laboratory Crocking & shrinkage Test(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Laboratory Crocking & shrinkage Test
        /// </summary>
        /// <param name="Item">Laboratory Crocking & shrinkage Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Delete(FIR_Laboratory Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [FIR_Laboratory]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
