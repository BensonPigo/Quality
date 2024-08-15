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
            SbSql.Append("        ,CrockingTestPicture1" + Environment.NewLine);
            SbSql.Append("        ,CrockingTestPicture2" + Environment.NewLine);
            SbSql.Append("        ,CrockingTestPicture3" + Environment.NewLine);
            SbSql.Append("        ,CrockingTestPicture4" + Environment.NewLine);
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
