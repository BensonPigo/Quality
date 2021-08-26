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
    public class FIRProvider : SQLDAL, IFIRProvider
    {
        #region 底層連線
        public FIRProvider(string ConString) : base(ConString) { }
        public FIRProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Fabric Inspection Report(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Fabric Inspection Report
        /// </summary>
        /// <param name="Item">Fabric Inspection Report成員</param>
        /// <returns>回傳Fabric Inspection Report</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public IList<FIR> Get(FIR Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,Suppid"+ Environment.NewLine);
            SbSql.Append("        ,SCIRefno"+ Environment.NewLine);
            SbSql.Append("        ,Refno"+ Environment.NewLine);
            SbSql.Append("        ,ReceivingID"+ Environment.NewLine);
            SbSql.Append("        ,ReplacementReportID"+ Environment.NewLine);
            SbSql.Append("        ,ArriveQty"+ Environment.NewLine);
            SbSql.Append("        ,TotalInspYds"+ Environment.NewLine);
            SbSql.Append("        ,TotalDefectPoint"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Nonphysical"+ Environment.NewLine);
            SbSql.Append("        ,Physical"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalEncode"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalDate"+ Environment.NewLine);
            SbSql.Append("        ,nonWeight"+ Environment.NewLine);
            SbSql.Append("        ,Weight"+ Environment.NewLine);
            SbSql.Append("        ,WeightEncode"+ Environment.NewLine);
            SbSql.Append("        ,WeightDate"+ Environment.NewLine);
            SbSql.Append("        ,nonShadebond"+ Environment.NewLine);
            SbSql.Append("        ,ShadeBond"+ Environment.NewLine);
            SbSql.Append("        ,ShadebondEncode"+ Environment.NewLine);
            SbSql.Append("        ,ShadeBondDate"+ Environment.NewLine);
            SbSql.Append("        ,nonContinuity"+ Environment.NewLine);
            SbSql.Append("        ,Continuity"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityEncode"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityDate"+ Environment.NewLine);
            SbSql.Append("        ,InspDeadline"+ Environment.NewLine);
            SbSql.Append("        ,Approve"+ Environment.NewLine);
            SbSql.Append("        ,ApproveDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,OldFabricUkey"+ Environment.NewLine);
            SbSql.Append("        ,OldFabricVer"+ Environment.NewLine);
            SbSql.Append("        ,nonOdor"+ Environment.NewLine);
            SbSql.Append("        ,Odor"+ Environment.NewLine);
            SbSql.Append("        ,OdorEncode"+ Environment.NewLine);
            SbSql.Append("        ,OdorDate"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalInspector"+ Environment.NewLine);
            SbSql.Append("        ,WeightInspector"+ Environment.NewLine);
            SbSql.Append("        ,ShadeboneInspector"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityInspector"+ Environment.NewLine);
            SbSql.Append("        ,OdorInspector"+ Environment.NewLine);
            SbSql.Append("        ,Moisture"+ Environment.NewLine);
            SbSql.Append("        ,NonMoisture"+ Environment.NewLine);
            SbSql.Append("        ,MoistureDate"+ Environment.NewLine);
            SbSql.Append("        ,MaterialCompositionGrp"+ Environment.NewLine);
            SbSql.Append("        ,MaterialCompositionItem"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandardDesc"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard1"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard1_Comparison"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard2"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard2_Comparison"+ Environment.NewLine);
            SbSql.Append("FROM [FIR]"+ Environment.NewLine);



            return ExecuteList<FIR>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Fabric Inspection Report(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Fabric Inspection Report
        /// </summary>
        /// <param name="Item">Fabric Inspection Report成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Create(FIR Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [FIR]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,Suppid"+ Environment.NewLine);
            SbSql.Append("        ,SCIRefno"+ Environment.NewLine);
            SbSql.Append("        ,Refno"+ Environment.NewLine);
            SbSql.Append("        ,ReceivingID"+ Environment.NewLine);
            SbSql.Append("        ,ReplacementReportID"+ Environment.NewLine);
            SbSql.Append("        ,ArriveQty"+ Environment.NewLine);
            SbSql.Append("        ,TotalInspYds"+ Environment.NewLine);
            SbSql.Append("        ,TotalDefectPoint"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Nonphysical"+ Environment.NewLine);
            SbSql.Append("        ,Physical"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalEncode"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalDate"+ Environment.NewLine);
            SbSql.Append("        ,nonWeight"+ Environment.NewLine);
            SbSql.Append("        ,Weight"+ Environment.NewLine);
            SbSql.Append("        ,WeightEncode"+ Environment.NewLine);
            SbSql.Append("        ,WeightDate"+ Environment.NewLine);
            SbSql.Append("        ,nonShadebond"+ Environment.NewLine);
            SbSql.Append("        ,ShadeBond"+ Environment.NewLine);
            SbSql.Append("        ,ShadebondEncode"+ Environment.NewLine);
            SbSql.Append("        ,ShadeBondDate"+ Environment.NewLine);
            SbSql.Append("        ,nonContinuity"+ Environment.NewLine);
            SbSql.Append("        ,Continuity"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityEncode"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityDate"+ Environment.NewLine);
            SbSql.Append("        ,InspDeadline"+ Environment.NewLine);
            SbSql.Append("        ,Approve"+ Environment.NewLine);
            SbSql.Append("        ,ApproveDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,OldFabricUkey"+ Environment.NewLine);
            SbSql.Append("        ,OldFabricVer"+ Environment.NewLine);
            SbSql.Append("        ,nonOdor"+ Environment.NewLine);
            SbSql.Append("        ,Odor"+ Environment.NewLine);
            SbSql.Append("        ,OdorEncode"+ Environment.NewLine);
            SbSql.Append("        ,OdorDate"+ Environment.NewLine);
            SbSql.Append("        ,PhysicalInspector"+ Environment.NewLine);
            SbSql.Append("        ,WeightInspector"+ Environment.NewLine);
            SbSql.Append("        ,ShadeboneInspector"+ Environment.NewLine);
            SbSql.Append("        ,ContinuityInspector"+ Environment.NewLine);
            SbSql.Append("        ,OdorInspector"+ Environment.NewLine);
            SbSql.Append("        ,Moisture"+ Environment.NewLine);
            SbSql.Append("        ,NonMoisture"+ Environment.NewLine);
            SbSql.Append("        ,MoistureDate"+ Environment.NewLine);
            SbSql.Append("        ,MaterialCompositionGrp"+ Environment.NewLine);
            SbSql.Append("        ,MaterialCompositionItem"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandardDesc"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard1"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard1_Comparison"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard2"+ Environment.NewLine);
            SbSql.Append("        ,MoistureStandard2_Comparison"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@SEQ1"); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);
            SbSql.Append("        ,@SEQ2"); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);
            SbSql.Append("        ,@Suppid"); objParameter.Add("@Suppid", DbType.String, Item.Suppid);
            SbSql.Append("        ,@SCIRefno"); objParameter.Add("@SCIRefno", DbType.String, Item.SCIRefno);
            SbSql.Append("        ,@Refno"); objParameter.Add("@Refno", DbType.String, Item.Refno);
            SbSql.Append("        ,@ReceivingID"); objParameter.Add("@ReceivingID", DbType.String, Item.ReceivingID);
            SbSql.Append("        ,@ReplacementReportID"); objParameter.Add("@ReplacementReportID", DbType.String, Item.ReplacementReportID);
            SbSql.Append("        ,@ArriveQty"); objParameter.Add("@ArriveQty", DbType.String, Item.ArriveQty);
            SbSql.Append("        ,@TotalInspYds"); objParameter.Add("@TotalInspYds", DbType.String, Item.TotalInspYds);
            SbSql.Append("        ,@TotalDefectPoint"); objParameter.Add("@TotalDefectPoint", DbType.String, Item.TotalDefectPoint);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@Nonphysical"); objParameter.Add("@Nonphysical", DbType.String, Item.Nonphysical);
            SbSql.Append("        ,@Physical"); objParameter.Add("@Physical", DbType.String, Item.Physical);
            SbSql.Append("        ,@PhysicalEncode"); objParameter.Add("@PhysicalEncode", DbType.String, Item.PhysicalEncode);
            SbSql.Append("        ,@PhysicalDate"); objParameter.Add("@PhysicalDate", DbType.DateTime, Item.PhysicalDate);
            SbSql.Append("        ,@nonWeight"); objParameter.Add("@nonWeight", DbType.String, Item.nonWeight);
            SbSql.Append("        ,@Weight"); objParameter.Add("@Weight", DbType.String, Item.Weight);
            SbSql.Append("        ,@WeightEncode"); objParameter.Add("@WeightEncode", DbType.String, Item.WeightEncode);
            SbSql.Append("        ,@WeightDate"); objParameter.Add("@WeightDate", DbType.DateTime, Item.WeightDate);
            SbSql.Append("        ,@nonShadebond"); objParameter.Add("@nonShadebond", DbType.String, Item.nonShadebond);
            SbSql.Append("        ,@ShadeBond"); objParameter.Add("@ShadeBond", DbType.String, Item.ShadeBond);
            SbSql.Append("        ,@ShadebondEncode"); objParameter.Add("@ShadebondEncode", DbType.String, Item.ShadebondEncode);
            SbSql.Append("        ,@ShadeBondDate"); objParameter.Add("@ShadeBondDate", DbType.DateTime, Item.ShadeBondDate);
            SbSql.Append("        ,@nonContinuity"); objParameter.Add("@nonContinuity", DbType.String, Item.nonContinuity);
            SbSql.Append("        ,@Continuity"); objParameter.Add("@Continuity", DbType.String, Item.Continuity);
            SbSql.Append("        ,@ContinuityEncode"); objParameter.Add("@ContinuityEncode", DbType.String, Item.ContinuityEncode);
            SbSql.Append("        ,@ContinuityDate"); objParameter.Add("@ContinuityDate", DbType.DateTime, Item.ContinuityDate);
            SbSql.Append("        ,@InspDeadline"); objParameter.Add("@InspDeadline", DbType.String, Item.InspDeadline);
            SbSql.Append("        ,@Approve"); objParameter.Add("@Approve", DbType.String, Item.Approve);
            SbSql.Append("        ,@ApproveDate"); objParameter.Add("@ApproveDate", DbType.DateTime, Item.ApproveDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
            SbSql.Append("        ,@OldFabricUkey"); objParameter.Add("@OldFabricUkey", DbType.String, Item.OldFabricUkey);
            SbSql.Append("        ,@OldFabricVer"); objParameter.Add("@OldFabricVer", DbType.String, Item.OldFabricVer);
            SbSql.Append("        ,@nonOdor"); objParameter.Add("@nonOdor", DbType.String, Item.nonOdor);
            SbSql.Append("        ,@Odor"); objParameter.Add("@Odor", DbType.String, Item.Odor);
            SbSql.Append("        ,@OdorEncode"); objParameter.Add("@OdorEncode", DbType.String, Item.OdorEncode);
            SbSql.Append("        ,@OdorDate"); objParameter.Add("@OdorDate", DbType.DateTime, Item.OdorDate);
            SbSql.Append("        ,@PhysicalInspector"); objParameter.Add("@PhysicalInspector", DbType.String, Item.PhysicalInspector);
            SbSql.Append("        ,@WeightInspector"); objParameter.Add("@WeightInspector", DbType.String, Item.WeightInspector);
            SbSql.Append("        ,@ShadeboneInspector"); objParameter.Add("@ShadeboneInspector", DbType.String, Item.ShadeboneInspector);
            SbSql.Append("        ,@ContinuityInspector"); objParameter.Add("@ContinuityInspector", DbType.String, Item.ContinuityInspector);
            SbSql.Append("        ,@OdorInspector"); objParameter.Add("@OdorInspector", DbType.String, Item.OdorInspector);
            SbSql.Append("        ,@Moisture"); objParameter.Add("@Moisture", DbType.String, Item.Moisture);
            SbSql.Append("        ,@NonMoisture"); objParameter.Add("@NonMoisture", DbType.String, Item.NonMoisture);
            SbSql.Append("        ,@MoistureDate"); objParameter.Add("@MoistureDate", DbType.String, Item.MoistureDate);
            SbSql.Append("        ,@MaterialCompositionGrp"); objParameter.Add("@MaterialCompositionGrp", DbType.String, Item.MaterialCompositionGrp);
            SbSql.Append("        ,@MaterialCompositionItem"); objParameter.Add("@MaterialCompositionItem", DbType.String, Item.MaterialCompositionItem);
            SbSql.Append("        ,@MoistureStandardDesc"); objParameter.Add("@MoistureStandardDesc", DbType.String, Item.MoistureStandardDesc);
            SbSql.Append("        ,@MoistureStandard1"); objParameter.Add("@MoistureStandard1", DbType.String, Item.MoistureStandard1);
            SbSql.Append("        ,@MoistureStandard1_Comparison"); objParameter.Add("@MoistureStandard1_Comparison", DbType.String, Item.MoistureStandard1_Comparison);
            SbSql.Append("        ,@MoistureStandard2"); objParameter.Add("@MoistureStandard2", DbType.String, Item.MoistureStandard2);
            SbSql.Append("        ,@MoistureStandard2_Comparison"); objParameter.Add("@MoistureStandard2_Comparison", DbType.String, Item.MoistureStandard2_Comparison);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Fabric Inspection Report(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Fabric Inspection Report
        /// </summary>
        /// <param name="Item">Fabric Inspection Report成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Update(FIR Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [FIR]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.POID != null) { SbSql.Append(",POID=@POID"+ Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID);}
            if (Item.SEQ1 != null) { SbSql.Append(",SEQ1=@SEQ1"+ Environment.NewLine); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);}
            if (Item.SEQ2 != null) { SbSql.Append(",SEQ2=@SEQ2"+ Environment.NewLine); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);}
            if (Item.Suppid != null) { SbSql.Append(",Suppid=@Suppid"+ Environment.NewLine); objParameter.Add("@Suppid", DbType.String, Item.Suppid);}
            if (Item.SCIRefno != null) { SbSql.Append(",SCIRefno=@SCIRefno"+ Environment.NewLine); objParameter.Add("@SCIRefno", DbType.String, Item.SCIRefno);}
            if (Item.Refno != null) { SbSql.Append(",Refno=@Refno"+ Environment.NewLine); objParameter.Add("@Refno", DbType.String, Item.Refno);}
            if (Item.ReceivingID != null) { SbSql.Append(",ReceivingID=@ReceivingID"+ Environment.NewLine); objParameter.Add("@ReceivingID", DbType.String, Item.ReceivingID);}
            if (Item.ReplacementReportID != null) { SbSql.Append(",ReplacementReportID=@ReplacementReportID"+ Environment.NewLine); objParameter.Add("@ReplacementReportID", DbType.String, Item.ReplacementReportID);}
            if (Item.ArriveQty != null) { SbSql.Append(",ArriveQty=@ArriveQty"+ Environment.NewLine); objParameter.Add("@ArriveQty", DbType.String, Item.ArriveQty);}
            if (Item.TotalInspYds != null) { SbSql.Append(",TotalInspYds=@TotalInspYds"+ Environment.NewLine); objParameter.Add("@TotalInspYds", DbType.String, Item.TotalInspYds);}
            if (Item.TotalDefectPoint != null) { SbSql.Append(",TotalDefectPoint=@TotalDefectPoint"+ Environment.NewLine); objParameter.Add("@TotalDefectPoint", DbType.String, Item.TotalDefectPoint);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.Nonphysical != null) { SbSql.Append(",Nonphysical=@Nonphysical"+ Environment.NewLine); objParameter.Add("@Nonphysical", DbType.String, Item.Nonphysical);}
            if (Item.Physical != null) { SbSql.Append(",Physical=@Physical"+ Environment.NewLine); objParameter.Add("@Physical", DbType.String, Item.Physical);}
            if (Item.PhysicalEncode != null) { SbSql.Append(",PhysicalEncode=@PhysicalEncode"+ Environment.NewLine); objParameter.Add("@PhysicalEncode", DbType.String, Item.PhysicalEncode);}
            if (Item.PhysicalDate != null) { SbSql.Append(",PhysicalDate=@PhysicalDate"+ Environment.NewLine); objParameter.Add("@PhysicalDate", DbType.DateTime, Item.PhysicalDate);}
            if (Item.nonWeight != null) { SbSql.Append(",nonWeight=@nonWeight"+ Environment.NewLine); objParameter.Add("@nonWeight", DbType.String, Item.nonWeight);}
            if (Item.Weight != null) { SbSql.Append(",Weight=@Weight"+ Environment.NewLine); objParameter.Add("@Weight", DbType.String, Item.Weight);}
            if (Item.WeightEncode != null) { SbSql.Append(",WeightEncode=@WeightEncode"+ Environment.NewLine); objParameter.Add("@WeightEncode", DbType.String, Item.WeightEncode);}
            if (Item.WeightDate != null) { SbSql.Append(",WeightDate=@WeightDate"+ Environment.NewLine); objParameter.Add("@WeightDate", DbType.DateTime, Item.WeightDate);}
            if (Item.nonShadebond != null) { SbSql.Append(",nonShadebond=@nonShadebond"+ Environment.NewLine); objParameter.Add("@nonShadebond", DbType.String, Item.nonShadebond);}
            if (Item.ShadeBond != null) { SbSql.Append(",ShadeBond=@ShadeBond"+ Environment.NewLine); objParameter.Add("@ShadeBond", DbType.String, Item.ShadeBond);}
            if (Item.ShadebondEncode != null) { SbSql.Append(",ShadebondEncode=@ShadebondEncode"+ Environment.NewLine); objParameter.Add("@ShadebondEncode", DbType.String, Item.ShadebondEncode);}
            if (Item.ShadeBondDate != null) { SbSql.Append(",ShadeBondDate=@ShadeBondDate"+ Environment.NewLine); objParameter.Add("@ShadeBondDate", DbType.DateTime, Item.ShadeBondDate);}
            if (Item.nonContinuity != null) { SbSql.Append(",nonContinuity=@nonContinuity"+ Environment.NewLine); objParameter.Add("@nonContinuity", DbType.String, Item.nonContinuity);}
            if (Item.Continuity != null) { SbSql.Append(",Continuity=@Continuity"+ Environment.NewLine); objParameter.Add("@Continuity", DbType.String, Item.Continuity);}
            if (Item.ContinuityEncode != null) { SbSql.Append(",ContinuityEncode=@ContinuityEncode"+ Environment.NewLine); objParameter.Add("@ContinuityEncode", DbType.String, Item.ContinuityEncode);}
            if (Item.ContinuityDate != null) { SbSql.Append(",ContinuityDate=@ContinuityDate"+ Environment.NewLine); objParameter.Add("@ContinuityDate", DbType.DateTime, Item.ContinuityDate);}
            if (Item.InspDeadline != null) { SbSql.Append(",InspDeadline=@InspDeadline"+ Environment.NewLine); objParameter.Add("@InspDeadline", DbType.String, Item.InspDeadline);}
            if (Item.Approve != null) { SbSql.Append(",Approve=@Approve"+ Environment.NewLine); objParameter.Add("@Approve", DbType.String, Item.Approve);}
            if (Item.ApproveDate != null) { SbSql.Append(",ApproveDate=@ApproveDate"+ Environment.NewLine); objParameter.Add("@ApproveDate", DbType.DateTime, Item.ApproveDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine); objParameter.Add("@Status", DbType.String, Item.Status);}
            if (Item.OldFabricUkey != null) { SbSql.Append(",OldFabricUkey=@OldFabricUkey"+ Environment.NewLine); objParameter.Add("@OldFabricUkey", DbType.String, Item.OldFabricUkey);}
            if (Item.OldFabricVer != null) { SbSql.Append(",OldFabricVer=@OldFabricVer"+ Environment.NewLine); objParameter.Add("@OldFabricVer", DbType.String, Item.OldFabricVer);}
            if (Item.nonOdor != null) { SbSql.Append(",nonOdor=@nonOdor"+ Environment.NewLine); objParameter.Add("@nonOdor", DbType.String, Item.nonOdor);}
            if (Item.Odor != null) { SbSql.Append(",Odor=@Odor"+ Environment.NewLine); objParameter.Add("@Odor", DbType.String, Item.Odor);}
            if (Item.OdorEncode != null) { SbSql.Append(",OdorEncode=@OdorEncode"+ Environment.NewLine); objParameter.Add("@OdorEncode", DbType.String, Item.OdorEncode);}
            if (Item.OdorDate != null) { SbSql.Append(",OdorDate=@OdorDate"+ Environment.NewLine); objParameter.Add("@OdorDate", DbType.DateTime, Item.OdorDate);}
            if (Item.PhysicalInspector != null) { SbSql.Append(",PhysicalInspector=@PhysicalInspector"+ Environment.NewLine); objParameter.Add("@PhysicalInspector", DbType.String, Item.PhysicalInspector);}
            if (Item.WeightInspector != null) { SbSql.Append(",WeightInspector=@WeightInspector"+ Environment.NewLine); objParameter.Add("@WeightInspector", DbType.String, Item.WeightInspector);}
            if (Item.ShadeboneInspector != null) { SbSql.Append(",ShadeboneInspector=@ShadeboneInspector"+ Environment.NewLine); objParameter.Add("@ShadeboneInspector", DbType.String, Item.ShadeboneInspector);}
            if (Item.ContinuityInspector != null) { SbSql.Append(",ContinuityInspector=@ContinuityInspector"+ Environment.NewLine); objParameter.Add("@ContinuityInspector", DbType.String, Item.ContinuityInspector);}
            if (Item.OdorInspector != null) { SbSql.Append(",OdorInspector=@OdorInspector"+ Environment.NewLine); objParameter.Add("@OdorInspector", DbType.String, Item.OdorInspector);}
            if (Item.Moisture != null) { SbSql.Append(",Moisture=@Moisture"+ Environment.NewLine); objParameter.Add("@Moisture", DbType.String, Item.Moisture);}
            if (Item.NonMoisture != null) { SbSql.Append(",NonMoisture=@NonMoisture"+ Environment.NewLine); objParameter.Add("@NonMoisture", DbType.String, Item.NonMoisture);}
            if (Item.MoistureDate != null) { SbSql.Append(",MoistureDate=@MoistureDate"+ Environment.NewLine); objParameter.Add("@MoistureDate", DbType.String, Item.MoistureDate);}
            if (Item.MaterialCompositionGrp != null) { SbSql.Append(",MaterialCompositionGrp=@MaterialCompositionGrp"+ Environment.NewLine); objParameter.Add("@MaterialCompositionGrp", DbType.String, Item.MaterialCompositionGrp);}
            if (Item.MaterialCompositionItem != null) { SbSql.Append(",MaterialCompositionItem=@MaterialCompositionItem"+ Environment.NewLine); objParameter.Add("@MaterialCompositionItem", DbType.String, Item.MaterialCompositionItem);}
            if (Item.MoistureStandardDesc != null) { SbSql.Append(",MoistureStandardDesc=@MoistureStandardDesc"+ Environment.NewLine); objParameter.Add("@MoistureStandardDesc", DbType.String, Item.MoistureStandardDesc);}
            if (Item.MoistureStandard1 != null) { SbSql.Append(",MoistureStandard1=@MoistureStandard1"+ Environment.NewLine); objParameter.Add("@MoistureStandard1", DbType.String, Item.MoistureStandard1);}
            if (Item.MoistureStandard1_Comparison != null) { SbSql.Append(",MoistureStandard1_Comparison=@MoistureStandard1_Comparison"+ Environment.NewLine); objParameter.Add("@MoistureStandard1_Comparison", DbType.String, Item.MoistureStandard1_Comparison);}
            if (Item.MoistureStandard2 != null) { SbSql.Append(",MoistureStandard2=@MoistureStandard2"+ Environment.NewLine); objParameter.Add("@MoistureStandard2", DbType.String, Item.MoistureStandard2);}
            if (Item.MoistureStandard2_Comparison != null) { SbSql.Append(",MoistureStandard2_Comparison=@MoistureStandard2_Comparison"+ Environment.NewLine); objParameter.Add("@MoistureStandard2_Comparison", DbType.String, Item.MoistureStandard2_Comparison);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Fabric Inspection Report(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Fabric Inspection Report
        /// </summary>
        /// <param name="Item">Fabric Inspection Report成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/25  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/25  1.00    Admin        Create
        /// </history>
        public int Delete(FIR Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [FIR]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
