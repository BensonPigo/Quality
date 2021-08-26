using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class StyleProvider : SQLDAL, IStyleProvider
    {
        #region 底層連線
        public StyleProvider(string ConString) : base(ConString) { }
        public StyleProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<Style> GetSizeUnit(Int64 ukey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                 { "@ukey", DbType.Int64, ukey } ,
            };

            string sqlcmd = @"
select SizeUnit from Style 
where 1=1
and ukey = @ukey
";

            return ExecuteList<Style>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<Style> GetStyleID()
        {
            string sqlcmd = @"select distinct id from Style where Junk = 0";

            return ExecuteList<Style>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }


        /*回傳款式資料基本檔(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳款式資料基本檔
        /// </summary>
        /// <param name="Item">款式資料基本檔成員</param>
        /// <returns>回傳款式資料基本檔</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<Style> Get(Style Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                 { "@ukey", DbType.Int64, Item.Ukey } ,
            };
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,ProgramID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,Model"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,StyleName"+ Environment.NewLine);
            SbSql.Append("        ,ComboType"+ Environment.NewLine);
            SbSql.Append("        ,CdCodeID"+ Environment.NewLine);
            SbSql.Append("        ,ApparelType"+ Environment.NewLine);
            SbSql.Append("        ,FabricType"+ Environment.NewLine);
            SbSql.Append("        ,Contents"+ Environment.NewLine);
            SbSql.Append("        ,GMTLT"+ Environment.NewLine);
            SbSql.Append("        ,CPU"+ Environment.NewLine);
            SbSql.Append("        ,Factories"+ Environment.NewLine);
            SbSql.Append("        ,FTYRemark"+ Environment.NewLine);
            SbSql.Append("        ,SampleSMR"+ Environment.NewLine);
            SbSql.Append("        ,SampleMRHandle"+ Environment.NewLine);
            SbSql.Append("        ,BulkSMR"+ Environment.NewLine);
            SbSql.Append("        ,BulkMRHandle"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,RainwearTestPassed"+ Environment.NewLine);
            SbSql.Append("        ,SizePage"+ Environment.NewLine);
            SbSql.Append("        ,SizeRange"+ Environment.NewLine);
            SbSql.Append("        ,CTNQty"+ Environment.NewLine);
            SbSql.Append("        ,StdCost"+ Environment.NewLine);
            SbSql.Append("        ,Processes"+ Environment.NewLine);
            SbSql.Append("        ,ArtworkCost"+ Environment.NewLine);
            SbSql.Append("        ,Picture1"+ Environment.NewLine);
            SbSql.Append("        ,Picture2"+ Environment.NewLine);
            SbSql.Append("        ,Label"+ Environment.NewLine);
            SbSql.Append("        ,Packing"+ Environment.NewLine);
            SbSql.Append("        ,IETMSID"+ Environment.NewLine);
            SbSql.Append("        ,IETMSVersion"+ Environment.NewLine);
            SbSql.Append("        ,IEImportName"+ Environment.NewLine);
            SbSql.Append("        ,IEImportDate"+ Environment.NewLine);
            SbSql.Append("        ,ApvDate"+ Environment.NewLine);
            SbSql.Append("        ,ApvName"+ Environment.NewLine);
            SbSql.Append("        ,CareCode"+ Environment.NewLine);
            SbSql.Append("        ,SpecialMark"+ Environment.NewLine);
            SbSql.Append("        ,Lining"+ Environment.NewLine);
            SbSql.Append("        ,StyleUnit"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionForm"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormRemark"+ Environment.NewLine);
            SbSql.Append("        ,LocalMR"+ Environment.NewLine);
            SbSql.Append("        ,LocalStyle"+ Environment.NewLine);
            SbSql.Append("        ,PPMeeting"+ Environment.NewLine);
            SbSql.Append("        ,NoNeedPPMeeting"+ Environment.NewLine);
            SbSql.Append("        ,SampleApv"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,SizeUnit"+ Environment.NewLine);
            SbSql.Append("        ,ModularParent"+ Environment.NewLine);
            SbSql.Append("        ,CPUAdjusted"+ Environment.NewLine);
            SbSql.Append("        ,Phase"+ Environment.NewLine);
            SbSql.Append("        ,Gender"+ Environment.NewLine);
            SbSql.Append("        ,ThreadEditname"+ Environment.NewLine);
            SbSql.Append("        ,ThreadEditdate"+ Environment.NewLine);
            SbSql.Append("        ,ThickFabric"+ Environment.NewLine);
            SbSql.Append("        ,DyeingID"+ Environment.NewLine);
            SbSql.Append("        ,TPEEditName"+ Environment.NewLine);
            SbSql.Append("        ,TPEEditDate"+ Environment.NewLine);
            SbSql.Append("        ,Pressing1"+ Environment.NewLine);
            SbSql.Append("        ,Pressing2"+ Environment.NewLine);
            SbSql.Append("        ,Folding1"+ Environment.NewLine);
            SbSql.Append("        ,Folding2"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormStatus"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormDate"+ Environment.NewLine);
            SbSql.Append("        ,ThickFabricBulk"+ Environment.NewLine);
            SbSql.Append("        ,HangerPack"+ Environment.NewLine);
            SbSql.Append("        ,Construction"+ Environment.NewLine);
            SbSql.Append("        ,CDCodeNew"+ Environment.NewLine);
            SbSql.Append("        ,FitType"+ Environment.NewLine);
            SbSql.Append("        ,GearLine"+ Environment.NewLine);
            SbSql.Append("        ,ThreadVersion"+ Environment.NewLine);
            SbSql.Append("FROM [Style]"+ Environment.NewLine);
            SbSql.Append("where 1 = 1" + Environment.NewLine);
            if (Item.Ukey != 0)
            {
                SbSql.Append("and ukey = @Ukey" + Environment.NewLine);
            }


            return ExecuteList<Style>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立款式資料基本檔(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立款式資料基本檔
        /// </summary>
        /// <param name="Item">款式資料基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(Style Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Style]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,ProgramID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,Model"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,StyleName"+ Environment.NewLine);
            SbSql.Append("        ,ComboType"+ Environment.NewLine);
            SbSql.Append("        ,CdCodeID"+ Environment.NewLine);
            SbSql.Append("        ,ApparelType"+ Environment.NewLine);
            SbSql.Append("        ,FabricType"+ Environment.NewLine);
            SbSql.Append("        ,Contents"+ Environment.NewLine);
            SbSql.Append("        ,GMTLT"+ Environment.NewLine);
            SbSql.Append("        ,CPU"+ Environment.NewLine);
            SbSql.Append("        ,Factories"+ Environment.NewLine);
            SbSql.Append("        ,FTYRemark"+ Environment.NewLine);
            SbSql.Append("        ,SampleSMR"+ Environment.NewLine);
            SbSql.Append("        ,SampleMRHandle"+ Environment.NewLine);
            SbSql.Append("        ,BulkSMR"+ Environment.NewLine);
            SbSql.Append("        ,BulkMRHandle"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,RainwearTestPassed"+ Environment.NewLine);
            SbSql.Append("        ,SizePage"+ Environment.NewLine);
            SbSql.Append("        ,SizeRange"+ Environment.NewLine);
            SbSql.Append("        ,CTNQty"+ Environment.NewLine);
            SbSql.Append("        ,StdCost"+ Environment.NewLine);
            SbSql.Append("        ,Processes"+ Environment.NewLine);
            SbSql.Append("        ,ArtworkCost"+ Environment.NewLine);
            SbSql.Append("        ,Picture1"+ Environment.NewLine);
            SbSql.Append("        ,Picture2"+ Environment.NewLine);
            SbSql.Append("        ,Label"+ Environment.NewLine);
            SbSql.Append("        ,Packing"+ Environment.NewLine);
            SbSql.Append("        ,IETMSID"+ Environment.NewLine);
            SbSql.Append("        ,IETMSVersion"+ Environment.NewLine);
            SbSql.Append("        ,IEImportName"+ Environment.NewLine);
            SbSql.Append("        ,IEImportDate"+ Environment.NewLine);
            SbSql.Append("        ,ApvDate"+ Environment.NewLine);
            SbSql.Append("        ,ApvName"+ Environment.NewLine);
            SbSql.Append("        ,CareCode"+ Environment.NewLine);
            SbSql.Append("        ,SpecialMark"+ Environment.NewLine);
            SbSql.Append("        ,Lining"+ Environment.NewLine);
            SbSql.Append("        ,StyleUnit"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionForm"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormRemark"+ Environment.NewLine);
            SbSql.Append("        ,LocalMR"+ Environment.NewLine);
            SbSql.Append("        ,LocalStyle"+ Environment.NewLine);
            SbSql.Append("        ,PPMeeting"+ Environment.NewLine);
            SbSql.Append("        ,NoNeedPPMeeting"+ Environment.NewLine);
            SbSql.Append("        ,SampleApv"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,SizeUnit"+ Environment.NewLine);
            SbSql.Append("        ,ModularParent"+ Environment.NewLine);
            SbSql.Append("        ,CPUAdjusted"+ Environment.NewLine);
            SbSql.Append("        ,Phase"+ Environment.NewLine);
            SbSql.Append("        ,Gender"+ Environment.NewLine);
            SbSql.Append("        ,ThreadEditname"+ Environment.NewLine);
            SbSql.Append("        ,ThreadEditdate"+ Environment.NewLine);
            SbSql.Append("        ,ThickFabric"+ Environment.NewLine);
            SbSql.Append("        ,DyeingID"+ Environment.NewLine);
            SbSql.Append("        ,TPEEditName"+ Environment.NewLine);
            SbSql.Append("        ,TPEEditDate"+ Environment.NewLine);
            SbSql.Append("        ,Pressing1"+ Environment.NewLine);
            SbSql.Append("        ,Pressing2"+ Environment.NewLine);
            SbSql.Append("        ,Folding1"+ Environment.NewLine);
            SbSql.Append("        ,Folding2"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormStatus"+ Environment.NewLine);
            SbSql.Append("        ,ExpectionFormDate"+ Environment.NewLine);
            SbSql.Append("        ,ThickFabricBulk"+ Environment.NewLine);
            SbSql.Append("        ,HangerPack"+ Environment.NewLine);
            SbSql.Append("        ,Construction"+ Environment.NewLine);
            SbSql.Append("        ,CDCodeNew"+ Environment.NewLine);
            SbSql.Append("        ,FitType"+ Environment.NewLine);
            SbSql.Append("        ,GearLine"+ Environment.NewLine);
            SbSql.Append("        ,ThreadVersion"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@ProgramID"); objParameter.Add("@ProgramID", DbType.String, Item.ProgramID);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            SbSql.Append("        ,@Model"); objParameter.Add("@Model", DbType.String, Item.Model);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@StyleName"); objParameter.Add("@StyleName", DbType.String, Item.StyleName);
            SbSql.Append("        ,@ComboType"); objParameter.Add("@ComboType", DbType.String, Item.ComboType);
            SbSql.Append("        ,@CdCodeID"); objParameter.Add("@CdCodeID", DbType.String, Item.CdCodeID);
            SbSql.Append("        ,@ApparelType"); objParameter.Add("@ApparelType", DbType.String, Item.ApparelType);
            SbSql.Append("        ,@FabricType"); objParameter.Add("@FabricType", DbType.String, Item.FabricType);
            SbSql.Append("        ,@Contents"); objParameter.Add("@Contents", DbType.String, Item.Contents);
            SbSql.Append("        ,@GMTLT"); objParameter.Add("@GMTLT", DbType.String, Item.GMTLT);
            SbSql.Append("        ,@CPU"); objParameter.Add("@CPU", DbType.String, Item.CPU);
            SbSql.Append("        ,@Factories"); objParameter.Add("@Factories", DbType.String, Item.Factories);
            SbSql.Append("        ,@FTYRemark"); objParameter.Add("@FTYRemark", DbType.String, Item.FTYRemark);
            SbSql.Append("        ,@SampleSMR"); objParameter.Add("@SampleSMR", DbType.String, Item.SampleSMR);
            SbSql.Append("        ,@SampleMRHandle"); objParameter.Add("@SampleMRHandle", DbType.String, Item.SampleMRHandle);
            SbSql.Append("        ,@BulkSMR"); objParameter.Add("@BulkSMR", DbType.String, Item.BulkSMR);
            SbSql.Append("        ,@BulkMRHandle"); objParameter.Add("@BulkMRHandle", DbType.String, Item.BulkMRHandle);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@RainwearTestPassed"); objParameter.Add("@RainwearTestPassed", DbType.String, Item.RainwearTestPassed);
            SbSql.Append("        ,@SizePage"); objParameter.Add("@SizePage", DbType.String, Item.SizePage);
            SbSql.Append("        ,@SizeRange"); objParameter.Add("@SizeRange", DbType.String, Item.SizeRange);
            SbSql.Append("        ,@CTNQty"); objParameter.Add("@CTNQty", DbType.String, Item.CTNQty);
            SbSql.Append("        ,@StdCost"); objParameter.Add("@StdCost", DbType.String, Item.StdCost);
            SbSql.Append("        ,@Processes"); objParameter.Add("@Processes", DbType.String, Item.Processes);
            SbSql.Append("        ,@ArtworkCost"); objParameter.Add("@ArtworkCost", DbType.String, Item.ArtworkCost);
            SbSql.Append("        ,@Picture1"); objParameter.Add("@Picture1", DbType.String, Item.Picture1);
            SbSql.Append("        ,@Picture2"); objParameter.Add("@Picture2", DbType.String, Item.Picture2);
            SbSql.Append("        ,@Label"); objParameter.Add("@Label", DbType.String, Item.Label);
            SbSql.Append("        ,@Packing"); objParameter.Add("@Packing", DbType.String, Item.Packing);
            SbSql.Append("        ,@IETMSID"); objParameter.Add("@IETMSID", DbType.String, Item.IETMSID);
            SbSql.Append("        ,@IETMSVersion"); objParameter.Add("@IETMSVersion", DbType.String, Item.IETMSVersion);
            SbSql.Append("        ,@IEImportName"); objParameter.Add("@IEImportName", DbType.String, Item.IEImportName);
            SbSql.Append("        ,@IEImportDate"); objParameter.Add("@IEImportDate", DbType.DateTime, Item.IEImportDate);
            SbSql.Append("        ,@ApvDate"); objParameter.Add("@ApvDate", DbType.DateTime, Item.ApvDate);
            SbSql.Append("        ,@ApvName"); objParameter.Add("@ApvName", DbType.String, Item.ApvName);
            SbSql.Append("        ,@CareCode"); objParameter.Add("@CareCode", DbType.String, Item.CareCode);
            SbSql.Append("        ,@SpecialMark"); objParameter.Add("@SpecialMark", DbType.String, Item.SpecialMark);
            SbSql.Append("        ,@Lining"); objParameter.Add("@Lining", DbType.String, Item.Lining);
            SbSql.Append("        ,@StyleUnit"); objParameter.Add("@StyleUnit", DbType.String, Item.StyleUnit);
            SbSql.Append("        ,@ExpectionForm"); objParameter.Add("@ExpectionForm", DbType.String, Item.ExpectionForm);
            SbSql.Append("        ,@ExpectionFormRemark"); objParameter.Add("@ExpectionFormRemark", DbType.String, Item.ExpectionFormRemark);
            SbSql.Append("        ,@LocalMR"); objParameter.Add("@LocalMR", DbType.String, Item.LocalMR);
            SbSql.Append("        ,@LocalStyle"); objParameter.Add("@LocalStyle", DbType.String, Item.LocalStyle);
            SbSql.Append("        ,@PPMeeting"); objParameter.Add("@PPMeeting", DbType.String, Item.PPMeeting);
            SbSql.Append("        ,@NoNeedPPMeeting"); objParameter.Add("@NoNeedPPMeeting", DbType.String, Item.NoNeedPPMeeting);
            SbSql.Append("        ,@SampleApv"); objParameter.Add("@SampleApv", DbType.DateTime, Item.SampleApv);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@SizeUnit"); objParameter.Add("@SizeUnit", DbType.String, Item.SizeUnit);
            SbSql.Append("        ,@ModularParent"); objParameter.Add("@ModularParent", DbType.String, Item.ModularParent);
            SbSql.Append("        ,@CPUAdjusted"); objParameter.Add("@CPUAdjusted", DbType.String, Item.CPUAdjusted);
            SbSql.Append("        ,@Phase"); objParameter.Add("@Phase", DbType.String, Item.Phase);
            SbSql.Append("        ,@Gender"); objParameter.Add("@Gender", DbType.String, Item.Gender);
            SbSql.Append("        ,@ThreadEditname"); objParameter.Add("@ThreadEditname", DbType.String, Item.ThreadEditname);
            SbSql.Append("        ,@ThreadEditdate"); objParameter.Add("@ThreadEditdate", DbType.DateTime, Item.ThreadEditdate);
            SbSql.Append("        ,@ThickFabric"); objParameter.Add("@ThickFabric", DbType.String, Item.ThickFabric);
            SbSql.Append("        ,@DyeingID"); objParameter.Add("@DyeingID", DbType.String, Item.DyeingID);
            SbSql.Append("        ,@TPEEditName"); objParameter.Add("@TPEEditName", DbType.String, Item.TPEEditName);
            SbSql.Append("        ,@TPEEditDate"); objParameter.Add("@TPEEditDate", DbType.DateTime, Item.TPEEditDate);
            SbSql.Append("        ,@Pressing1"); objParameter.Add("@Pressing1", DbType.Int32, Item.Pressing1);
            SbSql.Append("        ,@Pressing2"); objParameter.Add("@Pressing2", DbType.Int32, Item.Pressing2);
            SbSql.Append("        ,@Folding1"); objParameter.Add("@Folding1", DbType.Int32, Item.Folding1);
            SbSql.Append("        ,@Folding2"); objParameter.Add("@Folding2", DbType.Int32, Item.Folding2);
            SbSql.Append("        ,@ExpectionFormStatus"); objParameter.Add("@ExpectionFormStatus", DbType.String, Item.ExpectionFormStatus);
            SbSql.Append("        ,@ExpectionFormDate"); objParameter.Add("@ExpectionFormDate", DbType.String, Item.ExpectionFormDate);
            SbSql.Append("        ,@ThickFabricBulk"); objParameter.Add("@ThickFabricBulk", DbType.String, Item.ThickFabricBulk);
            SbSql.Append("        ,@HangerPack"); objParameter.Add("@HangerPack", DbType.String, Item.HangerPack);
            SbSql.Append("        ,@Construction"); objParameter.Add("@Construction", DbType.String, Item.Construction);
            SbSql.Append("        ,@CDCodeNew"); objParameter.Add("@CDCodeNew", DbType.String, Item.CDCodeNew);
            SbSql.Append("        ,@FitType"); objParameter.Add("@FitType", DbType.String, Item.FitType);
            SbSql.Append("        ,@GearLine"); objParameter.Add("@GearLine", DbType.String, Item.GearLine);
            SbSql.Append("        ,@ThreadVersion"); objParameter.Add("@ThreadVersion", DbType.String, Item.ThreadVersion);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新款式資料基本檔(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新款式資料基本檔
        /// </summary>
        /// <param name="Item">款式資料基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(Style Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Style]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.Ukey != null) { SbSql.Append(",Ukey=@Ukey"+ Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey);}
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID"+ Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID);}
            if (Item.ProgramID != null) { SbSql.Append(",ProgramID=@ProgramID"+ Environment.NewLine); objParameter.Add("@ProgramID", DbType.String, Item.ProgramID);}
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID"+ Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);}
            if (Item.Model != null) { SbSql.Append(",Model=@Model"+ Environment.NewLine); objParameter.Add("@Model", DbType.String, Item.Model);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.StyleName != null) { SbSql.Append(",StyleName=@StyleName"+ Environment.NewLine); objParameter.Add("@StyleName", DbType.String, Item.StyleName);}
            if (Item.ComboType != null) { SbSql.Append(",ComboType=@ComboType"+ Environment.NewLine); objParameter.Add("@ComboType", DbType.String, Item.ComboType);}
            if (Item.CdCodeID != null) { SbSql.Append(",CdCodeID=@CdCodeID"+ Environment.NewLine); objParameter.Add("@CdCodeID", DbType.String, Item.CdCodeID);}
            if (Item.ApparelType != null) { SbSql.Append(",ApparelType=@ApparelType"+ Environment.NewLine); objParameter.Add("@ApparelType", DbType.String, Item.ApparelType);}
            if (Item.FabricType != null) { SbSql.Append(",FabricType=@FabricType"+ Environment.NewLine); objParameter.Add("@FabricType", DbType.String, Item.FabricType);}
            if (Item.Contents != null) { SbSql.Append(",Contents=@Contents"+ Environment.NewLine); objParameter.Add("@Contents", DbType.String, Item.Contents);}
            if (Item.GMTLT != null) { SbSql.Append(",GMTLT=@GMTLT"+ Environment.NewLine); objParameter.Add("@GMTLT", DbType.String, Item.GMTLT);}
            if (Item.CPU != null) { SbSql.Append(",CPU=@CPU"+ Environment.NewLine); objParameter.Add("@CPU", DbType.String, Item.CPU);}
            if (Item.Factories != null) { SbSql.Append(",Factories=@Factories"+ Environment.NewLine); objParameter.Add("@Factories", DbType.String, Item.Factories);}
            if (Item.FTYRemark != null) { SbSql.Append(",FTYRemark=@FTYRemark"+ Environment.NewLine); objParameter.Add("@FTYRemark", DbType.String, Item.FTYRemark);}
            if (Item.SampleSMR != null) { SbSql.Append(",SampleSMR=@SampleSMR"+ Environment.NewLine); objParameter.Add("@SampleSMR", DbType.String, Item.SampleSMR);}
            if (Item.SampleMRHandle != null) { SbSql.Append(",SampleMRHandle=@SampleMRHandle"+ Environment.NewLine); objParameter.Add("@SampleMRHandle", DbType.String, Item.SampleMRHandle);}
            if (Item.BulkSMR != null) { SbSql.Append(",BulkSMR=@BulkSMR"+ Environment.NewLine); objParameter.Add("@BulkSMR", DbType.String, Item.BulkSMR);}
            if (Item.BulkMRHandle != null) { SbSql.Append(",BulkMRHandle=@BulkMRHandle"+ Environment.NewLine); objParameter.Add("@BulkMRHandle", DbType.String, Item.BulkMRHandle);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.RainwearTestPassed != null) { SbSql.Append(",RainwearTestPassed=@RainwearTestPassed"+ Environment.NewLine); objParameter.Add("@RainwearTestPassed", DbType.String, Item.RainwearTestPassed);}
            if (Item.SizePage != null) { SbSql.Append(",SizePage=@SizePage"+ Environment.NewLine); objParameter.Add("@SizePage", DbType.String, Item.SizePage);}
            if (Item.SizeRange != null) { SbSql.Append(",SizeRange=@SizeRange"+ Environment.NewLine); objParameter.Add("@SizeRange", DbType.String, Item.SizeRange);}
            if (Item.CTNQty != null) { SbSql.Append(",CTNQty=@CTNQty"+ Environment.NewLine); objParameter.Add("@CTNQty", DbType.String, Item.CTNQty);}
            if (Item.StdCost != null) { SbSql.Append(",StdCost=@StdCost"+ Environment.NewLine); objParameter.Add("@StdCost", DbType.String, Item.StdCost);}
            if (Item.Processes != null) { SbSql.Append(",Processes=@Processes"+ Environment.NewLine); objParameter.Add("@Processes", DbType.String, Item.Processes);}
            if (Item.ArtworkCost != null) { SbSql.Append(",ArtworkCost=@ArtworkCost"+ Environment.NewLine); objParameter.Add("@ArtworkCost", DbType.String, Item.ArtworkCost);}
            if (Item.Picture1 != null) { SbSql.Append(",Picture1=@Picture1"+ Environment.NewLine); objParameter.Add("@Picture1", DbType.String, Item.Picture1);}
            if (Item.Picture2 != null) { SbSql.Append(",Picture2=@Picture2"+ Environment.NewLine); objParameter.Add("@Picture2", DbType.String, Item.Picture2);}
            if (Item.Label != null) { SbSql.Append(",Label=@Label"+ Environment.NewLine); objParameter.Add("@Label", DbType.String, Item.Label);}
            if (Item.Packing != null) { SbSql.Append(",Packing=@Packing"+ Environment.NewLine); objParameter.Add("@Packing", DbType.String, Item.Packing);}
            if (Item.IETMSID != null) { SbSql.Append(",IETMSID=@IETMSID"+ Environment.NewLine); objParameter.Add("@IETMSID", DbType.String, Item.IETMSID);}
            if (Item.IETMSVersion != null) { SbSql.Append(",IETMSVersion=@IETMSVersion"+ Environment.NewLine); objParameter.Add("@IETMSVersion", DbType.String, Item.IETMSVersion);}
            if (Item.IEImportName != null) { SbSql.Append(",IEImportName=@IEImportName"+ Environment.NewLine); objParameter.Add("@IEImportName", DbType.String, Item.IEImportName);}
            if (Item.IEImportDate != null) { SbSql.Append(",IEImportDate=@IEImportDate"+ Environment.NewLine); objParameter.Add("@IEImportDate", DbType.DateTime, Item.IEImportDate);}
            if (Item.ApvDate != null) { SbSql.Append(",ApvDate=@ApvDate"+ Environment.NewLine); objParameter.Add("@ApvDate", DbType.DateTime, Item.ApvDate);}
            if (Item.ApvName != null) { SbSql.Append(",ApvName=@ApvName"+ Environment.NewLine); objParameter.Add("@ApvName", DbType.String, Item.ApvName);}
            if (Item.CareCode != null) { SbSql.Append(",CareCode=@CareCode"+ Environment.NewLine); objParameter.Add("@CareCode", DbType.String, Item.CareCode);}
            if (Item.SpecialMark != null) { SbSql.Append(",SpecialMark=@SpecialMark"+ Environment.NewLine); objParameter.Add("@SpecialMark", DbType.String, Item.SpecialMark);}
            if (Item.Lining != null) { SbSql.Append(",Lining=@Lining"+ Environment.NewLine); objParameter.Add("@Lining", DbType.String, Item.Lining);}
            if (Item.StyleUnit != null) { SbSql.Append(",StyleUnit=@StyleUnit"+ Environment.NewLine); objParameter.Add("@StyleUnit", DbType.String, Item.StyleUnit);}
            if (Item.ExpectionForm != null) { SbSql.Append(",ExpectionForm=@ExpectionForm"+ Environment.NewLine); objParameter.Add("@ExpectionForm", DbType.String, Item.ExpectionForm);}
            if (Item.ExpectionFormRemark != null) { SbSql.Append(",ExpectionFormRemark=@ExpectionFormRemark"+ Environment.NewLine); objParameter.Add("@ExpectionFormRemark", DbType.String, Item.ExpectionFormRemark);}
            if (Item.LocalMR != null) { SbSql.Append(",LocalMR=@LocalMR"+ Environment.NewLine); objParameter.Add("@LocalMR", DbType.String, Item.LocalMR);}
            if (Item.LocalStyle != null) { SbSql.Append(",LocalStyle=@LocalStyle"+ Environment.NewLine); objParameter.Add("@LocalStyle", DbType.String, Item.LocalStyle);}
            if (Item.PPMeeting != null) { SbSql.Append(",PPMeeting=@PPMeeting"+ Environment.NewLine); objParameter.Add("@PPMeeting", DbType.String, Item.PPMeeting);}
            if (Item.NoNeedPPMeeting != null) { SbSql.Append(",NoNeedPPMeeting=@NoNeedPPMeeting"+ Environment.NewLine); objParameter.Add("@NoNeedPPMeeting", DbType.String, Item.NoNeedPPMeeting);}
            if (Item.SampleApv != null) { SbSql.Append(",SampleApv=@SampleApv"+ Environment.NewLine); objParameter.Add("@SampleApv", DbType.DateTime, Item.SampleApv);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.SizeUnit != null) { SbSql.Append(",SizeUnit=@SizeUnit"+ Environment.NewLine); objParameter.Add("@SizeUnit", DbType.String, Item.SizeUnit);}
            if (Item.ModularParent != null) { SbSql.Append(",ModularParent=@ModularParent"+ Environment.NewLine); objParameter.Add("@ModularParent", DbType.String, Item.ModularParent);}
            if (Item.CPUAdjusted != null) { SbSql.Append(",CPUAdjusted=@CPUAdjusted"+ Environment.NewLine); objParameter.Add("@CPUAdjusted", DbType.String, Item.CPUAdjusted);}
            if (Item.Phase != null) { SbSql.Append(",Phase=@Phase"+ Environment.NewLine); objParameter.Add("@Phase", DbType.String, Item.Phase);}
            if (Item.Gender != null) { SbSql.Append(",Gender=@Gender"+ Environment.NewLine); objParameter.Add("@Gender", DbType.String, Item.Gender);}
            if (Item.ThreadEditname != null) { SbSql.Append(",ThreadEditname=@ThreadEditname"+ Environment.NewLine); objParameter.Add("@ThreadEditname", DbType.String, Item.ThreadEditname);}
            if (Item.ThreadEditdate != null) { SbSql.Append(",ThreadEditdate=@ThreadEditdate"+ Environment.NewLine); objParameter.Add("@ThreadEditdate", DbType.DateTime, Item.ThreadEditdate);}
            if (Item.ThickFabric != null) { SbSql.Append(",ThickFabric=@ThickFabric"+ Environment.NewLine); objParameter.Add("@ThickFabric", DbType.String, Item.ThickFabric);}
            if (Item.DyeingID != null) { SbSql.Append(",DyeingID=@DyeingID"+ Environment.NewLine); objParameter.Add("@DyeingID", DbType.String, Item.DyeingID);}
            if (Item.TPEEditName != null) { SbSql.Append(",TPEEditName=@TPEEditName"+ Environment.NewLine); objParameter.Add("@TPEEditName", DbType.String, Item.TPEEditName);}
            if (Item.TPEEditDate != null) { SbSql.Append(",TPEEditDate=@TPEEditDate"+ Environment.NewLine); objParameter.Add("@TPEEditDate", DbType.DateTime, Item.TPEEditDate);}
            if (Item.Pressing1 != null) { SbSql.Append(",Pressing1=@Pressing1"+ Environment.NewLine); objParameter.Add("@Pressing1", DbType.Int32, Item.Pressing1);}
            if (Item.Pressing2 != null) { SbSql.Append(",Pressing2=@Pressing2"+ Environment.NewLine); objParameter.Add("@Pressing2", DbType.Int32, Item.Pressing2);}
            if (Item.Folding1 != null) { SbSql.Append(",Folding1=@Folding1"+ Environment.NewLine); objParameter.Add("@Folding1", DbType.Int32, Item.Folding1);}
            if (Item.Folding2 != null) { SbSql.Append(",Folding2=@Folding2"+ Environment.NewLine); objParameter.Add("@Folding2", DbType.Int32, Item.Folding2);}
            if (Item.ExpectionFormStatus != null) { SbSql.Append(",ExpectionFormStatus=@ExpectionFormStatus"+ Environment.NewLine); objParameter.Add("@ExpectionFormStatus", DbType.String, Item.ExpectionFormStatus);}
            if (Item.ExpectionFormDate != null) { SbSql.Append(",ExpectionFormDate=@ExpectionFormDate"+ Environment.NewLine); objParameter.Add("@ExpectionFormDate", DbType.String, Item.ExpectionFormDate);}
            if (Item.ThickFabricBulk != null) { SbSql.Append(",ThickFabricBulk=@ThickFabricBulk"+ Environment.NewLine); objParameter.Add("@ThickFabricBulk", DbType.String, Item.ThickFabricBulk);}
            if (Item.HangerPack != null) { SbSql.Append(",HangerPack=@HangerPack"+ Environment.NewLine); objParameter.Add("@HangerPack", DbType.String, Item.HangerPack);}
            if (Item.Construction != null) { SbSql.Append(",Construction=@Construction"+ Environment.NewLine); objParameter.Add("@Construction", DbType.String, Item.Construction);}
            if (Item.CDCodeNew != null) { SbSql.Append(",CDCodeNew=@CDCodeNew"+ Environment.NewLine); objParameter.Add("@CDCodeNew", DbType.String, Item.CDCodeNew);}
            if (Item.FitType != null) { SbSql.Append(",FitType=@FitType"+ Environment.NewLine); objParameter.Add("@FitType", DbType.String, Item.FitType);}
            if (Item.GearLine != null) { SbSql.Append(",GearLine=@GearLine"+ Environment.NewLine); objParameter.Add("@GearLine", DbType.String, Item.GearLine);}
            if (Item.ThreadVersion != null) { SbSql.Append(",ThreadVersion=@ThreadVersion"+ Environment.NewLine); objParameter.Add("@ThreadVersion", DbType.String, Item.ThreadVersion);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除款式資料基本檔(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除款式資料基本檔
        /// </summary>
        /// <param name="Item">款式資料基本檔成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(Style Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Style]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public string GetSizeUnitByPOID(string POID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                 { "@POID", DbType.String, POID} ,
            };

            string sqlcmd = @"
select SizeUnit 
from Style  with (nolock)
where   ukey = (select styleUkey from orders with (nolock) where ID = @POID)
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.Rows[0]["SizeUnit"].ToString();
            }
            else
            {
                return string.Empty;
            }
        }
        #endregion
    }
}
