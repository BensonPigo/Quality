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
    public class BrandProvider : SQLDAL, IBrandProvider
    {
        #region 底層連線
        public BrandProvider(string ConString) : base(ConString) { }
        public BrandProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Brand(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Brand
        /// </summary>
        /// <param name="Item">Brand成員</param>
        /// <returns>回傳Brand</returns>
		/// <info>Author: Admin; Date: 2021/08/10  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/10  1.00    Admin        Create
        /// </history>
        public IList<Brand> Get()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,NameCH"+ Environment.NewLine);
            SbSql.Append("        ,NameEN"+ Environment.NewLine);
            SbSql.Append("        ,CountryID"+ Environment.NewLine);
            SbSql.Append("        ,BuyerID"+ Environment.NewLine);
            SbSql.Append("        ,Tel"+ Environment.NewLine);
            SbSql.Append("        ,Fax"+ Environment.NewLine);
            SbSql.Append("        ,Contact1"+ Environment.NewLine);
            SbSql.Append("        ,Contact2"+ Environment.NewLine);
            SbSql.Append("        ,AddressCH"+ Environment.NewLine);
            SbSql.Append("        ,AddressEN"+ Environment.NewLine);
            SbSql.Append("        ,CurrencyID"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Customize1"+ Environment.NewLine);
            SbSql.Append("        ,Customize2"+ Environment.NewLine);
            SbSql.Append("        ,Customize3"+ Environment.NewLine);
            SbSql.Append("        ,Commission"+ Environment.NewLine);
            SbSql.Append("        ,ZipCode"+ Environment.NewLine);
            SbSql.Append("        ,Email"+ Environment.NewLine);
            SbSql.Append("        ,MrTeam"+ Environment.NewLine);
            SbSql.Append("        ,BrandGroup"+ Environment.NewLine);
            SbSql.Append("        ,ApparelXlt"+ Environment.NewLine);
            SbSql.Append("        ,LossSampleFabric"+ Environment.NewLine);
            SbSql.Append("        ,PayTermARIDBulk"+ Environment.NewLine);
            SbSql.Append("        ,PayTermARIDSample"+ Environment.NewLine);
            SbSql.Append("        ,BrandFactoryAreaCaption"+ Environment.NewLine);
            SbSql.Append("        ,BrandFactoryCodeCaption"+ Environment.NewLine);
            SbSql.Append("        ,BrandFactoryVendorCaption"+ Environment.NewLine);
            SbSql.Append("        ,ShipCode"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,LossSampleAccessory"+ Environment.NewLine);
            SbSql.Append("        ,ShipLeader"+ Environment.NewLine);
            SbSql.Append("        ,ShipLeaderEditDate"+ Environment.NewLine);
            SbSql.Append("        ,OTDExtension"+ Environment.NewLine);
            SbSql.Append("        ,UseRatioRule"+ Environment.NewLine);
            SbSql.Append("        ,UseRatioRule_Thick"+ Environment.NewLine);
            SbSql.Append("FROM [Brand]"+ Environment.NewLine);



            return ExecuteList<Brand>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Brand(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Brand
        /// </summary>
        /// <param name="Item">Brand成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/10  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/10  1.00    Admin        Create
        /// </history>
        public int Update(Brand Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Brand]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.NameCH != null) { SbSql.Append(",NameCH=@NameCH"+ Environment.NewLine); objParameter.Add("@NameCH", DbType.String, Item.NameCH);}
            if (Item.NameEN != null) { SbSql.Append(",NameEN=@NameEN"+ Environment.NewLine); objParameter.Add("@NameEN", DbType.String, Item.NameEN);}
            if (Item.CountryID != null) { SbSql.Append(",CountryID=@CountryID"+ Environment.NewLine); objParameter.Add("@CountryID", DbType.String, Item.CountryID);}
            if (Item.BuyerID != null) { SbSql.Append(",BuyerID=@BuyerID"+ Environment.NewLine); objParameter.Add("@BuyerID", DbType.String, Item.BuyerID);}
            if (Item.Tel != null) { SbSql.Append(",Tel=@Tel"+ Environment.NewLine); objParameter.Add("@Tel", DbType.String, Item.Tel);}
            if (Item.Fax != null) { SbSql.Append(",Fax=@Fax"+ Environment.NewLine); objParameter.Add("@Fax", DbType.String, Item.Fax);}
            if (Item.Contact1 != null) { SbSql.Append(",Contact1=@Contact1"+ Environment.NewLine); objParameter.Add("@Contact1", DbType.String, Item.Contact1);}
            if (Item.Contact2 != null) { SbSql.Append(",Contact2=@Contact2"+ Environment.NewLine); objParameter.Add("@Contact2", DbType.String, Item.Contact2);}
            if (Item.AddressCH != null) { SbSql.Append(",AddressCH=@AddressCH"+ Environment.NewLine); objParameter.Add("@AddressCH", DbType.String, Item.AddressCH);}
            if (Item.AddressEN != null) { SbSql.Append(",AddressEN=@AddressEN"+ Environment.NewLine); objParameter.Add("@AddressEN", DbType.String, Item.AddressEN);}
            if (Item.CurrencyID != null) { SbSql.Append(",CurrencyID=@CurrencyID"+ Environment.NewLine); objParameter.Add("@CurrencyID", DbType.String, Item.CurrencyID);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark ?? "");}
            if (Item.Customize1 != null) { SbSql.Append(",Customize1=@Customize1"+ Environment.NewLine); objParameter.Add("@Customize1", DbType.String, Item.Customize1);}
            if (Item.Customize2 != null) { SbSql.Append(",Customize2=@Customize2"+ Environment.NewLine); objParameter.Add("@Customize2", DbType.String, Item.Customize2);}
            if (Item.Customize3 != null) { SbSql.Append(",Customize3=@Customize3"+ Environment.NewLine); objParameter.Add("@Customize3", DbType.String, Item.Customize3);}
            if (Item.Commission != null) { SbSql.Append(",Commission=@Commission"+ Environment.NewLine); objParameter.Add("@Commission", DbType.String, Item.Commission);}
            if (Item.ZipCode != null) { SbSql.Append(",ZipCode=@ZipCode"+ Environment.NewLine); objParameter.Add("@ZipCode", DbType.String, Item.ZipCode);}
            if (Item.Email != null) { SbSql.Append(",Email=@Email"+ Environment.NewLine); objParameter.Add("@Email", DbType.String, Item.Email);}
            if (Item.MrTeam != null) { SbSql.Append(",MrTeam=@MrTeam"+ Environment.NewLine); objParameter.Add("@MrTeam", DbType.String, Item.MrTeam);}
            if (Item.BrandGroup != null) { SbSql.Append(",BrandGroup=@BrandGroup"+ Environment.NewLine); objParameter.Add("@BrandGroup", DbType.String, Item.BrandGroup);}
            if (Item.ApparelXlt != null) { SbSql.Append(",ApparelXlt=@ApparelXlt"+ Environment.NewLine); objParameter.Add("@ApparelXlt", DbType.String, Item.ApparelXlt);}
            if (Item.LossSampleFabric != null) { SbSql.Append(",LossSampleFabric=@LossSampleFabric"+ Environment.NewLine); objParameter.Add("@LossSampleFabric", DbType.String, Item.LossSampleFabric);}
            if (Item.PayTermARIDBulk != null) { SbSql.Append(",PayTermARIDBulk=@PayTermARIDBulk"+ Environment.NewLine); objParameter.Add("@PayTermARIDBulk", DbType.String, Item.PayTermARIDBulk);}
            if (Item.PayTermARIDSample != null) { SbSql.Append(",PayTermARIDSample=@PayTermARIDSample"+ Environment.NewLine); objParameter.Add("@PayTermARIDSample", DbType.String, Item.PayTermARIDSample);}
            if (Item.BrandFactoryAreaCaption != null) { SbSql.Append(",BrandFactoryAreaCaption=@BrandFactoryAreaCaption"+ Environment.NewLine); objParameter.Add("@BrandFactoryAreaCaption", DbType.String, Item.BrandFactoryAreaCaption);}
            if (Item.BrandFactoryCodeCaption != null) { SbSql.Append(",BrandFactoryCodeCaption=@BrandFactoryCodeCaption"+ Environment.NewLine); objParameter.Add("@BrandFactoryCodeCaption", DbType.String, Item.BrandFactoryCodeCaption);}
            if (Item.BrandFactoryVendorCaption != null) { SbSql.Append(",BrandFactoryVendorCaption=@BrandFactoryVendorCaption"+ Environment.NewLine); objParameter.Add("@BrandFactoryVendorCaption", DbType.String, Item.BrandFactoryVendorCaption);}
            if (Item.ShipCode != null) { SbSql.Append(",ShipCode=@ShipCode"+ Environment.NewLine); objParameter.Add("@ShipCode", DbType.String, Item.ShipCode);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.LossSampleAccessory != null) { SbSql.Append(",LossSampleAccessory=@LossSampleAccessory"+ Environment.NewLine); objParameter.Add("@LossSampleAccessory", DbType.String, Item.LossSampleAccessory);}
            if (Item.ShipLeader != null) { SbSql.Append(",ShipLeader=@ShipLeader"+ Environment.NewLine); objParameter.Add("@ShipLeader", DbType.String, Item.ShipLeader);}
            if (Item.ShipLeaderEditDate != null) { SbSql.Append(",ShipLeaderEditDate=@ShipLeaderEditDate"+ Environment.NewLine); objParameter.Add("@ShipLeaderEditDate", DbType.DateTime, Item.ShipLeaderEditDate);}
            if (Item.OTDExtension != null) { SbSql.Append(",OTDExtension=@OTDExtension"+ Environment.NewLine); objParameter.Add("@OTDExtension", DbType.Int32, Item.OTDExtension);}
            if (Item.UseRatioRule != null) { SbSql.Append(",UseRatioRule=@UseRatioRule"+ Environment.NewLine); objParameter.Add("@UseRatioRule", DbType.String, Item.UseRatioRule);}
            if (Item.UseRatioRule_Thick != null) { SbSql.Append(",UseRatioRule_Thick=@UseRatioRule_Thick"+ Environment.NewLine); objParameter.Add("@UseRatioRule_Thick", DbType.String, Item.UseRatioRule_Thick);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Brand(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Brand
        /// </summary>
        /// <param name="Item">Brand成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/10  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/10  1.00    Admin        Create
        /// </history>
        public int Delete(Brand Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Brand]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
