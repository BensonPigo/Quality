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
    /*(MockupCrockingProvider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/19  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19  1.00    Admin        Create
    /// </history>
    public class MockupCrockingProvider : SQLDAL, IMockupCrockingProvider
    {
        #region 底層連線
        public MockupCrockingProvider(string ConString) : base(ConString) { }
        public MockupCrockingProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public IList<MockupCrocking> Get(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ReportNo"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,ArtworkTypeID"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,T1Subcon"+ Environment.NewLine);
            SbSql.Append("        ,TestDate"+ Environment.NewLine);
            SbSql.Append("        ,ReceivedDate"+ Environment.NewLine);
            SbSql.Append("        ,ReleasedDate"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Technician"+ Environment.NewLine);
            SbSql.Append("        ,MR"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("FROM [MockupCrocking]"+ Environment.NewLine);

            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }

            return ExecuteList<MockupCrocking>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Create(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [MockupCrocking]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ReportNo"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,ArtworkTypeID"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,T1Subcon"+ Environment.NewLine);
            SbSql.Append("        ,TestDate"+ Environment.NewLine);
            SbSql.Append("        ,ReceivedDate"+ Environment.NewLine);
            SbSql.Append("        ,ReleasedDate"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Technician"+ Environment.NewLine);
            SbSql.Append("        ,MR"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ReportNo"); objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@ArtworkTypeID"); objParameter.Add("@ArtworkTypeID", DbType.String, Item.ArtworkTypeID);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@T1Subcon"); objParameter.Add("@T1Subcon", DbType.String, Item.T1Subcon);
            SbSql.Append("        ,@TestDate"); objParameter.Add("@TestDate", DbType.String, Item.TestDate);
            SbSql.Append("        ,@ReceivedDate"); objParameter.Add("@ReceivedDate", DbType.String, Item.ReceivedDate);
            SbSql.Append("        ,@ReleasedDate"); objParameter.Add("@ReleasedDate", DbType.String, Item.ReleasedDate);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Technician"); objParameter.Add("@Technician", DbType.String, Item.Technician);
            SbSql.Append("        ,@MR"); objParameter.Add("@MR", DbType.String, Item.MR);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);

            SbSql.Append("        ,@TestBeforePicture");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            SbSql.Append("        ,@TestAfterPicture");
            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Update(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [MockupCrocking]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);

            if (Item.POID != null) { SbSql.Append(",POID=@POID"+ Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID);}
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID"+ Environment.NewLine); objParameter.Add("@StyleID", DbType.String, Item.StyleID);}
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID"+ Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);}
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID"+ Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.ArtworkTypeID != null) { SbSql.Append(",ArtworkTypeID=@ArtworkTypeID"+ Environment.NewLine); objParameter.Add("@ArtworkTypeID", DbType.String, Item.ArtworkTypeID);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.T1Subcon != null) { SbSql.Append(",T1Subcon=@T1Subcon"+ Environment.NewLine); objParameter.Add("@T1Subcon", DbType.String, Item.T1Subcon);}
            if (Item.TestDate != null) { SbSql.Append(",TestDate=@TestDate"+ Environment.NewLine); objParameter.Add("@TestDate", DbType.String, Item.TestDate);}
            if (Item.ReceivedDate != null) { SbSql.Append(",ReceivedDate=@ReceivedDate"+ Environment.NewLine); objParameter.Add("@ReceivedDate", DbType.String, Item.ReceivedDate);}
            if (Item.ReleasedDate != null) { SbSql.Append(",ReleasedDate=@ReleasedDate"+ Environment.NewLine); objParameter.Add("@ReleasedDate", DbType.String, Item.ReleasedDate);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Technician != null) { SbSql.Append(",Technician=@Technician"+ Environment.NewLine); objParameter.Add("@Technician", DbType.String, Item.Technician);}
            if (Item.MR != null) { SbSql.Append(",MR=@MR"+ Environment.NewLine); objParameter.Add("@MR", DbType.String, Item.MR);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}

            // 圖檔
            SbSql.Append(@"
    TestBeforePicture=@TestBeforePicture,
    TestAfterPicture=@TestAfterPicture
");
            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Delete(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE [MockupCrocking]"+ Environment.NewLine);
            SbSql.Append("WHERE ReportNo = @ReportNo" + Environment.NewLine);
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        public IList<MockupCrocking> GetMockupCrocking(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT
         ReportNo
        ,POID
        ,StyleID
        ,SeasonID
        ,BrandID
        ,Article
        ,ArtworkTypeID
        ,Remark
        ,T1Subcon
        ,TestDate
        ,ReceivedDate
        ,ReleasedDate
        ,Result
        ,Technician
        ,MR
        ,Type
        ,TestBeforePicture
        ,TestAfterPicture
        ,AddDate
        ,AddName
        ,EditDate
        ,EditName
        ,EditName
        ,Abb = (select Abb from LocalSupp where ID = T1Subcon)
        ,TechnicianName = (select name from pass1 p where p.ID = Technician)
        ,SignaturePic = (select PicPath from system) + (select t.SignaturePic from Technician t where t.ID = Technician)
FROM [MockupCrocking]
");
            SbSql.Append("Where 1 = 1" + Environment.NewLine);

            if (!string.IsNullOrEmpty(Item.ReportNo))
            {
                SbSql.Append("And ReportNo = @ReportNo" + Environment.NewLine);
                objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);
            }
            if (!string.IsNullOrEmpty(Item.BrandID))
            {
                SbSql.Append("And BrandID = @BrandID" + Environment.NewLine);
                objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            }
            if (!string.IsNullOrEmpty(Item.SeasonID))
            {
                SbSql.Append("And SeasonID = @SeasonID" + Environment.NewLine);
                objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            }
            if (!string.IsNullOrEmpty(Item.StyleID))
            {
                SbSql.Append("And StyleID = @StyleID" + Environment.NewLine);
                objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            }
            if (!string.IsNullOrEmpty(Item.Article))
            {
                SbSql.Append("And Article = @Article" + Environment.NewLine);
                objParameter.Add("@Article", DbType.String, Item.Article);
            }
            if (!string.IsNullOrEmpty(Item.Type))
            {
                SbSql.Append("And Type = @Type" + Environment.NewLine);
                objParameter.Add("@Type", DbType.String, Item.Type);
            }


            return ExecuteList<MockupCrocking>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int UpdatePicture(MockupCrocking Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            if (Item.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Item.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (Item.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Item.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            SbSql.Append(@"
UPDATE [MockupCrocking]
SET");
            if (Item.TestBeforePicture != null && Item.TestAfterPicture == null)
            {
                SbSql.Append(@"
    TestBeforePicture=@TestBeforePicture");

            }
            else if (Item.TestBeforePicture == null && Item.TestAfterPicture != null)
            {
                SbSql.Append(@"
    TestAfterPicture=@TestAfterPicture");

            }
            else
            {
                SbSql.Append(@"
    TestBeforePicture=@TestBeforePicture
    TestAfterPicture=@TestAfterPicture,");

            }

            SbSql.Append(@"
where ReportNo = @ReportNo
");
            objParameter.Add("@ReportNo", DbType.String, Item.ReportNo);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
