using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.Public;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolKit;
using System.Configuration;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class StickerTestProvider : SQLDAL
    {
        #region 底層連線
        public StickerTestProvider(string ConString) : base(ConString) { }
        public StickerTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<StickerTestItem> GetTestItems()
        {
            string sqlcmd = $@"select * from StickerTestItem";
            var tmp = ExecuteList<StickerTestItem>(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return tmp.Any() ? tmp.ToList() : new List<StickerTestItem>();
        }

        public List<SelectListItem> GetScales()
        {
            string sqlcmd = @"
select Text=ID , Value=ID
from Scale WITH(NOLOCK)  
WHERE Junk=0 
order by Value
";
            var tmp = ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return tmp.Any() ? tmp.ToList() : new List<SelectListItem>();
        }
        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(StickerTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();

            string SbSql = $@"

select o.*, oa.Article
from Orders o
inner join Order_Article oa on oa.id = o.ID
where o.Category ='B'  --只抓大貨單
";

            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql += $@" and o.ID = @OrderID" + Environment.NewLine;
                paras.Add("@OrderID", DbType.String, Req.OrderID);
            }

            var tmp = ExecuteList<DatabaseObject.ProductionDB.Orders>(CommandType.Text, SbSql, paras);


            return tmp.Any() ? tmp.ToList() : new List<DatabaseObject.ProductionDB.Orders>();
        }
        public List<StickerTest_Main> GetMainList(StickerTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
            };

            string sqlcmd = $@"
select a.*
        ,d.TestAfterPicture
        ,d.TestBeforePicture
from StickerTest a
left join PMSFile.dbo.StickerTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
where Junk = 0
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                sqlcmd += " and a.BrandID = @BrandID";
                objParameter.Add("@BrandID", Req.BrandID);
            }

            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                sqlcmd += " and a.SeasonID = @SeasonID";
                objParameter.Add("@SeasonID", Req.SeasonID);
            }

            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                sqlcmd += " and a.StyleID = @StyleID";
                objParameter.Add("@StyleID", Req.StyleID);
            }

            if (!string.IsNullOrEmpty(Req.Article))
            {
                sqlcmd += " and a.Article = @Article";
                objParameter.Add("@Article", Req.Article);
            }


            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo";
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<StickerTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<StickerTest_Main>();
        }
        public List<StickerTest_Detail> GetDetailList(StickerTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select distinct a.*
from StickerTest_Detail a WITH(NOLOCK)
inner join StickerTest b on a.ReportNo=b.ReportNo 
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and b.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<StickerTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<StickerTest_Detail>();
        }
        public List<StickerTest_Detail_Item> GetDetailItemList(StickerTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select c.*
from StickerTest a WITH(NOLOCK)
inner join StickerTest_Detail b  WITH(NOLOCK) on a.ReportNo = b.ReportNo
inner join StickerTest_Detail_Item c WITH(NOLOCK) on b.ReportNo = c.ReportNo AND b.EvaluationItem = c.EvaluationItem
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<StickerTest_Detail_Item>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<StickerTest_Detail_Item>();
        }
        public int Insert_StickerTest(StickerTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "ST", "StickerTest", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@FactoryID", DbType.String, Req.Main.FactoryID ?? ""} ,
                { "@BrandID", DbType.String, Req.Main.BrandID ?? ""} ,
                { "@SeasonID", DbType.String, Req.Main.SeasonID ?? ""} ,
                { "@StyleID", DbType.String, Req.Main.StyleID ?? ""} ,
                { "@Article", DbType.String, Req.Main.Article ?? ""} ,
                { "@OrderID", DbType.String, Req.Main.OrderID ?? ""} ,
                { "@FactoryID", DbType.String, Req.Main.FactoryID ?? ""} ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@AccRefNo", DbType.String, Req.Main.AccRefNo ?? "" } ,
                { "@AccColor", DbType.String, Req.Main.AccColor ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Main.Temperature } ,
                { "@Time", DbType.Int32, Req.Main.Time } ,
                { "@Humidity", DbType.Decimal, Req.Main.Humidity } ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestAfterPicture != null)
            {
                objParameter.Add("@TestAfterPicture", Req.Main.TestAfterPicture);
            }
            else
            {
                objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"
INSERT INTO dbo.StickerTest
           (ReportNo
           ,FactoryID
           ,OrderID
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,SubmitDate
           ,Seq1
           ,Seq2
           ,TestStandard
           ,FabricRefNo
           ,FabricColor
           ,FabricDescription
           ,AccRefNo
           ,AccColor
           ,Temperature
           ,Time
           ,Humidity
           ,Result
           ,Status
           ,AddDate
           ,AddName)
VALUES
           (@ReportNo
           ,(select top 1 FactoryID from SciProduction_Orders with(NOLOCK) where ID = @OrderID)
           ,@OrderID
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,@Article
           ,@SubmitDate
           ,@Seq1
           ,@Seq2
           ,@TestStandard
           ,@FabricRefNo
           ,@FabricColor
           ,@FabricDescription
           ,@AccRefNo
           ,@AccColor
           ,@Temperature
           ,@Time
           ,@Humidity
           ,@Result
           ,'New'
           ,GETDATE()
           ,@AddName)
;

INSERT INTO PMSFile.dbo.StickerTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_StickerTest(StickerTest_ViewModel Req, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@AccRefNo", DbType.String, Req.Main.AccRefNo ?? "" } ,
                { "@AccColor", DbType.String, Req.Main.AccColor ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Main.Temperature } ,
                { "@Time", DbType.Int32, Req.Main.Time } ,
                { "@Humidity", DbType.Decimal, Req.Main.Humidity } ,
                { "@Remark", DbType.String, Req.Main.Remark ?? "" } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestAfterPicture != null)
            {
                objParameter.Add("@TestAfterPicture", Req.Main.TestAfterPicture);
            }
            else
            {
                objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            string mainSqlCmd = $@"
UPDATE StickerTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,Remark = @Remark
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,TestStandard = @TestStandard
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricDescription = @FabricDescription
      ,AccRefNo = @AccRefNo
      ,AccColor = @AccColor
      ,Temperature = @Temperature
      ,Time = @Time
      ,Humidity = @Humidity
WHERE ReportNo = @ReportNo
;
UPDATE PMSFile.dbo.StickerTest
SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture
WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public string Delete_StickerTest(StickerTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@EditName", DbType.String, Req.Main.EditName } ,
            };

            string mainSqlCmd = $@"
delete from StickerTest WHERE ReportNo = @ReportNo
delete from StickerTest_Detail WHERE ReportNo = @ReportNo
delete from StickerTest_Detail_Item WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.StickerTest WHERE ReportNo = @ReportNo

";

            return ExecuteScalar(CommandType.Text, mainSqlCmd.ToString(), objParameter).ToString();
        }
        public bool Processe_StickerTest_Detail(StickerTest_ViewModel sources, string UserID)
        {
            List<StickerTest_Detail> oldDetailData = this.GetDetailList(new StickerTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<StickerTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<StickerTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "ReportNo,EvaluationItem",
                    "Scale,Result,Remark");

            string insertDetail = $@" ----寫入 StickerTest_Detail
INSERT INTO StickerTest_Detail
           (ReportNo,EvaluationItem,Scale,Result,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@Scale,@Result,@Remark,@UserID,GETDATE())
";
            string updateDetail = $@" ----更新 StickerTest_Detail
UPDATE StickerTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,Scale = @Scale
    ,Result = @Result
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND EvaluationItem = @EvaluationItem
;
";
            string deleteDetail = $@" ----刪除 StickerTest_Detail
DELETE FROM StickerTest_Detail where ReportNo = @ReportNo AND EvaluationItem = @EvaluationItem
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                listDetailPar.Add(new SqlParameter($"@Scale", detailItem.Scale ?? string.Empty));
                listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case CompareStateType.Delete:

                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case CompareStateType.None:

                        break;
                    default:
                        break;
                }

            }

            return true;

        }

        public bool Processe_StickerTest_Detail_Item(StickerTest_ViewModel sources, string UserID)
        {
            List<StickerTest_Detail_Item> oldDetailItemData = this.GetDetailItemList(new StickerTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<StickerTest_Detail_Item> needUpdateDetailItemList =
                PublicClass.CompareListValue<StickerTest_Detail_Item>(
                    sources.DetailItemList,
                    oldDetailItemData,
                    "ReportNo,EvaluationItem,ItemID",
                    "EvaluationItemDesc,Value,Result");

            string insertDetail = $@" ----寫入 StickerTest_Detail_Item
INSERT INTO StickerTest_Detail_Item
           (ReportNo,EvaluationItem,ItemID,EvaluationItemDesc,Value,Result,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@ItemID,@EvaluationItemDesc,@Value,@Result,@UserID,GETDATE())
";
            string updateDetail = $@" ----更新 StickerTest_Detail_Item，只有Value、Result是允許User 變更的
UPDATE StickerTest_Detail_Item
SET EditDate = GETDATE() , EditName = @UserID
    ,Value = @Value
    ,Result = @Result
WHERE ReportNo = @ReportNo
AND EvaluationItem = @EvaluationItem
AND ItemID = @ItemID
;
";
            string deleteDetail = $@" ----刪除 StickerTest_Detail_Item
DELETE FROM StickerTest_Detail_Item WHERE ReportNo = @ReportNo AND EvaluationItem = @EvaluationItem AND ItemID = @ItemID
";
            int idx = 0;
            foreach (var detailItem in needUpdateDetailItemList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@ItemID", $"{sources.Main.ReportNo}_{idx}"));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItemDesc", detailItem.EvaluationItemDesc));
                        listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        idx++;
                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@ItemID", detailItem.ItemID));
                        listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case CompareStateType.Delete:
                        listDetailPar.Add(new SqlParameter($"@ItemID", detailItem.ItemID));
                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case CompareStateType.None:

                        break;
                    default:
                        break;
                }

            }

            return true;

        }

        /// <summary>
        /// Encode / Amend StickerTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_StickerTest(StickerTest_Main request, string UserID)
        {

            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@EditName", UserID);
            paras.Add("@Status", request.Status);
            paras.Add("@Result", request.Result);
            paras.Add("@ReportNo", request.ReportNo);


            string sqlCmd;

            if (request.Status == "Confirmed")
            {
                sqlCmd = $@"
UPDATE StickerTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
;
";
            }
            else
            {
                sqlCmd = $@"
UPDATE StickerTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
;
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(StickerTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from StickerTest a
left join Pass1 mp on mp.ID = a.EditName 
left join MainServer.Production.dbo.Pass1 pp on pp.ID = a.EditName 
left join MainServer.Production.dbo.Technician t on t.ID = a.EditName 
where a.ReportNo = @ReportNo
;

";
            return ExecuteDataTable(CommandType.Text, sqlCmd, paras);
        }
    }
}
