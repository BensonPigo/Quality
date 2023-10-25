using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.Public;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolKit;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class WickingHeightTestProvider : SQLDAL
    {
        #region 底層連線
        public WickingHeightTestProvider(string ConString) : base(ConString) { }
        public WickingHeightTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(WickingHeightTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();

            string SbSql = $@"
select o.*, oa.Article
    , [FabricType] = case when s.FabricType = 'I' then 'KNIT' when s.FabricType = 'V' then 'WOVEN' else s.FabricType end        
from Orders o
inner join Order_Article oa on oa.id = o.ID
inner join Style s on o.StyleUkey = s.Ukey
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

        public List<WickingHeightTest_Main> GetMainList(WickingHeightTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
        ,d.TestWarpPicture
        ,d.TestWeftPicture
from WickingHeightTest a
left join SciPMSFile_WickingHeightTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
where 1=1
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

            var tmp = ExecuteList<WickingHeightTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<WickingHeightTest_Main>();
        }

        public List<WickingHeightTest_Detail> GetDetailList(WickingHeightTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from WickingHeightTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            sqlcmd += "Order by Ukey, EvaluationType";

            var tmp = ExecuteList<WickingHeightTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<WickingHeightTest_Detail>();
        }

        public List<WickingHeightTest_Detail_Item> GetDetaiItemlList(WickingHeightTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select wdi.* , wd.EvaluationType
from WickingHeightTest_Detail_Item wdi WITH(NOLOCK)
inner join WickingHeightTest_Detail wd WITH(NOLOCK) on wdi.WickingHeightTestDetailUkey = wd.Ukey
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and wdi.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            sqlcmd += "Order by wdi.WickingHeightTestDetailUkey, wdi.Ukey, wd.EvaluationType, wdi.EvaluationItem";

            var tmp = ExecuteList<WickingHeightTest_Detail_Item>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<WickingHeightTest_Detail_Item>();
        }

        public int Insert_WickingHeightTest(WickingHeightTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "WK", "WickingHeightTest", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@BrandID", DbType.String, Req.Main.BrandID ?? ""} ,
                { "@SeasonID", DbType.String, Req.Main.SeasonID ?? ""} ,
                { "@StyleID", DbType.String, Req.Main.StyleID ?? ""} ,
                { "@Article", DbType.String, Req.Main.Article ?? ""} ,
                { "@OrderID", DbType.String, Req.Main.OrderID ?? ""} ,
                { "@FactoryID", DbType.String, Req.Main.FactoryID ?? ""} ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@FabricType", DbType.String, Req.Main.FabricType ?? "" } ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestWarpPicture != null)
            {
                objParameter.Add("@TestWarpPicture", Req.Main.TestWarpPicture);
            }
            else
            {
                objParameter.Add("@TestWarpPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestWeftPicture != null)
            {
                objParameter.Add("@TestWeftPicture", Req.Main.TestWeftPicture);
            }
            else
            {
                objParameter.Add("@TestWeftPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"
SET XACT_ABORT ON

INSERT INTO dbo.WickingHeightTest
           (ReportNo
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,OrderID
           ,FactoryID
           ,SubmitDate
           ,FabricType
           ,Seq1
           ,Seq2
           ,FabricRefNo
           ,FabricColor
           ,FabricDescription
           ,Status
           ,Result
           ,AddDate
           ,AddName)
VALUES
           (@ReportNo
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,@Article
           ,@OrderID
           ,(select top 1 FactoryID from SciProduction_Orders with(NOLOCK) where ID = @OrderID)
           ,@SubmitDate
           ,@FabricType
           ,@Seq1
           ,@Seq2
           ,@FabricRefNo
           ,@FabricColor
           ,@FabricDescription
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
)
;

INSERT INTO SciPMSFile_WickingHeightTest
    ( ReportNo ,TestWarpPicture ,TestWeftPicture)
VALUES
    ( @ReportNo ,@TestWarpPicture ,@TestWeftPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update_WickingHeightTest(WickingHeightTest_ViewModel Req, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestWarpPicture != null)
            {
                objParameter.Add("@TestWarpPicture", Req.Main.TestWarpPicture);
            }
            else
            {
                objParameter.Add("@TestWarpPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestWeftPicture != null)
            {
                objParameter.Add("@TestWeftPicture", Req.Main.TestWeftPicture);
            }
            else
            {
                objParameter.Add("@TestWeftPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            string mainSqlCmd = $@"
SET XACT_ABORT ON

UPDATE WickingHeightTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricDescription = @FabricDescription
WHERE ReportNo = @ReportNo
;
UPDATE SciPMSFile_WickingHeightTest
SET TestWeftPicture = @TestWeftPicture , TestWarpPicture=@TestWarpPicture
WHERE ReportNo = @ReportNo
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }

        public int Delete_WickingHeightTest(WickingHeightTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
SET XACT_ABORT ON

delete from WickingHeightTest WHERE ReportNo = @ReportNo
delete from WickingHeightTest_Detail WHERE ReportNo = @ReportNo
delete from WickingHeightTest_Detail_Item WHERE ReportNo = @ReportNo
delete from SciPMSFile_WickingHeightTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }

        public bool Processe_WickingHeightTest_Detail(WickingHeightTest_ViewModel sources, string UserID)
        {
            List<WickingHeightTest_Detail> oldDetailData = this.GetDetailList(new WickingHeightTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            List<WickingHeightTest_Detail_Item> oldDetailItemData = this.GetDetaiItemlList(new WickingHeightTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            List<WickingHeightTest_Detail_Item> nowDetailItemData = sources.DetaiItemlList;
            List<WickingHeightTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<WickingHeightTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "ReportNo,Ukey",
                    "EvaluationType,WarpAverage,WarpResult,WeftAverage,WeftResult,Remark");
            List<WickingHeightTest_Detail_Item> needUpdateDetailItemList =
                PublicClass.CompareListValue<WickingHeightTest_Detail_Item>(
                    sources.DetaiItemlList,
                    oldDetailItemData,
                    "ReportNo,Ukey",
                    "WickingHeightTestDetailUkey,EvaluationItem,WarpValues,WarpTime,WeftValues,WeftTime");

            int i;
            string insertDetail = $@" ----寫入 WickingHeightTest_Detail

DECLARE @InsertOutput TABLE (ReportNo varchar(14), Ukey bigint, EvaluationType varchar(20))

INSERT INTO WickingHeightTest_Detail (ReportNo,EvaluationType,WarpAverage,WarpResult,WeftAverage,WeftResult,Remark,EditName,EditDate)
OUTPUT INSERTED.ReportNo, INSERTED.Ukey, INSERTED.EvaluationType INTO @InsertOutput
VALUES (@ReportNo,@EvaluationType,@WarpAverage,@WarpResult,@WeftAverage,@WeftResult,@Remark,@UserID,GETDATE())
";

            string updateDetail = $@" ----更新 WickingHeightTest_Detail

UPDATE WickingHeightTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationType = @EvaluationType
    ,WarpAverage = @WarpAverage
    ,WarpResult = @WarpResult
    ,WeftAverage = @WeftAverage
    ,WeftResult = @WeftResult
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey;
";

            string deleteDetail = $@" ----刪除 WickingHeightTest_Detail
DELETE FROM WickingHeightTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
DELETE FROM WickingHeightTest_Detail_Item where ReportNo = @ReportNo and WickingHeightTestDetailUkey = @Ukey
";

            
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@EvaluationType", detailItem.EvaluationType));
                        listDetailPar.Add(new SqlParameter($"@WarpAverage", detailItem.WarpAverage ?? 0) { Scale = 2 });
                        listDetailPar.Add(new SqlParameter($"@WarpResult", detailItem.WarpResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WeftAverage", detailItem.WeftAverage ?? 0) { Scale = 2 });
                        listDetailPar.Add(new SqlParameter($"@WeftResult", detailItem.WeftResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));

                        i = 1;
                        string insertDetailItem = string.Empty;
                        foreach (var item in nowDetailItemData.Where(x => x.EvaluationType == detailItem.EvaluationType))
                        {
                            insertDetailItem += $@"
INSERT INTO　WickingHeightTest_Detail_Item (ReportNo, WickingHeightTestDetailUkey, EvaluationItem, WarpValues, WarpTime, WeftValues, WeftTime, EditName, EditDate)
select t.ReportNo, [WickingHeightTestDetailUkey] = t.Ukey, @EvaluationItem{i}, @WarpValues{i}, @WarpTime{i}, @WeftValues{i}, @WeftTime{i}, @UserID, GETDATE()
from @InsertOutput t
";
                            listDetailPar.Add(new SqlParameter($"@EvaluationItem{i}", item.EvaluationItem));
                            listDetailPar.Add(new SqlParameter($"@WarpValues{i}", item.WarpValues ?? 0) { Scale = 2 });
                            listDetailPar.Add(new SqlParameter($"@WarpTime{i}", item.WarpTime ?? 0));
                            listDetailPar.Add(new SqlParameter($"@WeftValues{i}", item.WeftValues ?? 0) { Scale = 2 });
                            listDetailPar.Add(new SqlParameter($"@WeftTime{i}", item.WeftTime ?? 0));
                            i++;
                        }                       

                        ExecuteNonQuery(CommandType.Text, insertDetail + insertDetailItem, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationType", detailItem.EvaluationType));
                        listDetailPar.Add(new SqlParameter($"@WarpAverage", detailItem.WarpAverage ?? 0) { Scale = 2 });
                        listDetailPar.Add(new SqlParameter($"@WarpResult", detailItem.WarpResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WeftAverage", detailItem.WeftAverage ?? 0) { Scale = 2 });
                        listDetailPar.Add(new SqlParameter($"@WeftResult", detailItem.WeftResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));

                        i = 1;
                        string updateDetaillItem = string.Empty;
                        foreach (var item in nowDetailItemData.Where(x => x.EvaluationType == detailItem.EvaluationType))
                        {
                            updateDetaillItem += $@"
UPDATE WickingHeightTest_Detail_Item
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem{i}
    ,WarpValues = @WarpValues{i}
    ,WarpTime = @WarpTime{i}
    ,WeftValues = @WeftValues{i}
    ,WeftTime = @WeftTime{i}
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey{i}
AND WickingHeightTestDetailUkey = @WickingHeightTestDetailUkey{i}
";

                            listDetailPar.Add(new SqlParameter($"@WickingHeightTestDetailUkey{i}", detailItem.Ukey));
                            listDetailPar.Add(new SqlParameter($"@Ukey{i}", item.Ukey));
                            listDetailPar.Add(new SqlParameter($"@EvaluationItem{i}", item.EvaluationItem));
                            listDetailPar.Add(new SqlParameter($"@WarpValues{i}", item.WarpValues ?? 0) { Scale = 2 });
                            listDetailPar.Add(new SqlParameter($"@WarpTime{i}", item.WarpTime ?? 0));
                            listDetailPar.Add(new SqlParameter($"@WeftValues{i}", item.WeftValues ?? 0) { Scale = 2 });
                            listDetailPar.Add(new SqlParameter($"@WeftTime{i}", item.WeftTime ?? 0));
                            i++;
                        }
                        
                        ExecuteNonQuery(CommandType.Text, updateDetail += updateDetaillItem, listDetailPar);
                        break;
                    case CompareStateType.Delete:
                        listDetailPar.Add(new SqlParameter($"@ReportNo", detailItem.ReportNo));
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));

                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case CompareStateType.None:

                        break;
                    default:
                        break;
                }

            }


            updateDetail = @"
UPDATE WickingHeightTest_Detail_Item
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,WarpValues = @WarpValues
    ,WarpTime = @WarpTime
    ,WeftValues = @WeftValues
    ,WeftTime = @WeftTime
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
";
            // 只有第三層 update
            foreach (var item in needUpdateDetailItemList.Where(x => x.StateType == CompareStateType.Edit))
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@Ukey", item.Ukey));
                listDetailPar.Add(new SqlParameter($"@EvaluationItem", item.EvaluationItem));
                listDetailPar.Add(new SqlParameter($"@WarpValues", item.WarpValues ?? 0) { Scale = 2 });
                listDetailPar.Add(new SqlParameter($"@WarpTime", item.WarpTime ?? 0));
                listDetailPar.Add(new SqlParameter($"@WeftValues", item.WeftValues ?? 0) { Scale = 2 });
                listDetailPar.Add(new SqlParameter($"@WeftTime", item.WeftTime ?? 0));
                listDetailPar.Add(new SqlParameter($"@UserID", UserID));      

                ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
            }

            return true;

        }

        /// <summary>
        /// Encode / Amend WickingHeightTest
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_WickingHeightTest(WickingHeightTest_Main request, string UserID)
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
UPDATE WickingHeightTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE WickingHeightTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(WickingHeightTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection
            {
                { "@ReportNo", Req.ReportNo }
            };

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from WickingHeightTest a
left join Pass1 mp on mp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Pass1 pp on pp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Technician t on t.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
where a.ReportNo = @ReportNo
;

";
            return ExecuteDataTable(CommandType.Text, sqlCmd, paras);
        }
    }
}
