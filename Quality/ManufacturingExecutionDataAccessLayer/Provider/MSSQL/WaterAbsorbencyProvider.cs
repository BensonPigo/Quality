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
using ToolKit;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class WaterAbsorbencyProvider : SQLDAL
    {
        #region 底層連線
        public WaterAbsorbencyProvider(string ConString) : base(ConString) { }
        public WaterAbsorbencyProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(WaterAbsorbency_Request Req)
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

        public List<WaterAbsorbency_Main> GetMainList(WaterAbsorbency_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
        ,d.TestBeforePicture
        ,d.TestBeforeWashPicture
        ,d.TestAfterPicture
from WaterAbsorbencyTest a
left join SciPMSFile_WaterAbsorbencyTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<WaterAbsorbency_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<WaterAbsorbency_Main>();
        }

        public List<WaterAbsorbency_Detail> GetDetailList(WaterAbsorbency_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
from WaterAbsorbencyTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<WaterAbsorbency_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<WaterAbsorbency_Detail>();
        }

        public int Insert_WaterAbsorbency(WaterAbsorbency_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "WA", "WaterAbsorbencyTest", DateTime.Today, 2, "ReportNo");
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
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricType", DbType.String, Req.Main.FabricType ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
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

            if (Req.Main.TestBeforeWashPicture != null)
            {
                objParameter.Add("@TestBeforeWashPicture", Req.Main.TestBeforeWashPicture);
            }
            else
            {
                objParameter.Add("@TestBeforeWashPicture", System.Data.SqlTypes.SqlBinary.Null);
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
SET XACT_ABORT ON

INSERT INTO dbo.WaterAbsorbencyTest
           (ReportNo
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,OrderID
           ,FactoryID
           ,SubmitDate
           ,Seq1
           ,Seq2
           ,FabricRefNo
           ,FabricColor
           ,FabricType
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
           ,@Seq1
           ,@Seq2
           ,@FabricRefNo
           ,@FabricColor
           ,@FabricType
           ,@FabricDescription
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
)
;

INSERT INTO SciPMSFile_WaterAbsorbencyTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture, TestBeforeWashPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture, @TestBeforeWashPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_WaterAbsorbency(WaterAbsorbency_ViewModel Req, string UserID)
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

            if (Req.Main.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestBeforeWashPicture != null)
            {
                objParameter.Add("@TestBeforeWashPicture", Req.Main.TestBeforeWashPicture);
            }
            else
            {
                objParameter.Add("@TestBeforeWashPicture", System.Data.SqlTypes.SqlBinary.Null);
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
SET XACT_ABORT ON

UPDATE WaterAbsorbencyTest
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
if exists(select 1 from PMSFile.dbo.WaterAbsorbencyTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.WaterAbsorbencyTest
    SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture , TestBeforeWashPicture=@TestBeforeWashPicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.WaterAbsorbencyTest
        ( ReportNo ,TestAfterPicture ,TestBeforePicture ,TestBeforeWashPicture)
    VALUES
        ( @ReportNo ,@TestAfterPicture ,@TestBeforePicture ,@TestBeforeWashPicture)
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_WaterAbsorbency(WaterAbsorbency_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
SET XACT_ABORT ON

delete from WaterAbsorbencyTest WHERE ReportNo = @ReportNo
delete from WaterAbsorbencyTest_Detail WHERE ReportNo = @ReportNo
delete from SciPMSFile_WaterAbsorbencyTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_WaterAbsorbency_Detail(WaterAbsorbency_ViewModel sources, string UserID)
        {
            List<WaterAbsorbency_Detail> oldDetailData = this.GetDetailList(new WaterAbsorbency_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<WaterAbsorbency_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<WaterAbsorbency_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "EvaluationItem,NoOfDrops,Values,Result,Remark");

            string insertDetail = $@" ----寫入 WaterAbsorbencyTest_Detail

INSERT INTO WaterAbsorbencyTest_Detail
           (ReportNo,EvaluationItem,NoOfDrops,[Values],Result,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@NoOfDrops,@Values,@Result,@Remark,@UserID,GETDATE())

";
            string updateDetail = $@" ----更新 WaterAbsorbencyTest_Detail


UPDATE WaterAbsorbencyTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,NoOfDrops = @NoOfDrops
    ,[Values] = @Values
    ,Result = @Result
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 WaterAbsorbencyTest_Detail
DELETE FROM WaterAbsorbencyTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection
                {
                    new SqlParameter($"@ReportNo", sources.Main.ReportNo),
                    new SqlParameter($"@UserID", UserID)
                };
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@NoOfDrops", detailItem.NoOfDrops ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Values", detailItem.Values ?? 0));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@NoOfDrops", detailItem.NoOfDrops ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Values", detailItem.Values ?? 0));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
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

            return true;

        }

        /// <summary>
        /// Encode / Amend WaterAbsorbencyTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_WaterAbsorbency(WaterAbsorbency_Main request, string UserID)
        {

            SQLParameterCollection paras = new SQLParameterCollection
            {
                { "@EditName", UserID },
                { "@Status", request.Status },
                { "@Result", request.Result },
                { "@ReportNo", request.ReportNo }
            };


            string sqlCmd;

            if (request.Status == "Confirmed")
            {
                sqlCmd = $@"
UPDATE WaterAbsorbencyTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE WaterAbsorbencyTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(WaterAbsorbency_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection
            {
                { "@ReportNo", Req.ReportNo }
            };

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from WaterAbsorbencyTest a
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
