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

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class HydrostaticPressureWaterproofTestProvider : SQLDAL
    {
        #region 底層連線
        public HydrostaticPressureWaterproofTestProvider(string ConString) : base(ConString) { }
        public HydrostaticPressureWaterproofTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<HydrostaticPressureWaterproofStandard> GetStandards()
        {
            string sqlcmd = $@"select * from HydrostaticPressureWaterproofStandard";
            var tmp = ExecuteList<HydrostaticPressureWaterproofStandard>(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return tmp.Any() ? tmp.ToList() : new List<HydrostaticPressureWaterproofStandard>();
        }

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(HydrostaticPressureWaterproofTest_Request Req)
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
        public List<HydrostaticPressureWaterproofTest_Main> GetMainList(HydrostaticPressureWaterproofTest_Request Req)
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
from HydrostaticPressureWaterproofTest a
left join PMSFile.dbo.HydrostaticPressureWaterproofTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<HydrostaticPressureWaterproofTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<HydrostaticPressureWaterproofTest_Main>();
        }
        public List<HydrostaticPressureWaterproofTest_Detail> GetDetailList(HydrostaticPressureWaterproofTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select distinct a.*
    ,b.EvaluationTypeSeq
    ,b.EvaluationItemSeq
from HydrostaticPressureWaterproofTest_Detail a WITH(NOLOCK)
inner join HydrostaticPressureWaterproofStandard b on a.EvaluationType=b.EvaluationType and a.EvaluationItem=b.EvaluationItem 
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
            var tmp = ExecuteList<HydrostaticPressureWaterproofTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<HydrostaticPressureWaterproofTest_Detail>();
        }
        public int Insert_HydrostaticPressureWaterproofTest(HydrostaticPressureWaterproofTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "HW", "HydrostaticPressureWaterproofTest", DateTime.Today, 2, "ReportNo");
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
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Main.Temperature } ,
                { "@DryingCondition", DbType.String, Req.Main.DryingCondition ?? "" } ,
                { "@WashCycles", DbType.Int32, Req.Main.WashCycles } ,
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
INSERT INTO dbo.HydrostaticPressureWaterproofTest
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
           ,FabricRefNo
           ,FabricColor
           ,FabricDescription
           ,Temperature
           ,DryingCondition
           ,WashCycles
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
           ,@FabricRefNo
           ,@FabricColor
           ,@FabricDescription
           ,@Temperature
           ,@DryingCondition
           ,@WashCycles
           ,@Result
           ,'New'
           ,GETDATE()
           ,@AddName)
;

INSERT INTO PMSFile.dbo.HydrostaticPressureWaterproofTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_HydrostaticPressureWaterproofTest(HydrostaticPressureWaterproofTest_ViewModel Req, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
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
UPDATE HydrostaticPressureWaterproofTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,Remark = @Remark
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricDescription = @FabricDescription
WHERE ReportNo = @ReportNo
;
UPDATE PMSFile.dbo.HydrostaticPressureWaterproofTest
SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture
WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public string Delete_HydrostaticPressureWaterproofTest(HydrostaticPressureWaterproofTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@EditName", DbType.String, Req.Main.EditName } ,
            };

            string mainSqlCmd = $@"
delete from HydrostaticPressureWaterproofTest WHERE ReportNo = @ReportNo
delete from HydrostaticPressureWaterproofTest WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.HydrostaticPressureWaterproofTest WHERE ReportNo = @ReportNo

";

            return ExecuteScalar(CommandType.Text, mainSqlCmd.ToString(), objParameter).ToString();
        }
        public bool Processe_HydrostaticPressureWaterproofTest_Detail(HydrostaticPressureWaterproofTest_ViewModel sources, string UserID)
        {
            List<HydrostaticPressureWaterproofTest_Detail> oldDetailData = this.GetDetailList(new HydrostaticPressureWaterproofTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<HydrostaticPressureWaterproofTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<HydrostaticPressureWaterproofTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "EvaluationType,EvaluationItem,AsReceivedValue,AsReceivedResult,AfterWashValue,AfterWashResult,Remark");

            string insertDetail = $@" ----寫入 HydrostaticPressureWaterproofTest_Detail
INSERT INTO HydrostaticPressureWaterproofTest_Detail
           (ReportNo,EvaluationType,EvaluationItem,AsReceivedValue,AsReceivedResult,AfterWashValue,AfterWashResult,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationType,@EvaluationItem,@AsReceivedValue,@AsReceivedResult,@AfterWashValue,@AfterWashResult,@Remark,@UserID,GETDATE())
";
            string updateDetail = $@" ----更新 HydrostaticPressureWaterproofTest_Detail
UPDATE HydrostaticPressureWaterproofTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationType = @EvaluationType
    ,EvaluationItem = @EvaluationItem
    ,AsReceivedValue = @AsReceivedValue
    ,AsReceivedResult = @AsReceivedResult
    ,AfterWashValue = @AfterWashValue
    ,AfterWashResult = @AfterWashResult
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 HydrostaticPressureWaterproofTest_Detail
DELETE FROM HydrostaticPressureWaterproofTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
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
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@AsReceivedValue", detailItem.AsReceivedValue));
                        listDetailPar.Add(new SqlParameter($"@AsReceivedResult", detailItem.AsReceivedResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AfterWashValue", detailItem.AfterWashValue));
                        listDetailPar.Add(new SqlParameter($"@AfterWashResult", detailItem.AfterWashResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationType", detailItem.EvaluationType));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@AsReceivedValue", detailItem.AsReceivedValue));
                        listDetailPar.Add(new SqlParameter($"@AsReceivedResult", detailItem.AsReceivedResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AfterWashValue", detailItem.AfterWashValue));
                        listDetailPar.Add(new SqlParameter($"@AfterWashResult", detailItem.AfterWashResult ?? string.Empty));
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
        /// Encode / Amend HydrostaticPressureWaterproofTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_HydrostaticPressureWaterproofTest(HydrostaticPressureWaterproofTest_Main request, string UserID)
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
UPDATE HydrostaticPressureWaterproofTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , Result = @Result
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
;
";
            }
            else
            {
                sqlCmd = $@"
UPDATE HydrostaticPressureWaterproofTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , Result = ''
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
;
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(HydrostaticPressureWaterproofTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from HydrostaticPressureWaterproofTest a
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
