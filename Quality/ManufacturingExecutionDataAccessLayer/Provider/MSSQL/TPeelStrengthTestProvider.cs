using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
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
    public class TPeelStrengthTestProvider : SQLDAL
    {
        #region 底層連線
        public TPeelStrengthTestProvider(string ConString) : base(ConString) { }
        public TPeelStrengthTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(TPeelStrengthTest_Request Req)
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

        public List<TPeelStrengthTest_Main> GetMainList(TPeelStrengthTest_Request Req)
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
from TPeelStrengthTest a
left join SciPMSFile_TPeelStrengthTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<TPeelStrengthTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<TPeelStrengthTest_Main>();
        }

        public List<TPeelStrengthTest_Detail> GetDetailList(TPeelStrengthTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from TPeelStrengthTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<TPeelStrengthTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<TPeelStrengthTest_Detail>();
        }

        public int Insert_TPeelStrengthTest(TPeelStrengthTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "TS", "TPeelStrengthTest", DateTime.Today, 2, "ReportNo");
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
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@MachineNo", DbType.String, Req.Main.MachineNo ?? "" } ,
                { "@MachineReport", DbType.String, Req.Main.MachineReport ?? "" } ,
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

INSERT INTO dbo.TPeelStrengthTest
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
           ,FabricDescription
           ,MachineNo
           ,MachineReport
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
           ,@FabricDescription
           ,@MachineNo
           ,@MachineReport
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
)
;

INSERT INTO PMSFile.dbo.TPeelStrengthTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update_TPeelStrengthTest(TPeelStrengthTest_ViewModel Req, string UserID)
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
                { "@MachineNo", DbType.String, Req.Main.MachineNo ?? "" } ,
                { "@MachineReport", DbType.String, Req.Main.MachineReport ?? "" } ,

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

UPDATE TPeelStrengthTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricDescription = @FabricDescription
      ,MachineNo = @MachineNo
      ,MachineReport = @MachineReport
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.TPeelStrengthTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.TPeelStrengthTest
    SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.TPeelStrengthTest
        ( ReportNo ,TestAfterPicture ,TestBeforePicture )
    VALUES
        ( @ReportNo ,@TestAfterPicture ,@TestBeforePicture )
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_TPeelStrengthTest(TPeelStrengthTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
delete from TPeelStrengthTest WHERE ReportNo = @ReportNo
delete from TPeelStrengthTest_Detail WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.TPeelStrengthTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_TPeelStrengthTest_Detail(TPeelStrengthTest_ViewModel sources, string UserID)
        {
            List<TPeelStrengthTest_Detail> oldDetailData = this.GetDetailList(new TPeelStrengthTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<TPeelStrengthTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<TPeelStrengthTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "EvaluationItem,AllResult,WarpValue,WarpResult,WeftValue,WeftResult,Remark");

            string insertDetail = $@" ----寫入 TPeelStrengthTest_Detail

INSERT INTO TPeelStrengthTest_Detail
           (ReportNo,EvaluationItem,AllResult,WarpValue,WarpResult,WeftValue,WeftResult,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@AllResult,@WarpValue,@WarpResult,@WeftValue,@WeftResult,@Remark,@UserID,GETDATE())

";
            string updateDetail = $@" ----更新 TPeelStrengthTest_Detail


UPDATE TPeelStrengthTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,AllResult = @AllResult
    ,Remark = @Remark
    ,WarpValue = @WarpValue
    ,WarpResult = @WarpResult
    ,WeftValue = @WeftValue
    ,WeftResult = @WeftResult
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 TPeelStrengthTest_Detail
DELETE FROM TPeelStrengthTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@AllResult", detailItem.AllResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WarpValue", detailItem.WarpValue));
                        listDetailPar.Add(new SqlParameter($"@WarpResult", detailItem.WarpResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WeftValue", detailItem.WeftValue));
                        listDetailPar.Add(new SqlParameter($"@WeftResult", detailItem.WeftResult ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@AllResult", detailItem.AllResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WarpValue", detailItem.WarpValue));
                        listDetailPar.Add(new SqlParameter($"@WarpResult", detailItem.WarpResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WeftValue", detailItem.WeftValue));
                        listDetailPar.Add(new SqlParameter($"@WeftResult", detailItem.WeftResult ?? string.Empty));

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
        /// Encode / Amend TPeelStrengthTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_TPeelStrengthTest(TPeelStrengthTest_Main request, string UserID)
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
UPDATE TPeelStrengthTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE TPeelStrengthTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(TPeelStrengthTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from TPeelStrengthTest a
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
