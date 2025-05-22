using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Linq;
using System.Security.Cryptography;
using System.Data.SqlTypes;
using ToolKit;
using DatabaseObject.Public;
using System.Data.SqlClient;
using DatabaseObject.ResultModel;
using System.Web.Mvc;
using Microsoft.SqlServer.Server;
using System.Reflection;
using DatabaseObject.ResultModel.EtoEFlowChart;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class PhenolicYellowTestProvider : SQLDAL
    {
        public string GetFactoryNameEN(string ReportNo, string FactoryID)
        {
            string factoryNameEN = string.Empty;
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ReportNo", DbType.String, ReportNo } ,
                { "@FactoryID",DbType.String, FactoryID } ,
            };
            string sql = $@"
            SELECT
            o.FactoryID
            INTO #tmp
            FROM PhenolicYellowTest P WITH(NOLOCK)
            INNER JOIN Production.dbo.Orders O WITH(NOLOCK) ON O.ID = P.OrderID
            WHERE P.ReportNo = @ReportNo
			
            SELECT
            F.NameEN,*
            FROM Production.dbo.Factory F WITH(NOLOCK)
            LEFT JOIN #TMP T WITH(NOLOCK) ON T.FactoryID = F.ID
            WHERE F.ID = IIF((SELECT count(1) from #tmp) > 0 ,T.FactoryID,@FactoryID)";

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sql, objParameter);
            factoryNameEN = dt.Rows[0]["NameEN"].ToString();
            return factoryNameEN;
        }

        #region 底層連線
        public PhenolicYellowTestProvider(string ConString) : base(ConString) { }
        public PhenolicYellowTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

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

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(PhenolicYellowTest_Request Req)
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

        public List<PhenolicYellowTest_Main> GetMainList(PhenolicYellowTest_Request Req)
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
,[ApproverName] = (select Name from pass1 where id = a.Approver)
,[PreparerName] = (select Name from pass1 where id = a.Preparer)
from PhenolicYellowTest a
left join SciPMSFile_PhenolicYellowTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<PhenolicYellowTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<PhenolicYellowTest_Main>();
        }

        public List<PhenolicYellowTest_Detail> GetDetailList(PhenolicYellowTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from PhenolicYellowTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<PhenolicYellowTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<PhenolicYellowTest_Detail>();
        }

        public int Insert_PhenolicYellowTest(PhenolicYellowTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "PY", "PhenolicYellowTest", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
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
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@Temperature", Req.Main.Temperature } ,
                { "@Time", Req.Main.Time } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
                { "@ReportDate" ,Req.Main.ReportDate } ,
                { "@Approver",DbType.String, Req.Main.Approver},
                { "@Preparer",DbType.String, Req.Main.Preparer},
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
SET XACT_ABORT ON

INSERT INTO dbo.PhenolicYellowTest
           (ReportNo
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,OrderID
           ,SubmitDate
           ,Seq1
           ,Seq2
           ,FabricRefNo
           ,FabricColor
           ,Temperature
           ,Time
           ,Status
           ,Result
           ,AddDate
           ,AddName
           ,ReportDate
           ,Approver
           ,Preparer
)
VALUES
           (@ReportNo
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,@Article
           ,@OrderID
           ,@SubmitDate
           ,@Seq1
           ,@Seq2
           ,@FabricRefNo
           ,@FabricColor
           ,@Temperature
           ,@Time
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
           ,@ReportDate
           ,@Approver
           ,@Preparer
)
;

INSERT INTO SciPMSFile_PhenolicYellowTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_PhenolicYellowTest(PhenolicYellowTest_ViewModel Req, string UserID)
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
                { "@Temperature", Req.Main.Temperature } ,
                { "@Time", Req.Main.Time } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
                { "@ReportDate",DbType.Date, Req.Main.ReportDate},
                { "@Approver",DbType.String, Req.Main.Approver},
                { "@Preparer",DbType.String, Req.Main.Preparer},
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
SET XACT_ABORT ON

UPDATE PhenolicYellowTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,Temperature = @Temperature
      ,Time = @Time
      ,ReportDate = @ReportDate
      ,Approver = @Approver
      ,Preparer = @Preparer
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.PhenolicYellowTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.PhenolicYellowTest
    SET TestBeforePicture = @TestBeforePicture , TestAfterPicture=@TestAfterPicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.PhenolicYellowTest
        ( ReportNo ,TestBeforePicture ,TestAfterPicture )
    VALUES
        ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture )
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_PhenolicYellowTest(PhenolicYellowTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
SET XACT_ABORT ON

delete from PhenolicYellowTest WHERE ReportNo = @ReportNo
delete from PhenolicYellowTest_Detail WHERE ReportNo = @ReportNo
delete from SciPMSFile_PhenolicYellowTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_PhenolicYellowTest_Detail(PhenolicYellowTest_ViewModel sources, string UserID)
        {
            List<PhenolicYellowTest_Detail> oldDetailData = this.GetDetailList(new PhenolicYellowTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<PhenolicYellowTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<PhenolicYellowTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "EvaluationItem,Roll,Dyelot,Scale,Result,Remark");

            string insertDetail = $@" ----寫入 PhenolicYellowTest_Detail

INSERT INTO PhenolicYellowTest_Detail
           (ReportNo,EvaluationItem,Roll,Dyelot,Scale,Result,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@Roll,@Dyelot,@Scale,@Result,@Remark,@UserID,GETDATE())

";
            string updateDetail = $@" ----更新 PhenolicYellowTest_Detail


UPDATE PhenolicYellowTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,Roll = @Roll
    ,Dyelot = @Dyelot
    ,Scale = @Scale
    ,Result = @Result
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 PhenolicYellowTest_Detail
DELETE FROM PhenolicYellowTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
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
                        listDetailPar.Add(new SqlParameter($"@Roll", detailItem.Roll ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Dyelot", detailItem.Dyelot ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Scale", detailItem.Scale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@Roll", detailItem.Roll ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Dyelot", detailItem.Dyelot ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Scale", detailItem.Scale ?? string.Empty));
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
        /// Encode / Amend PhenolicYellowTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_PhenolicYellowTest(PhenolicYellowTest_Main request, string UserID)
        {

            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@EditName", UserID);
            paras.Add("@Status", request.Status);
            paras.Add("@Result", request.Result);
            paras.Add("@ReportNo", request.ReportNo);


            string sqlCmd ;

            if (request.Status == "Confirmed")
            {
                sqlCmd = $@"
UPDATE PhenolicYellowTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE PhenolicYellowTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(PhenolicYellowTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from PhenolicYellowTest a
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
