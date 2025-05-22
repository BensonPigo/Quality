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
    public class MartindalePillingTestProvider : SQLDAL
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
            FROM MartindalePillingTest P WITH(NOLOCK)
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
        public MartindalePillingTestProvider(string ConString) : base(ConString) { }
        public MartindalePillingTestProvider(SQLDataTransaction tra) : base(tra) { }
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

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(MartindalePillingTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();

            string SbSql = $@"

select o.*, oa.Article 
,FabricType = CASE WHEN s.FabricType = 'I' THEN 'KNIT'
                   WHEN s.FabricType = 'V' THEN 'WOVEN'
                   ELSE s.FabricType
              END
from Orders o
inner join Order_Article oa on oa.id = o.ID
inner join Style s on s.Ukey = o.StyleUkey
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

        public List<MartindalePillingTest_Main> GetMainList(MartindalePillingTest_Request Req)
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
        ,d.TestBeforePicture
        ,d.Test500AfterPicture
        ,d.Test2000AfterPicture
,[ApproverName] = (select Name from [MainServer].Production.dbo.pass1 where id = a.Approver)
,[PreparerName] = (select Name from [MainServer].Production.dbo.pass1 where id = a.Preparer)
from MartindalePillingTest a
left join PMSFile.dbo.MartindalePillingTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<MartindalePillingTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<MartindalePillingTest_Main>();
        }

        public List<MartindalePillingTest_Detail> GetDetailList(MartindalePillingTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from MartindalePillingTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<MartindalePillingTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<MartindalePillingTest_Detail>();
        }

        public int Insert_MartindalePillingTest(MartindalePillingTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "MP", "MartindalePillingTest", DateTime.Today, 2, "ReportNo");
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
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
                { "@Approver",DbType.String,Req.Main.Approver ?? ""},
                { "@Preparer",DbType.String,Req.Main.Preparer ?? ""},
            };

            if (Req.Main.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.Test500AfterPicture != null)
            {
                objParameter.Add("@Test500AfterPicture", Req.Main.Test500AfterPicture);
            }
            else
            {
                objParameter.Add("@Test500AfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.Test2000AfterPicture != null)
            {
                objParameter.Add("@Test2000AfterPicture", Req.Main.Test2000AfterPicture);
            }
            else
            {
                objParameter.Add("@Test2000AfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"

INSERT INTO dbo.MartindalePillingTest
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
           ,TestStandard
           ,Status
           ,Result
           ,AddDate
           ,AddName
           ,Approver
           ,Preparer)
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
           ,@TestStandard
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
           ,@Approver
           ,@Preparer
)
;

INSERT INTO PMSFile.dbo.MartindalePillingTest
    ( ReportNo ,TestBeforePicture ,Test500AfterPicture ,Test2000AfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@Test500AfterPicture ,@Test2000AfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_MartindalePillingTest(MartindalePillingTest_ViewModel Req, string UserID)
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
                { "@FabricType", DbType.String, Req.Main.FabricType ?? "" } ,
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
                { "@ReportDate",DbType.Date,Req.Main.ReportDate },
                { "@Approver",DbType.String,Req.Main.Approver ?? ""},
                { "@Preparer",DbType.String,Req.Main.Preparer ?? ""},
            };

            if (Req.Main.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.Test500AfterPicture != null)
            {
                objParameter.Add("@Test500AfterPicture", Req.Main.Test500AfterPicture);
            }
            else
            {
                objParameter.Add("@Test500AfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.Test2000AfterPicture != null)
            {
                objParameter.Add("@Test2000AfterPicture", Req.Main.Test2000AfterPicture);
            }
            else
            {
                objParameter.Add("@Test2000AfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            string mainSqlCmd = $@"

UPDATE MartindalePillingTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricType = @FabricType
      ,TestStandard = @TestStandard
    ,ReportDate = @ReportDate
    ,Approver=@Approver
    ,Preparer=@Preparer
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.MartindalePillingTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.MartindalePillingTest
    SET TestBeforePicture = @TestBeforePicture , Test500AfterPicture=@Test500AfterPicture , Test2000AfterPicture=@Test2000AfterPicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.MartindalePillingTest
        ( ReportNo ,TestBeforePicture ,Test500AfterPicture ,Test2000AfterPicture)
    VALUES
        ( @ReportNo ,@TestBeforePicture ,@Test500AfterPicture ,@Test2000AfterPicture)
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_MartindalePillingTest(MartindalePillingTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"

delete from MartindalePillingTest WHERE ReportNo = @ReportNo
delete from MartindalePillingTest_Detail WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.MartindalePillingTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_MartindalePillingTest_Detail(MartindalePillingTest_ViewModel sources, string UserID)
        {
            List<MartindalePillingTest_Detail> oldDetailData = this.GetDetailList(new MartindalePillingTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<MartindalePillingTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<MartindalePillingTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "RubTimes,EvaluationItem,Scale,Result,Remark");

            string insertDetail = $@" ----寫入 MartindalePillingTest_Detail

INSERT INTO MartindalePillingTest_Detail
           (ReportNo,Result,RubTimes,EvaluationItem,Scale,Remark,EditDate,EditName)
VALUES 
           (@ReportNo,@Result,@RubTimes,@EvaluationItem,@Scale,@Remark,GETDATE(),@UserID)

";
            string updateDetail = $@" ----更新 MartindalePillingTest_Detail


UPDATE MartindalePillingTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,Result = @Result
    ,Remark = @Remark
    ,RubTimes = @RubTimes
    ,Scale = @Scale
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 MartindalePillingTest_Detail
DELETE FROM MartindalePillingTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
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
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Scale", detailItem.Scale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@RubTimes", detailItem.RubTimes ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Scale", detailItem.Scale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@RubTimes", detailItem.RubTimes ?? string.Empty));

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
        /// Encode / Amend MartindalePillingTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_MartindalePillingTest(MartindalePillingTest_Main request, string UserID)
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
UPDATE MartindalePillingTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE MartindalePillingTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(MartindalePillingTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from MartindalePillingTest a
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
