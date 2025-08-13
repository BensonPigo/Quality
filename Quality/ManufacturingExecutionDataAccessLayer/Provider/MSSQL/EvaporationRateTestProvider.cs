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
using System.Data.SqlTypes;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class EvaporationRateTestProvider : SQLDAL
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
            FROM EvaporationRateTest P WITH(NOLOCK)
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
        public EvaporationRateTestProvider(string ConString) : base(ConString) { }
        public EvaporationRateTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();

            string SbSql = $@"
select DISTINCT sa.Article, o.*
from Production.dbo.Style_Article sa WITH(NOLOCK)
inner join Production.dbo.Orders o WITH(NOLOCK) on o.StyleUkey = sa.StyleUkey
where Category ='B'
";

            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql += $@" AND o.ID = @OrderID ";
                paras.Add("@OrderID", DbType.String, Req.OrderID ?? "");
            }
            else if (!string.IsNullOrEmpty(Req.BrandID) && !string.IsNullOrEmpty(Req.SeasonID) && !string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += $@" AND o.StyleID = @StyleID AND o.BrandID = @BrandID AND o.SeasonID = @SeasonID )";
                paras.Add("@StyleID", DbType.String, Req.StyleID ?? "");
                paras.Add("@BrandID", DbType.String, Req.BrandID ?? "");
                paras.Add("@SeasonID", DbType.String, Req.SeasonID ?? "");
            }
            else
            {
                SbSql += $@"AND 1=0";
            }



            var tmp = ExecuteList<DatabaseObject.ProductionDB.Orders>(CommandType.Text, SbSql, paras);


            return tmp.Any() ? tmp.ToList() : new List<DatabaseObject.ProductionDB.Orders>();
        }
        public List<EvaporationRateTest_Main> GetMainList(EvaporationRateTest_Request Req)
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
		,ApproverName = ISNULL(p1.Name,p2.Name)
        ,d.TestAfterPicture
        ,d.TestBeforePicture
,[ApproverName] = (select Name from [MainServer].Production.dbo.pass1 where id = a.Approver)
,[PreparerName] = (select Name from [MainServer].Production.dbo.pass1 where id = a.Preparer)
from EvaporationRateTest a
left join PMSFile.dbo.EvaporationRateTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
left join ManufacturingExecution.dbo.Pass1 p1 on a.Approver = p1.ID
left join SciProduction_Pass1 p2 on a.Approver = p2.ID
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

            var tmp = ExecuteList<EvaporationRateTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Main>();
        }
        public List<EvaporationRateTest_Detail> GetDetailList(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select distinct a.*
from EvaporationRateTest_Detail a WITH(NOLOCK)
inner join EvaporationRateTest b on a.ReportNo=b.ReportNo 
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and b.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<EvaporationRateTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Detail>();
        }
        public List<EvaporationRateTest_Specimen> GetSpecimenList(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select c.*
,DetailType = b.Type
from EvaporationRateTest a WITH(NOLOCK)
inner join EvaporationRateTest_Detail b  WITH(NOLOCK) on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c WITH(NOLOCK) on b.Ukey = c.DetailUkey
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<EvaporationRateTest_Specimen>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Specimen>();
        }
        public List<EvaporationRateTest_Specimen_Time> GetTimeList(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select d.*
,DetailType = b.Type
,c.SpecimenID
,InitialTime = (SELECT top 1 Time from EvaporationRateTest_Specimen_Time e WITH(NOLOCK) where e.Ukey = d.InitialTimeUkey)
from EvaporationRateTest a WITH(NOLOCK)
inner join EvaporationRateTest_Detail b  WITH(NOLOCK) on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c WITH(NOLOCK) on c.DetailUkey = b.Ukey
inner join EvaporationRateTest_Specimen_Time d WITH(NOLOCK) on d.SpecimenUkey = c.Ukey
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<EvaporationRateTest_Specimen_Time>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Specimen_Time>();
        }
        public List<EvaporationRateTest_Specimen_Rate> GetRateList(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select d.*
,DetailType = b.Type
,c.SpecimenID
,Subtrahend_Time = (select top 1 Time from EvaporationRateTest_Specimen_Time e where e.Ukey=d.Subtrahend_TimeUkey)
,Minuend_Time = (select top 1 Time from EvaporationRateTest_Specimen_Time e where e.Ukey=d.Minuend_TimeUkey)
from EvaporationRateTest a WITH(NOLOCK)
inner join EvaporationRateTest_Detail b  WITH(NOLOCK) on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c WITH(NOLOCK) on c.DetailUkey = b.Ukey
inner join EvaporationRateTest_Specimen_Rate d WITH(NOLOCK) on d.SpecimenUkey = c.Ukey
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<EvaporationRateTest_Specimen_Rate>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Specimen_Rate>();
        }
        public int Insert_EvaporationRateTest(EvaporationRateTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "ER", "EvaporationRateTest", DateTime.Today, 2, "ReportNo");
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
                { "@Approver", DbType.String, Req.Main.Approver ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@BeforeAverageRate", DbType.Decimal, Req.Main.BeforeAverageRate } ,
                { "@AfterAverageRate", DbType.Decimal, Req.Main.AfterAverageRate } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
                { "@ReportDate", DbType.Date, Req.Main.ReportDate} ,
                { "@Preparer", DbType.String, Req.Main.Preparer ?? ""} ,
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
INSERT INTO dbo.EvaporationRateTest
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
           ,Approver
           ,FabricRefNo
           ,FabricColor
           ,FabricDescription
           ,BeforeAverageRate
           ,AfterAverageRate
           ,Status
           ,AddDate
           ,AddName
,ReportDate
,Preparer)
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
           ,@Approver
           ,@FabricRefNo
           ,@FabricColor
           ,@FabricDescription
           ,@BeforeAverageRate
           ,@AfterAverageRate
           ,'New'
           ,GETDATE()
           ,@AddName
           ,@Preparer
,@ReportDate)
;

IF EXISTS(
    SELECT 1 FROM PMSFile.dbo.EvaporationRateTest WHERE ReportNo = @ReportNo
)
BEGIN
    UPDATE PMSFile.dbo.EvaporationRateTest
    SET TestBeforePicture = @TestBeforePicture , TestAfterPicture = @TestAfterPicture
    WHERE ReportNo = @ReportNo
END
ELSE
BEGIN
    INSERT INTO PMSFile.dbo.EvaporationRateTest (ReportNo,TestBeforePicture,TestAfterPicture)
    VALUES(@ReportNo,@TestBeforePicture,@TestAfterPicture)
END
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_EvaporationRateTest(EvaporationRateTest_ViewModel Req, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
                { "@SubmitDate", Req.Main.SubmitDate} ,
                { "@Seq1", DbType.String, Req.Main.Seq1 ?? "" } ,
                { "@Seq2", DbType.String, Req.Main.Seq2 ?? "" } ,
                { "@Approver", DbType.String, Req.Main.Approver ?? "" } ,
                { "@FabricRefNo", DbType.String, Req.Main.FabricRefNo ?? "" } ,
                { "@FabricColor", DbType.String, Req.Main.FabricColor ?? "" } ,
                { "@FabricDescription", DbType.String, Req.Main.FabricDescription ?? "" } ,
                { "@BeforeAverageRate", DbType.Decimal, Req.Main.BeforeAverageRate } ,
                { "@AfterAverageRate", DbType.Decimal, Req.Main.AfterAverageRate } ,
                { "@Remark", DbType.String, Req.Main.Remark ?? "" } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
                { "@ReportDate", DbType.Date, Req.Main.ReportDate} ,
                { "@Article", DbType.String, Req.Main.Article ?? ""} ,
                { "@Preparer", DbType.String, Req.Main.Preparer ?? ""} ,
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
UPDATE EvaporationRateTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Remark = @Remark
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,Approver = @Approver
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricDescription = @FabricDescription
      ,BeforeAverageRate = @BeforeAverageRate
      ,AfterAverageRate = @AfterAverageRate
      ,ReportDate = @ReportDate
      ,Preparer  = @Preparer
,Article =@Article
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.EvaporationRateTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.EvaporationRateTest
    SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.EvaporationRateTest
        ( ReportNo ,TestAfterPicture ,TestBeforePicture)
    VALUES
        ( @ReportNo ,@TestAfterPicture ,@TestBeforePicture)
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public string Delete_EvaporationRateTest(EvaporationRateTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
SET XACT_ABORT ON

delete d
from EvaporationRateTest a 
inner join EvaporationRateTest_Detail b  on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c  on c.DetailUkey = b.Ukey
inner join EvaporationRateTest_Specimen_Rate d  on d.SpecimenUkey = c.Ukey
where a.ReportNo = @ReportNo
;
delete d
from EvaporationRateTest a 
inner join EvaporationRateTest_Detail b  on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c  on c.DetailUkey = b.Ukey
inner join EvaporationRateTest_Specimen_Time d WITH(NOLOCK) on d.SpecimenUkey = c.Ukey
where a.ReportNo = @ReportNo
;
delete c
from EvaporationRateTest a
inner join EvaporationRateTest_Detail b on a.ReportNo = b.ReportNo
inner join EvaporationRateTest_Specimen c on b.Ukey = c.DetailUkey
where a.ReportNo = @ReportNo
;
delete from EvaporationRateTest_Detail WHERE ReportNo = @ReportNo
delete from EvaporationRateTest WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.EvaporationRateTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter).ToString();
        }
        public bool Processe_EvaporationRateTest_Detail(EvaporationRateTest_ViewModel sources, string UserID, out List<EvaporationRateTest_Detail> NewDetailList)
        {
            List<EvaporationRateTest_Detail> oldDetailData = this.GetDetailList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            NewDetailList = new List<EvaporationRateTest_Detail>();

            List<EvaporationRateTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<EvaporationRateTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "ReportNo,Type",
                    "Ukey,EvaporationRateAverage");

            string insertDetail = $@" ----寫入 EvaporationRateTest_Detail
INSERT INTO EvaporationRateTest_Detail
           (ReportNo,Type,EvaporationRateAverage,EditName,EditDate)
VALUES 
           (@ReportNo ,@Type ,@EvaporationRateAverage ,@UserID ,GETDATE())
";
            string updateDetail = $@" ----更新 EvaporationRateTest_Detail
UPDATE EvaporationRateTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaporationRateAverage = @EvaporationRateAverage
WHERE ReportNo = @ReportNo
AND Type = @Type
;
";
            string deleteDetail = $@" ----刪除 EvaporationRateTest_Detail
DELETE FROM EvaporationRateTest_Detail where ReportNo = @ReportNo AND Type = @Type
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@Type", detailItem.Type));
                listDetailPar.Add(new SqlParameter($"@EvaporationRateAverage", detailItem.EvaporationRateAverage));

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        // 下一層已經有Detail Ukey，因此不另外查詢
                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        NewDetailList = new List<EvaporationRateTest_Detail>();
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

            //SQLParameterCollection objlPar = new SQLParameterCollection();
            //objlPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
            //// 需要取出新的Detail Ukey，提供下一層寫入
            //string sql = "select * from EvaporationRateTest_Detail WITH(NOLOCK) where ReportNo = @ReportNo ";
            //var tmp = ExecuteList<EvaporationRateTest_Detail>(CommandType.Text, sql, objlPar);
            //NewDetailList = tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Detail>();
            NewDetailList = this.GetDetailList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            return true;

        }

        public bool Processe_EvaporationRateTest_Specimen(EvaporationRateTest_ViewModel sources, string UserID, out List<EvaporationRateTest_Specimen> NewSpecimenList)
        {
            List<EvaporationRateTest_Specimen> oldDetailData = this.GetSpecimenList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            NewSpecimenList = new List<EvaporationRateTest_Specimen>();

            List<EvaporationRateTest_Specimen> needUpdateDetailList =
                PublicClass.CompareListValue<EvaporationRateTest_Specimen>(
                    sources.SpecimenList,
                    oldDetailData,
                    "DetailUkey,SpecimenID",
                    "RateAverage");

            string insertDetail = $@" ----寫入 EvaporationRateTest_Specimen
INSERT INTO EvaporationRateTest_Specimen
           (DetailUkey,SpecimenID,RateAverage,EditName,EditDate)
VALUES 
           (@DetailUkey,@SpecimenID,@RateAverage,@UserID,GETDATE())
";
            string updateDetail = $@" ----更新 EvaporationRateTest_Specimen
UPDATE EvaporationRateTest_Specimen
SET EditDate = GETDATE() , EditName = @UserID
    ,RateAverage = @RateAverage
WHERE DetailUkey = @DetailUkey
AND SpecimenID = @SpecimenID
;
";
            string deleteDetail = $@" ----刪除 EvaporationRateTest_Specimen
DELETE FROM EvaporationRateTest_Specimen where DetailUkey = @DetailUkey AND SpecimenID = @SpecimenID
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@DetailUkey", detailItem.DetailUkey));
                listDetailPar.Add(new SqlParameter($"@SpecimenID", detailItem.SpecimenID));
                listDetailPar.Add(new SqlParameter($"@RateAverage", detailItem.RateAverage));

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);

                        // 下一層已經有Specimen Ukey，因此不另外查詢
                        NewSpecimenList = new List<EvaporationRateTest_Specimen>();
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


            // 需要取出新的Specimen Ukey，提供下一層寫入
//            string sql = $@"
//select b.*
//,DetailType = a.Type
//from EvaporationRateTest_Detail a WITH(NOLOCK)
//inner join EvaporationRateTest_Specimen b WITH(NOLOCK) on a.Ukey=b.DetailUkey
//where b.DetailUkey IN ({(string.Join(",", needUpdateDetailList.Select(o => o.DetailUkey)))}) ";
//            var tmp = ExecuteList<EvaporationRateTest_Specimen>(CommandType.Text, sql, new SQLParameterCollection());
//            NewSpecimenList = tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Specimen>();
            NewSpecimenList = this.GetSpecimenList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            return true;

        }
        public bool Processe_EvaporationRateTest_Specimen_Time(EvaporationRateTest_ViewModel sources, string UserID, out List<EvaporationRateTest_Specimen_Time> NewTimeList)
        {
            List<EvaporationRateTest_Specimen_Time> oldDetailData = this.GetTimeList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            NewTimeList = new List<EvaporationRateTest_Specimen_Time>();

            List<EvaporationRateTest_Specimen_Time> needUpdateDetailList =
                PublicClass.CompareListValue<EvaporationRateTest_Specimen_Time>(
                    sources.TimeList,
                    oldDetailData,
                    "SpecimenUkey,Time",
                    "IsInitialMass,Ukey,Mass,Evaporation,InitialTimeUkey");

            string insertDetail = $@" ----寫入 EvaporationRateTest_Specimen_Time
IF @IsInitialMass = 0 
BEGIN
    INSERT INTO EvaporationRateTest_Specimen_Time
               (SpecimenUkey,Time,IsInitialMass,Mass,Evaporation,InitialTimeUkey,EditName,EditDate)
    VALUES 
               (@SpecimenUkey ,@Time ,@IsInitialMass ,@Mass ,@Evaporation 
            ,(SELECT Top 1 Ukey From EvaporationRateTest_Specimen_Time WITH(NOLOCK) WHERE SpecimenUkey = @SpecimenUkey AND IsInitialMass = 1)  ----取得初始值標本的 SpecimenUkey
            ,@UserID ,GETDATE())
END
ELSE
BEGIN
    INSERT INTO EvaporationRateTest_Specimen_Time
               (SpecimenUkey,Time,IsInitialMass,Mass,Evaporation,InitialTimeUkey,EditName,EditDate)
    VALUES 
               (@SpecimenUkey ,@Time ,@IsInitialMass ,@Mass ,@Evaporation 
                ,0
                ,@UserID ,GETDATE())

END
";
            string updateDetail = $@" ----更新 EvaporationRateTest_Specimen_Time
UPDATE EvaporationRateTest_Specimen_Time
SET EditDate = GETDATE() , EditName = @UserID
    ,Mass = @Mass
    ,Evaporation = @Evaporation
WHERE SpecimenUkey = @SpecimenUkey
AND Time = @Time
;
";
            string deleteDetail = $@" ----刪除 EvaporationRateTest_Specimen_Time
DELETE FROM EvaporationRateTest_Specimen_Time where SpecimenUkey = @SpecimenUkey AND Time = @Time
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@SpecimenUkey", detailItem.SpecimenUkey));
                listDetailPar.Add(new SqlParameter($"@Time", detailItem.Time));

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@Mass", detailItem.Mass));
                        listDetailPar.Add(new SqlParameter($"@Evaporation", detailItem.Evaporation));
                        listDetailPar.Add(new SqlParameter($"@IsInitialMass", detailItem.IsInitialMass));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Mass", detailItem.Mass));
                        listDetailPar.Add(new SqlParameter($"@Evaporation", detailItem.Evaporation));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);

                        // EvaporationRateTest_Specimen_Rate已經有Specimen Ukey，因此不另外查詢
                        NewTimeList = new List<EvaporationRateTest_Specimen_Time>();
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

            // 需要取出新的Specimen Ukey，提供下一層寫入
//            string sql = $@"
//-----把初始值的Ukey寫回自己身上
//update a
//set a.InitialTimeUkey = a.Ukey
//from EvaporationRateTest_Specimen_Time a 
//where a.IsInitialMass = 1 and SpecimenUkey IN ({(string.Join(",", needUpdateDetailList.Select(o => o.SpecimenUkey)))}) 

//select c.*
//,DetailType = a.Type
//,b.SpecimenID
//from EvaporationRateTest_Detail a WITH(NOLOCK)
//inner join EvaporationRateTest_Specimen b WITH(NOLOCK) on a.Ukey=b.DetailUkey
//inner join EvaporationRateTest_Specimen_Time c WITH(NOLOCK) on b.Ukey=c.SpecimenUkey
//where c.SpecimenUkey IN ({(string.Join(",", needUpdateDetailList.Select(o => o.SpecimenUkey)))}) ";
//            var tmp = ExecuteList<EvaporationRateTest_Specimen_Time>(CommandType.Text, sql, new SQLParameterCollection());
//            NewTimeList = tmp.Any() ? tmp.ToList() : new List<EvaporationRateTest_Specimen_Time>();
            NewTimeList = this.GetTimeList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();
            return true;

        }
        public bool Processe_EvaporationRateTest_Specimen_Rate(EvaporationRateTest_ViewModel sources, string UserID)
        {
            List<EvaporationRateTest_Specimen_Rate> oldDetailData = this.GetRateList(new EvaporationRateTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<EvaporationRateTest_Specimen_Rate> needUpdateDetailList =
                PublicClass.CompareListValue<EvaporationRateTest_Specimen_Rate>(
                    sources.RateList,
                    oldDetailData,
                    "SpecimenUkey,RateName",
                    "Ukey,Value,Subtrahend_TimeUkey,Minuend_TimeUkey,Ratio");

            string insertDetail = $@" ----寫入 EvaporationRateTest_Specimen_Rate
----Value計算方式：(Subtrahend_TimeUkey的Mass - Minuend_TimeUkeyy的Mass) * Ratio
INSERT INTO EvaporationRateTest_Specimen_Rate
           (SpecimenUkey,RateName,Value,Subtrahend_TimeUkey,Minuend_TimeUkey,Ratio,EditName,EditDate)
VALUES 
           (@SpecimenUkey ,@RateName ,@Value,@Subtrahend_TimeUkey ,@Minuend_TimeUkey ,@Ratio ,@UserID ,GETDATE())
";
            string updateDetail = $@" ----更新 EvaporationRateTest_Specimen_Rate
UPDATE EvaporationRateTest_Specimen_Rate
SET EditDate = GETDATE() , EditName = @UserID
    ,Value = @Value
    ,Subtrahend_TimeUkey = @Subtrahend_TimeUkey
    ,Minuend_TimeUkey = @Minuend_TimeUkey
WHERE SpecimenUkey = @SpecimenUkey
AND RateName = @RateName
;
";
            string deleteDetail = $@" ----刪除 EvaporationRateTest_Specimen_Rate
DELETE FROM EvaporationRateTest_Specimen_Rate where SpecimenUkey = @SpecimenUkey AND RateName = @RateName
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@SpecimenUkey", detailItem.SpecimenUkey));
                listDetailPar.Add(new SqlParameter($"@RateName", detailItem.RateName));

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        //listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));
                        listDetailPar.Add(new SqlParameter($"@Subtrahend_Time", detailItem.Subtrahend_Time));
                        listDetailPar.Add(new SqlParameter($"@Minuend_Time", detailItem.Minuend_Time));
                        listDetailPar.Add(new SqlParameter($"@Subtrahend_TimeUkey", detailItem.Subtrahend_TimeUkey));
                        listDetailPar.Add(new SqlParameter($"@Minuend_TimeUkey", detailItem.Minuend_TimeUkey));
                        listDetailPar.Add(new SqlParameter($"@Ratio", detailItem.Ratio));
                        listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));

                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        //listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));
                        listDetailPar.Add(new SqlParameter($"@Subtrahend_Time", detailItem.Subtrahend_Time));
                        listDetailPar.Add(new SqlParameter($"@Minuend_Time", detailItem.Minuend_Time));
                        listDetailPar.Add(new SqlParameter($"@Subtrahend_TimeUkey", detailItem.Subtrahend_TimeUkey));
                        listDetailPar.Add(new SqlParameter($"@Minuend_TimeUkey", detailItem.Minuend_TimeUkey));
                        listDetailPar.Add(new SqlParameter($"@Value", detailItem.Value));
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

        public bool UpdateAverage(EvaporationRateTest_Main request)
        {

            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", request.ReportNo);


            string sqlCmd = $@"
---- Before：EvaporationRateTest_Specimen
SELECT b.SpecimenID,AvgRate = AVG(c.Value)
INTO #BeforeSpecimenAvg
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
INNER JOIN EvaporationRateTest_Specimen_Rate c on b.Ukey = c.SpecimenUkey 
where a.ReportNo = @ReportNo and a.Type='Before'
GROUP BY b.SpecimenID


UPDATE b
SET b.RateAverage = c.AvgRate
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
INNER JOIN #BeforeSpecimenAvg c on b.SpecimenID = c.SpecimenID
where a.ReportNo = @ReportNo and a.Type='Before'

----After：EvaporationRateTest_Specimen
SELECT b.SpecimenID,AvgRate = AVG(c.Value)
INTO #AfterSpecimenAvg
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
INNER JOIN EvaporationRateTest_Specimen_Rate c on b.Ukey = c.SpecimenUkey 
where a.ReportNo = @ReportNo and a.Type='After'
GROUP BY b.SpecimenID
UPDATE b
SET b.RateAverage = c.AvgRate
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
INNER JOIN #AfterSpecimenAvg c on b.SpecimenID = c.SpecimenID
where a.ReportNo = @ReportNo and a.Type='After'

----Before：EvaporationRateTest_Detail
SELECT a.Type,AvgRate=AVG(b.RateAverage)
INTO #BeforeDetailAvg
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
where a.ReportNo = @ReportNo and a.Type='Before'
GROUP BY a.Type

UPDATE a
SET a.EvaporationRateAverage = b.AvgRate
from EvaporationRateTest_Detail a
inner join #BeforeDetailAvg b on a.Type=b.Type  
where a.ReportNo = @ReportNo

----After：EvaporationRateTest_Detail
SELECT a.Type,AvgRate=AVG(b.RateAverage)
INTO #AfterDetailAvg
from EvaporationRateTest_Detail a
inner join EvaporationRateTest_Specimen b on a.Ukey=b.DetailUkey  
where a.ReportNo = @ReportNo and a.Type='After'
GROUP BY a.Type

UPDATE a
SET a.EvaporationRateAverage = b.AvgRate
from EvaporationRateTest_Detail a
inner join #AfterDetailAvg b on a.Type=b.Type  
where a.ReportNo = @ReportNo


----Final：EvaporationRateTest
UPDATE a
SET a.BeforeAverageRate = (select b.EvaporationRateAverage from EvaporationRateTest_Detail b where a.ReportNo = b.ReportNo and b.Type='Before')
	,a.AfterAverageRate = (select b.EvaporationRateAverage from EvaporationRateTest_Detail b where a.ReportNo = b.ReportNo and b.Type='After')
FROM EvaporationRateTest a
where a.ReportNo = @ReportNo


DROP TABLE #BeforeSpecimenAvg,#AfterSpecimenAvg,#BeforeDetailAvg,#AfterDetailAvg
";

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }
        /// <summary>
        /// Encode / Amend EvaporationRateTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_EvaporationRateTest(EvaporationRateTest_Main request, string UserID)
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
UPDATE EvaporationRateTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , Result = IIF(BeforeAverageRate >= 0.2 AND AfterAverageRate >= 0.2 , 'Pass' , 'Fail')
WHERE ReportNo = @ReportNo
;
";
            }
            else
            {
                sqlCmd = $@"
UPDATE EvaporationRateTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , Result = ''
WHERE ReportNo = @ReportNo
;
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(EvaporationRateTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select TechnicianName = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
	   ,ApproverName = ISNULL(mp2.Name,pp2.Name)
	   ,ApproverSignture = t2.Signature
from EvaporationRateTest a
left join Pass1 mp on mp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Pass1 pp on pp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Technician t on t.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join Pass1 mp2 on mp2.ID = a.Approver
left join MainServer.Production.dbo.Pass1 pp2 on pp2.ID = a.Approver
left join MainServer.Production.dbo.Technician t2 on t2.ID = a.Approver
where a.ReportNo = @ReportNo
;

";
            return ExecuteDataTable(CommandType.Text, sqlCmd, paras);
        }
    }
}
