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
    public class RandomTumblePillingTestProvider : SQLDAL
    {
        #region 底層連線
        public RandomTumblePillingTestProvider(string ConString) : base(ConString) { }
        public RandomTumblePillingTestProvider(SQLDataTransaction tra) : base(tra) { }
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

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(RandomTumblePillingTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();

            string SbSql = $@"

select o.*, oa.Article 
from Orders o
inner join Order_Article oa on oa.id = o.ID
inner join Style s on s.Ukey = o.StyleUkey
where o.Category ='B'  --只抓大貨單
and s.FabricType IN('KNIT' ,'I')
";

            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql += $@" and o.ID = @OrderID" + Environment.NewLine;
                paras.Add("@OrderID", DbType.String, Req.OrderID);
            }

            var tmp = ExecuteList<DatabaseObject.ProductionDB.Orders>(CommandType.Text, SbSql, paras);


            return tmp.Any() ? tmp.ToList() : new List<DatabaseObject.ProductionDB.Orders>();
        }

        public List<RandomTumblePillingTest_Main> GetMainList(RandomTumblePillingTest_Request Req)
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
        ,d.TestFaceSideBeforePicture
        ,d.TestFaceSideAfterPicture
        ,d.TestBackSideBeforePicture
        ,d.TestBackSideAfterPicture
from RandomTumblePillingTest a
left join PMSFile.dbo.RandomTumblePillingTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<RandomTumblePillingTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<RandomTumblePillingTest_Main>();
        }

        public List<RandomTumblePillingTest_Detail> GetDetailList(RandomTumblePillingTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from RandomTumblePillingTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<RandomTumblePillingTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<RandomTumblePillingTest_Detail>();
        }

        public int Insert_RandomTumblePillingTest(RandomTumblePillingTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "TP", "RandomTumblePillingTest", DateTime.Today, 2, "ReportNo");
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
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestFaceSideBeforePicture != null)
            {
                objParameter.Add("@TestFaceSideBeforePicture", Req.Main.TestFaceSideBeforePicture);
            }
            else
            {
                objParameter.Add("@TestFaceSideBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestFaceSideAfterPicture != null)
            {
                objParameter.Add("@TestFaceSideAfterPicture", Req.Main.TestFaceSideAfterPicture);
            }
            else
            {
                objParameter.Add("@TestFaceSideAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestBackSideBeforePicture != null)
            {
                objParameter.Add("@TestBackSideBeforePicture", Req.Main.TestBackSideBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBackSideBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestBackSideAfterPicture != null)
            {
                objParameter.Add("@TestBackSideAfterPicture", Req.Main.TestBackSideAfterPicture);
            }
            else
            {
                objParameter.Add("@TestBackSideAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"

INSERT INTO dbo.RandomTumblePillingTest
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
           ,TestStandard
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
           ,@TestStandard
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
)
;

INSERT INTO PMSFile.dbo.RandomTumblePillingTest
    ( ReportNo ,TestFaceSideBeforePicture ,TestFaceSideAfterPicture ,TestBackSideBeforePicture,TestBackSideAfterPicture)
VALUES
    ( @ReportNo ,@TestFaceSideBeforePicture ,@TestFaceSideAfterPicture ,@TestBackSideBeforePicture,@TestBackSideAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_RandomTumblePillingTest(RandomTumblePillingTest_ViewModel Req, string UserID)
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
                { "@TestStandard", DbType.String, Req.Main.TestStandard ?? "" } ,
                { "@EditName", DbType.String, UserID ?? "" } ,
            };

            if (Req.Main.TestFaceSideBeforePicture != null)
            {
                objParameter.Add("@TestFaceSideBeforePicture", Req.Main.TestFaceSideBeforePicture);
            }
            else
            {
                objParameter.Add("@TestFaceSideBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestFaceSideAfterPicture != null)
            {
                objParameter.Add("@TestFaceSideAfterPicture", Req.Main.TestFaceSideAfterPicture);
            }
            else
            {
                objParameter.Add("@TestFaceSideAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestBackSideBeforePicture != null)
            {
                objParameter.Add("@TestBackSideBeforePicture", Req.Main.TestBackSideBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBackSideBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.Main.TestBackSideAfterPicture != null)
            {
                objParameter.Add("@TestBackSideAfterPicture", Req.Main.TestBackSideAfterPicture);
            }
            else
            {
                objParameter.Add("@TestBackSideAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }


            string mainSqlCmd = $@"

UPDATE RandomTumblePillingTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,TestStandard = @TestStandard
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.RandomTumblePillingTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.RandomTumblePillingTest
    SET TestFaceSideBeforePicture = @TestFaceSideBeforePicture , TestFaceSideAfterPicture=@TestFaceSideAfterPicture ,TestBackSideBeforePicture=@TestBackSideBeforePicture ,TestBackSideAfterPicture=@TestBackSideAfterPicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.RandomTumblePillingTest
        ( ReportNo ,TestFaceSideBeforePicture ,TestFaceSideAfterPicture ,TestBackSideBeforePicture ,TestBackSideAfterPicture)
    VALUES
        ( @ReportNo ,@TestFaceSideBeforePicture ,@TestFaceSideAfterPicture ,@TestBackSideBeforePicture ,@TestBackSideAfterPicture)
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_RandomTumblePillingTest(RandomTumblePillingTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"

delete from RandomTumblePillingTest WHERE ReportNo = @ReportNo
delete from RandomTumblePillingTest_Detail WHERE ReportNo = @ReportNo
delete from PMSFile.dbo.RandomTumblePillingTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_RandomTumblePillingTest_Detail(RandomTumblePillingTest_ViewModel sources, string UserID)
        {
            List<RandomTumblePillingTest_Detail> oldDetailData = this.GetDetailList(new RandomTumblePillingTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<RandomTumblePillingTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<RandomTumblePillingTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "Side,EvaluationItem,FirstScale,SecondScale,IsEvenly,Result,Remark");

            string insertDetail = $@" ----寫入 RandomTumblePillingTest_Detail

INSERT INTO RandomTumblePillingTest_Detail
           (ReportNo,Side,EvaluationItem,FirstScale,SecondScale,IsEvenly,Result,Remark,EditDate,EditName)
VALUES 
           (@ReportNo,@Side,@EvaluationItem,@FirstScale,@SecondScale,@IsEvenly,@Result,@Remark,GETDATE(),@UserID)

";
            string updateDetail = $@" ----更新 RandomTumblePillingTest_Detail

UPDATE RandomTumblePillingTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,Side = @Side
    ,EvaluationItem = @EvaluationItem
    ,FirstScale = @FirstScale
    ,SecondScale = @SecondScale
    ,IsEvenly = @IsEvenly
    ,Result = @Result
    ,Remark = @Remark
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 RandomTumblePillingTest_Detail
DELETE FROM RandomTumblePillingTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@Side", detailItem.Side));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@FirstScale", detailItem.FirstScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@SecondScale", detailItem.SecondScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@IsEvenly", detailItem.IsEvenly));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@Side", detailItem.Side));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@FirstScale", detailItem.FirstScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@SecondScale", detailItem.SecondScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@IsEvenly", detailItem.IsEvenly));

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
        /// Encode / Amend RandomTumblePillingTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_RandomTumblePillingTest(RandomTumblePillingTest_Main request, string UserID)
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
UPDATE RandomTumblePillingTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE RandomTumblePillingTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New' ,Result=''
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(RandomTumblePillingTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from RandomTumblePillingTest a
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
