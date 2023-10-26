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
    public class SalivaFastnessTestProvider : SQLDAL
    {
        #region 底層連線
        public SalivaFastnessTestProvider(string ConString) : base(ConString) { }
        public SalivaFastnessTestProvider(SQLDataTransaction tra) : base(tra) { }
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

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(SalivaFastnessTest_Request Req)
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

        public List<SalivaFastnessTest_Main> GetMainList(SalivaFastnessTest_Request Req)
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
from SalivaFastnessTest a
left join SciPMSFile_SalivaFastnessTest d WITH(NOLOCK) on a.ReportNo = d.ReportNo
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

            var tmp = ExecuteList<SalivaFastnessTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<SalivaFastnessTest_Main>();
        }

        public List<SalivaFastnessTest_Detail> GetDetailList(SalivaFastnessTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   a.*
from SalivaFastnessTest_Detail a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<SalivaFastnessTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<SalivaFastnessTest_Detail>();
        }

        public int Insert_SalivaFastnessTest(SalivaFastnessTest_ViewModel Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "SF", "SalivaFastnessTest", DateTime.Today, 2, "ReportNo");
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
                { "@ItemTested", DbType.String, Req.Main.ItemTested ?? "" } ,
                { "@TypeOfPrint", DbType.String, Req.Main.TypeOfPrint ?? "" } ,
                { "@PrintColor", DbType.String, Req.Main.PrintColor ?? "" } ,
                { "@Result", DbType.String, Req.Main.Result ?? "Pass" } ,
                { "@Temperature", Req.Main.Temperature } ,
                { "@Time", Req.Main.Time } ,
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
SET XACT_ABORT ON

INSERT INTO dbo.SalivaFastnessTest
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
           ,TypeOfPrint
           ,ItemTested
           ,PrintColor
           ,Temperature
           ,Time
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
           ,@TypeOfPrint
           ,@ItemTested
           ,@PrintColor
           ,@Temperature
           ,@Time
           ,'New'
           ,@Result
           ,GETDATE()
           ,@AddName
)
;

INSERT INTO SciPMSFile_SalivaFastnessTest
    ( ReportNo ,TestBeforePicture ,TestAfterPicture)
VALUES
    ( @ReportNo ,@TestBeforePicture ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int Update_SalivaFastnessTest(SalivaFastnessTest_ViewModel Req, string UserID)
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
                { "@TypeOfPrint", DbType.String, Req.Main.TypeOfPrint ?? "" } ,
                { "@ItemTested", DbType.String, Req.Main.ItemTested ?? "" } ,
                { "@PrintColor", DbType.String, Req.Main.PrintColor ?? "" } ,
                { "@Temperature", Req.Main.Temperature } ,
                { "@Time", Req.Main.Time } ,
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
SET XACT_ABORT ON

UPDATE SalivaFastnessTest
   SET EditDate = GETDATE()
      ,EditName = @EditName
      ,Result = @Result
      ,SubmitDate = @SubmitDate
      ,Seq1 = @Seq1
      ,Seq2 = @Seq2
      ,FabricRefNo = @FabricRefNo
      ,FabricColor = @FabricColor
      ,FabricDescription = @FabricDescription
      ,TypeOfPrint = @TypeOfPrint
      ,ItemTested = @ItemTested
      ,PrintColor = @PrintColor
      ,Temperature = @Temperature
      ,Time = @Time
WHERE ReportNo = @ReportNo
;
if exists(select 1 from PMSFile.dbo.SalivaFastnessTest WHERE ReportNo = @ReportNo)
begin
    UPDATE PMSFile.dbo.SalivaFastnessTest
    SET TestAfterPicture = @TestAfterPicture , TestBeforePicture=@TestBeforePicture
    WHERE ReportNo = @ReportNo
end
else
begin
    INSERT INTO PMSFile.dbo.SalivaFastnessTest
        ( ReportNo ,TestAfterPicture ,TestBeforePicture )
    VALUES
        ( @ReportNo ,@TestAfterPicture ,@TestBeforePicture )
end
";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public int Delete_SalivaFastnessTest(SalivaFastnessTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.Main.ReportNo } ,
            };

            string mainSqlCmd = $@"
SET XACT_ABORT ON

delete from SalivaFastnessTest WHERE ReportNo = @ReportNo
delete from SalivaFastnessTest_Detail WHERE ReportNo = @ReportNo
delete from SciPMSFile_SalivaFastnessTest WHERE ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, mainSqlCmd.ToString(), objParameter);
        }
        public bool Processe_SalivaFastnessTest_Detail(SalivaFastnessTest_ViewModel sources, string UserID)
        {
            List<SalivaFastnessTest_Detail> oldDetailData = this.GetDetailList(new SalivaFastnessTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            List<SalivaFastnessTest_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<SalivaFastnessTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "Ukey",
                    "EvaluationItem,AllResult,AcetateScale,CottonScale,NylonScale,PolyesterScale,AcrylicScale,WoolScale,AcetateResult,CottonResult,NylonResult,PolyesterResult,AcrylicResult,WoolResult,Remark");

            string insertDetail = $@" ----寫入 SalivaFastnessTest_Detail

INSERT INTO SalivaFastnessTest_Detail
           (ReportNo,EvaluationItem,AllResult,AcetateScale,CottonScale,NylonScale,PolyesterScale,AcrylicScale,WoolScale,AcetateResult,CottonResult,NylonResult,PolyesterResult,AcrylicResult,WoolResult,Remark,EditName,EditDate)
VALUES 
           (@ReportNo,@EvaluationItem,@AllResult,@AcetateScale,@CottonScale,@NylonScale,@PolyesterScale,@AcrylicScale,@WoolScale,@AcetateResult,@CottonResult,@NylonResult,@PolyesterResult,@AcrylicResult,@WoolResult,@Remark,@UserID,GETDATE())

";
            string updateDetail = $@" ----更新 SalivaFastnessTest_Detail


UPDATE SalivaFastnessTest_Detail
SET EditDate = GETDATE() , EditName = @UserID
    ,EvaluationItem = @EvaluationItem
    ,AllResult = @AllResult
    ,Remark = @Remark

    ,AcetateScale = @AcetateScale
    ,CottonScale = @CottonScale
    ,NylonScale = @NylonScale
    ,PolyesterScale = @PolyesterScale
    ,AcrylicScale = @AcrylicScale
    ,WoolScale = @WoolScale

    ,AcetateResult = @AcetateResult
    ,CottonResult = @CottonResult
    ,NylonResult = @NylonResult
    ,PolyesterResult = @PolyesterResult
    ,AcrylicResult = @AcrylicResult
    ,WoolResult = @WoolResult
WHERE ReportNo = @ReportNo
AND Ukey = @Ukey
;
";
            string deleteDetail = $@" ----刪除 SalivaFastnessTest_Detail
DELETE FROM SalivaFastnessTest_Detail where ReportNo = @ReportNo AND Ukey = @Ukey
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
                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcetateResult", detailItem.AcetateResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonResult", detailItem.CottonResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonResult", detailItem.NylonResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterResult", detailItem.PolyesterResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicResult", detailItem.AcrylicResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolResult", detailItem.WoolResult ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@EvaluationItem", detailItem.EvaluationItem));
                        listDetailPar.Add(new SqlParameter($"@AllResult", detailItem.AllResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcetateResult", detailItem.AcetateResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonResult", detailItem.CottonResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonResult", detailItem.NylonResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterResult", detailItem.PolyesterResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicResult", detailItem.AcrylicResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolResult", detailItem.WoolResult ?? string.Empty));

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
        /// Encode / Amend SalivaFastnessTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_SalivaFastnessTest(SalivaFastnessTest_Main request, string UserID)
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
UPDATE SalivaFastnessTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE SalivaFastnessTest
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(SalivaFastnessTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from SalivaFastnessTest a
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
