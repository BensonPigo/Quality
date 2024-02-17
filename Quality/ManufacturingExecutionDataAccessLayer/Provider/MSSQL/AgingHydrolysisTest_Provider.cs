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
    public class AgingHydrolysisTest_Provider : SQLDAL
    {
        #region 底層連線
        public AgingHydrolysisTest_Provider(string ConString) : base(ConString) { }
        public AgingHydrolysisTest_Provider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<SelectListItem> GetScales()
        {
            string sqlcmd = @"
select Text='', Value = ''
union all
select Text=ID , Value=ID
from Scale WITH(NOLOCK)  
WHERE Junk=0 
order by Value
";
            var tmp = ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return tmp.Any() ? tmp.ToList() : new List<SelectListItem>();
        }

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(AgingHydrolysisTest_Request Req)
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

        public List<AgingHydrolysisTest_Main> GetMainList(AgingHydrolysisTest_Request Req)
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
    , b.MDivisionID
from AgingHydrolysisTest a
left join SciProduction_Factory b on a.FactoryID = b.ID
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

            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                sqlcmd += " and a.OrderID = @OrderID";
                objParameter.Add("@OrderID", Req.OrderID);
            }
            if (Req.AgingHydrolysisTestID != 0)
            {
                sqlcmd += " and a.ID = @ID";
                objParameter.Add("@ID", Req.AgingHydrolysisTestID);
            }

            var tmp = ExecuteList<AgingHydrolysisTest_Main>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<AgingHydrolysisTest_Main>();
        }

        public List<AgingHydrolysisTest_Detail> GetDetailList(AgingHydrolysisTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select   b.BrandID
        ,b.SeasonID
        ,b.StyleID
        ,b.Article
        ,b.OrderID
        ,b.Temperature
        ,b.Time
        ,b.TimeUnit
        ,b.Humidity
        ,a.*
        ,c.BuyerDelivery
        ,d.TestAfterPicture
        ,d.TestBeforePicture
from AgingHydrolysisTest_Detail a WITH(NOLOCK)
inner join AgingHydrolysisTest b WITH(NOLOCK) on a.AgingHydrolysisTestID = b.ID
inner join SciProduction_Orders c on b.OrderID = c.ID
left join PMSFile.dbo.AgingHydrolysisTest_Image d WITH(NOLOCK) on a.ReportNo = d.ReportNo
where 1=1
";
            if (Req.AgingHydrolysisTestID > 0)
            {
                sqlcmd += " and a.AgingHydrolysisTestID = @AgingHydrolysisTestID" + Environment.NewLine;
                objParameter.Add("@AgingHydrolysisTestID", Req.AgingHydrolysisTestID);
            }
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and a.ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<AgingHydrolysisTest_Detail>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<AgingHydrolysisTest_Detail>();
        }

        public List<AgingHydrolysisTest_Detail_Mockup> GetMockupList(AgingHydrolysisTest_Detail_Mockup Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
from AgingHydrolysisTest_Detail_Mockup a WITH(NOLOCK)
where 1=1
";
            if (Req.Ukey > 0)
            {
                sqlcmd += " and Ukey = @Ukey" + Environment.NewLine;
                objParameter.Add("@Ukey", Req.Ukey);
            }
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<AgingHydrolysisTest_Detail_Mockup>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<AgingHydrolysisTest_Detail_Mockup>();
        }
        public List<AgingHydrolysisTest_Detail_Mockup> GetDetailMockupList(AgingHydrolysisTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
from AgingHydrolysisTest_Detail_Mockup a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<AgingHydrolysisTest_Detail_Mockup>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<AgingHydrolysisTest_Detail_Mockup>();
        }

        public AgingHydrolysisTest_Main InsertUpdate_AgingHydrolysisTest(AgingHydrolysisTest_ViewModel sources, string UserID)
        {
            AgingHydrolysisTest_Main result = new AgingHydrolysisTest_Main();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", sources.MainData.ID } ,
                { "@BrandID", sources.MainData.BrandID } ,
                { "@SeasonID", sources.MainData.SeasonID } ,
                { "@StyleID", sources.MainData.StyleID } ,
                { "@Article", sources.MainData.Article } ,
                { "@OrderID", sources.MainData.OrderID } ,
                { "@Temperature", DbType.Decimal, Convert.ToDecimal(sources.MainData.Temperature) },
                { "@Time", DbType.Decimal, Convert.ToDecimal(sources.MainData.Time) },
                { "@TimeUnit", sources.MainData.TimeUnit ?? ""} ,
                { "@Humidity", DbType.Decimal, sources.MainData.Humidity },
                { "@AddName", UserID },
                { "@EditName", UserID },
            };

            // 自動判斷AgingHydrolysisTest 是要新增還是更新
            string sqlcmd = $@"
Declare @FactoryID as varchar(10) = (select top 1 FactoryID from SciProduction_Orders where ID = @OrderID)
Declare @MDivisionID as varchar(10) = (select top 1 MDivisionID from SciProduction_Factory where ID = @FactoryID)

if exists(
    select 1 from AgingHydrolysisTest where BrandID = @BrandID AND SeasonID = @SeasonID AND StyleID = @StyleID AND Article = @Article AND OrderID = @OrderID
)
or exists(
    select 1 from AgingHydrolysisTest where ID = @ID
)
begin
    UPDATE AgingHydrolysisTest
       SET EditDate = GETDATE()
          ,Temperature = @Temperature
          ,Time = @Time
          ,TimeUnit = @TimeUnit
          ,Humidity = @Humidity
          ,EditName = @EditName
     WHERE ID = @ID
    ;
    SELECT ID = CAST( @ID as bigint) ,BrandID = @BrandID ,SeasonID = @SeasonID ,StyleID = @StyleID ,Article = @Article ,MDivisionID = @MDivisionID ,OrderID=@OrderID
end
else 
begin
    INSERT INTO dbo.AgingHydrolysisTest
               (BrandID           ,SeasonID           ,StyleID           ,Article           ,FactoryID         ,TimeUnit  
                ,OrderID           ,Temperature               ,Time           ,Humidity           ,AddDate           ,AddName)
    VALUES
               (@BrandID           ,@SeasonID           ,@StyleID           ,@Article           ,@FactoryID         ,@TimeUnit  
                ,@OrderID           ,@Temperature               ,@Time           ,@Humidity           ,GETDATE()           ,@AddName)
    ;
    SELECT ID = CAST( @@IDENTITY as bigint),BrandID = @BrandID ,SeasonID = @SeasonID ,StyleID = @StyleID ,Article = @Article ,MDivisionID = @MDivisionID ,OrderID=@OrderID
end

";

            var tmp = ExecuteList<AgingHydrolysisTest_Main>(CommandType.Text, sqlcmd, objParameter);

            if (tmp != null && tmp.Any())
            {
                // 回傳AgingHydrolysisTest
                result = tmp.FirstOrDefault();
            }

            return result;
        }
        public bool Delete_AgingHydrolysisTest(AgingHydrolysisTest_ViewModel sources)
        {
            AgingHydrolysisTest_Main result = new AgingHydrolysisTest_Main();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", sources.MainData.ID } ,
            };

            // 自動判斷AgingHydrolysisTest 是要新增還是更新
            string sqlcmd = $@"
SET XACT_ABORT ON

Delete c
from AgingHydrolysisTest_Detail_Mockup c
inner join AgingHydrolysisTest_Detail b on b.ReportNo = c.ReportNo
inner join AgingHydrolysisTest a on a.ID = b.AgingHydrolysisTestID
where a.ID = @ID

Delete c
from PMSFile.dbo.AgingHydrolysisTest_Image c
inner join AgingHydrolysisTest_Detail b on b.ReportNo = c.ReportNo
inner join AgingHydrolysisTest a on a.ID = b.AgingHydrolysisTestID
where a.ID = @ID

Delete b
from AgingHydrolysisTest_Detail b
inner join AgingHydrolysisTest a on a.ID = b.AgingHydrolysisTestID
where a.ID = @ID

DELETE FROM AgingHydrolysisTest where ID = @ID
";

            ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
            return true;
        }
        public bool Processe_AgingHydrolysisTest_Detail(AgingHydrolysisTest_ViewModel sources, string UserID ,bool isSaveDetailPage = false)
        {

            List<AgingHydrolysisTest_Detail> oldDetailData = this.GetDetailList(new AgingHydrolysisTest_Request() { AgingHydrolysisTestID = sources.MainData.ID }).ToList();

            // 若是Detail頁面的Save，只需比對相同ReportNo的資料
            if (isSaveDetailPage && sources.DetailList.Any())
            {
                oldDetailData = oldDetailData.Where(o => o.ReportNo == sources.DetailList.FirstOrDefault().ReportNo).ToList();
            }

            List<AgingHydrolysisTest_Detail> needUpdateDetailList =new List<AgingHydrolysisTest_Detail>();

            // 若是Detail頁面的Save，才需比對所有欄位；Index的Save只能異動MaterialType
            if (isSaveDetailPage)
            {
                needUpdateDetailList =
                PublicClass.CompareListValue<AgingHydrolysisTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "ReportNo",
                    "MaterialType,ReportDate,ReceivedDate,FabricRefNo,AccRefNo,FabricColor,AccColor,Result,Comment");

            }
            else
            {
                needUpdateDetailList =
                PublicClass.CompareListValue<AgingHydrolysisTest_Detail>(
                    sources.DetailList,
                    oldDetailData,
                    "ReportNo",
                    "MaterialType");

            }

            string insertDetail = $@" ----寫入AgingHydrolysisTest_Detail
SET XACT_ABORT ON

INSERT INTO AgingHydrolysisTest_Detail
           (AgingHydrolysisTestID
           ,ReportNo
           ,MaterialType
           ,Status
           ,AddDate
           ,AddName)
VALUES 
           (@AgingHydrolysisTestID
           ,@ReportNo
           ,@MaterialType
           ,'New'
           ,GETDATE()
           ,@AddName)
;

if not exists(
    select * from PMSFile.dbo.AgingHydrolysisTest_Image 
)
begin
    INSERT INTO PMSFile.dbo.AgingHydrolysisTest_Image 
        ( ReportNo)
    VALUES
        ( @ReportNo)
end
";
            string updateDetail = $@" ----更新AgingHydrolysisTest_Detail
SET XACT_ABORT ON

UPDATE AgingHydrolysisTest_Detail
SET EditDate = GETDATE() , EditName = @EditName
    ,MaterialType = @MaterialType
    ,ReportDate = @ReportDate
    ,ReceivedDate = @ReceivedDate
    ,FabricRefNo = @FabricRefNo
    ,AccRefNo = @AccRefNo
    ,FabricColor = @FabricColor
    ,AccColor = @AccColor
    ,Result = @Result
    ,Comment = @Comment
WHERE ReportNo = @ReportNo
;
if @MaterialType != 'Mockup'
begin
    delete from AgingHydrolysisTest_Detail_Mockup where ReportNo = @ReportNo
end 
;
UPDATE  PMSFile.dbo.AgingHydrolysisTest_Image 
SET TestBeforePicture = @TestBeforePicture ,TestAfterPicture = @TestAfterPicture
WHERE ReportNo = @ReportNo
";
            string deleteDetail = $@" ----刪除AgingHydrolysisTest_Detail
SET XACT_ABORT ON

DELETE FROM AgingHydrolysisTest_Detail_Mockup where ReportNo = @ReportNo
DELETE FROM AgingHydrolysisTest_Detail where ReportNo = @ReportNo
DELETE FROM PMSFile.dbo.AgingHydrolysisTest_Image  where ReportNo = @ReportNo
";
            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@AgingHydrolysisTestID", sources.MainData.ID));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:

                        // 取得新ReportNo
                        string newReportNo = GetID(sources.MainData.MDivisionID + "AI", "AgingHydrolysisTest_Detail", DateTime.Today, 2, "ReportNo");

                        listDetailPar.Add(new SqlParameter($"@ReportNo", newReportNo));
                        listDetailPar.Add(new SqlParameter($"@MaterialType", detailItem.MaterialType));
                        listDetailPar.Add(new SqlParameter($"@AddName", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@ReportNo", detailItem.ReportNo));
                        listDetailPar.Add(new SqlParameter($"@MaterialType", detailItem.MaterialType ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ReportDate", detailItem.ReportDate));
                        listDetailPar.Add(new SqlParameter($"@ReceivedDate", detailItem.ReceivedDate));
                        listDetailPar.Add(new SqlParameter($"@FabricRefNo", detailItem.FabricRefNo ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AccRefNo", detailItem.AccRefNo ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@FabricColor", detailItem.FabricColor ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AccColor", detailItem.AccColor ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Comment", detailItem.Comment ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Result", detailItem.Result ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EditName", UserID));

                        if (detailItem.TestBeforePicture != null)
                        {
                            listDetailPar.Add("@TestBeforePicture", detailItem.TestBeforePicture);
                        }
                        else
                        {
                            listDetailPar.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
                        }

                        if (detailItem.TestAfterPicture != null)
                        {
                            listDetailPar.Add("@TestAfterPicture", detailItem.TestAfterPicture);
                        }
                        else
                        {
                            listDetailPar.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
                        }

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case CompareStateType.Delete:
                        listDetailPar.Add(new SqlParameter($"@ReportNo", detailItem.ReportNo));

                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case CompareStateType.None:

                        break;
                    default:
                        break;
                }


            }

            // AgingHydrolysisTest_Detail沒有異動，但AgingHydrolysisTest_Image有異動，走以下路線
            if (!needUpdateDetailList.Any() && sources.DetailList.Any())
            {

                string updateImageOnly = $@" ----修改AgingHydrolysisTest_Detail圖片
SET XACT_ABORT ON
UPDATE  PMSFile.dbo.AgingHydrolysisTest_Image 
SET TestBeforePicture = @TestBeforePicture ,TestAfterPicture = @TestAfterPicture
WHERE ReportNo = @ReportNo
";

                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.DetailList.FirstOrDefault().ReportNo));
                foreach (var detailItem in sources.DetailList)
                {

                    if (detailItem.TestBeforePicture != null)
                    {
                        listDetailPar.Add("@TestBeforePicture", detailItem.TestBeforePicture);
                    }
                    else
                    {
                        listDetailPar.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
                    }

                    if (detailItem.TestAfterPicture != null)
                    {
                        listDetailPar.Add("@TestAfterPicture", detailItem.TestAfterPicture);
                    }
                    else
                    {
                        listDetailPar.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
                    }

                    ExecuteNonQuery(CommandType.Text, updateImageOnly, listDetailPar);

                }
            }
            return true;

        }
        public bool Process_AgingHydrolysisTest_Detail_Mockup(AgingHydrolysisTest_Detail_ViewModel sources, string UserID)
        {

            List<AgingHydrolysisTest_Detail_Mockup> oldDetailData = this.GetMockupList(new AgingHydrolysisTest_Detail_Mockup() { ReportNo = sources.MainDetailData.ReportNo }).ToList();

            List<AgingHydrolysisTest_Detail_Mockup> needUpdateDetailList =
                PublicClass.CompareListValue<AgingHydrolysisTest_Detail_Mockup>(
                    sources.MockupList,
                    oldDetailData,
                    "ReportNo,SpecimenName",
                    "ChangeScale,ChangeResult,StainingScale,StainingResult,Comment");

            string insertDetail = $@" ----AgingHydrolysisTest_Detail_Mockup
INSERT INTO AgingHydrolysisTest_Detail_Mockup
           (ReportNo
           ,SpecimenName
           ,ChangeScaleStandard
           ,ChangeScale
           ,ChangeResult
           ,StainingScaleStandard
           ,StainingScale
           ,StainingResult
           ,Comment
           ,EditDate
           ,EditName)
VALUES 
           (@ReportNo
           ,@SpecimenName
           ,'4-5'
           ,@ChangeScale
           ,@ChangeResult
           ,'4'
           ,@StainingScale
           ,@StainingResult
           ,@Comment
           ,GETDATE()
           ,@EditName)
";
            string updateDetail = $@" ----AgingHydrolysisTest_Detail_Mockup
UPDATE AgingHydrolysisTest_Detail_Mockup
SET EditDate = GETDATE() , EditName = @EditName
    ,ChangeScale = @ChangeScale
    ,ChangeResult = @ChangeResult
    ,StainingScale = @StainingScale
    ,StainingResult = @StainingResult
    ,Comment = @Comment
WHERE Ukey = @Ukey AND ReportNo = @ReportNo
";
            string deleteDetail = $@" ----刪除AgingHydrolysisTest_Detail_Mockup
DELETE FROM AgingHydrolysisTest_Detail_Mockup 
where Ukey = @Ukey AND ReportNo = @ReportNo
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@ReportNo", sources.MainDetailData.ReportNo));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@SpecimenName", detailItem.SpecimenName ?? string.Empty));

                        listDetailPar.Add(new SqlParameter($"@ChangeScaleStandard", detailItem.ChangeScaleStandard ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ChangeScale", detailItem.ChangeScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ChangeResult", detailItem.ChangeResult ?? string.Empty));

                        listDetailPar.Add(new SqlParameter($"@StainingScaleStandard", detailItem.StainingScaleStandard ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@StainingScale", detailItem.StainingScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@StainingResult", detailItem.StainingResult ?? string.Empty));

                        listDetailPar.Add(new SqlParameter($"@Comment", detailItem.Comment ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EditName", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    case CompareStateType.Edit:
                        // 目前不允許修改 SpecimenName
                        listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                        listDetailPar.Add(new SqlParameter($"@ChangeScale", detailItem.ChangeScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ChangeResult", detailItem.ChangeResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@StainingScale", detailItem.StainingScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@StainingResult", detailItem.StainingResult ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@Comment", detailItem.Comment ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@EditName", UserID));

                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case CompareStateType.Delete:
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
        /// Encode / Amend AgingHydrolysisTest_Detail
        /// </summary>
        /// <param name="request"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public bool EncodeAmend_AgingHydrolysisTest_Detail(AgingHydrolysisTest_Detail request, string UserID)
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
UPDATE AgingHydrolysisTest_Detail
SET EditDate = GETDATE() , EditName = @EditName
    , Status = @Status
    , Result = @Result
    , ReportDate = GETDATE()
WHERE ReportNo = @ReportNo
";
            }
            else
            {
                sqlCmd = $@"
UPDATE AgingHydrolysisTest_Detail
SET EditDate = GETDATE() , EditName = @EditName
    , Status = 'New'
    , Result = ''
    , ReportDate = NULL
WHERE ReportNo = @ReportNo
";
            }

            ExecuteNonQuery(CommandType.Text, sqlCmd, paras);

            return true;
        }

        public DataTable GetReportTechnician(AgingHydrolysisTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from AgingHydrolysisTest_Detail a
left join Pass1 mp on mp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Pass1 pp on pp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Technician t on t.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
where a.ReportNo = @ReportNo
;

";
            return ExecuteDataTable(CommandType.Text, sqlCmd, paras);
        }
        public DataSet GetReport(AgingHydrolysisTest_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select b.BrandID
	,a.ReportNo
	,b.OrderID
	,b.FactoryID
	,a.ReceivedDate
	,b.StyleID
	,a.ReportDate
	,b.Article
	,a.FabricRefNo
	,b.SeasonID
	,a.FabricColor
	,a.MaterialType
	,a.Result
	,d.TestBeforePicture
	,d.TestAfterPicture
	,Technician = ISNULL(mp.Name,pp.Name)
	,a.Comment
from AgingHydrolysisTest_Detail a
inner join AgingHydrolysisTest b on b.ID = a.AgingHydrolysisTestID
left join PMSFile.dbo.AgingHydrolysisTest_Image  d on a.ReportNo = d.ReportNo
left join Pass1 mp on a.EditName = mp.ID
left join SciProduction_Pass1 pp on a.EditName = pp.ID
where a.ReportNo = @ReportNo

select *
from AgingHydrolysisTest_Detail_Mockup
where ReportNo = @ReportNo
";
            return ExecuteDataSet(CommandType.Text, sqlCmd, paras);
        }
    }
}
