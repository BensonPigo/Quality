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
    public class BrandBulkTestProvider : SQLDAL
    {
        #region 底層連線
        public BrandBulkTestProvider(string ConString) : base(ConString) { }
        public BrandBulkTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public List<SelectListItem> GetArtworkSource()
        {
            string sqlcmd = @"
select distinct  Text=ID , Value=ID
from ArtworkType WITH(NOLOCK)  
WHERE Junk=0 
";
            var tmp = ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return tmp.Any() ? tmp.ToList() : new List<SelectListItem>();
        }

        public List<DatabaseObject.ProductionDB.Orders> GetOrderInfo(BrandBulkTest_Request Req)
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

        public List<BrandBulkTest> GetMainList(BrandBulkTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@TestItem", DbType.String, Req.TestItem } ,
            };

            string sqlcmd = $@"
select a.*
,b.TestItem
from BrandBulkTest a WITH(NOLOCK)
left join BrandBulkTestItem b  WITH(NOLOCK) on a.TestItemUkey = b.Ukey
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

            if (!string.IsNullOrEmpty(Req.TestItem))
            {
                sqlcmd += " and b.TestItem = @TestItem";
                objParameter.Add("@TestItem", Req.TestItem);
            }


            var tmp = ExecuteList<BrandBulkTest>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<BrandBulkTest>();
        }
        public List<BrandBulkTestItem> GetBrandBulkTestItemList(BrandBulkTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select *
from BrandBulkTestItem a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                sqlcmd += " and a.BrandID = @BrandID" + Environment.NewLine;
                objParameter.Add("@BrandID", Req.BrandID);
            }

            var tmp = ExecuteList<BrandBulkTestItem>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<BrandBulkTestItem>();
        }
        public List<BrandBulkTestDox> GetBrandBulkTestDoxList(BrandBulkTest_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlcmd = $@"
select a.*
,IsOldFile = Cast(1 as bit)
from BrandBulkTestDox a WITH(NOLOCK)
where 1=1
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                sqlcmd += " and ReportNo = @ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", Req.ReportNo);
            }

            var tmp = ExecuteList<BrandBulkTestDox>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<BrandBulkTestDox>();
        }


        public List<BrandBulkTestDox> GetBrandBulkTestDoxList(List<BrandBulkTestDox> ReqList)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string ukeys = string.Join(",", ReqList.Select(o => o.Ukey));

            string sqlcmd = $@"
select a.*
,IsOldFile = Cast(1 as bit)
from BrandBulkTestDox a WITH(NOLOCK)
where 1=1
AND Ukey IN ({ukeys})
";          
            var tmp = ExecuteList<BrandBulkTestDox>(CommandType.Text, sqlcmd, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<BrandBulkTestDox>();
        }
        public int Insert_BrandBulkTest(BrandBulkTest_ViewModel sources, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "BB", "BrandBulkTest", DateTime.Today, 2, "ReportNo");
            BrandBulkTest result = new BrandBulkTest();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ReportNo", NewReportNo } ,
                { "@BrandID", sources.Main.BrandID } ,
                { "@SeasonID", sources.Main.SeasonID } ,
                { "@StyleID", sources.Main.StyleID } ,
                { "@Article", sources.Main.Article } ,
                { "@OrderID", sources.Main.OrderID } ,
                { "@ReportDate", sources.Main.ReportDate } ,
                { "@TestItemUkey", sources.Main.TestItemUkey } ,

                { "@FabricRefno", sources.Main.FabricRefno } ,
                { "@FabricColor", sources.Main.FabricColor } ,
                { "@AccessoryRefno", sources.Main.AccessoryRefno } ,
                { "@AccessoryColor", sources.Main.AccessoryColor } ,
                { "@Artwork", sources.Main.Artwork } ,
                { "@ArtworkRefno", sources.Main.ArtworkRefno } ,
                { "@ArtworkColor", sources.Main.ArtworkColor } ,

                { "@Result", sources.Main.Result ?? "Pass"} ,
                { "@Remark", sources.Main.Remark ?? ""} ,
                { "@AddName", UserID },
                { "@EditName", UserID },
            };

            // 自動判斷BrandBulkTest 是要新增還是更新
            string sqlcmd = $@"
    INSERT INTO dbo.BrandBulkTest
               (ReportNo           ,BrandID           ,SeasonID           ,StyleID           ,Article           ,OrderID           ,FactoryID
                ,FabricRefno           ,FabricColor           ,AccessoryRefno           ,AccessoryColor           ,Artwork           ,ArtworkRefno           ,ArtworkColor
               ,ReportDate           ,TestItemUkey           ,Result           ,Remark           ,AddDate           ,AddName           ,EditDate           ,EditName)
    VALUES
               (@ReportNo           ,@BrandID           ,@SeasonID           ,@StyleID           ,@Article           ,@OrderID           ,(select top 1 FactoryID from SciProduction_Orders with(NOLOCK) where ID = @OrderID)
               ,@FabricRefno           ,@FabricColor           ,@AccessoryRefno           ,@AccessoryColor           ,@Artwork           ,@ArtworkRefno           ,@ArtworkColor
               ,@ReportDate           ,@TestItemUkey           ,@Result           ,@Remark           ,GETDATE()           ,@AddName           ,GETDATE()           ,@EditName)

";

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }
        public int Update_BrandBulkTest(BrandBulkTest_ViewModel sources, string UserID)
        {
            BrandBulkTest result = new BrandBulkTest();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ReportNo", sources.Main.ReportNo } ,
                { "@ReportDate", sources.Main.ReportDate } ,

                { "@FabricRefno", sources.Main.FabricRefno } ,
                { "@FabricColor", sources.Main.FabricColor } ,
                { "@AccessoryRefno", sources.Main.AccessoryRefno } ,
                { "@AccessoryColor", sources.Main.AccessoryColor } ,
                { "@Artwork", sources.Main.Artwork } ,
                { "@ArtworkRefno", sources.Main.ArtworkRefno } ,
                { "@ArtworkColor", sources.Main.ArtworkColor } ,

                { "@Result", sources.Main.Result ?? "Pass"} ,
                { "@Remark", sources.Main.Remark ?? ""} ,
                { "@EditName", UserID },
            };

            // 自動判斷BrandBulkTest 是要新增還是更新
            string sqlcmd = $@"
UPDATE BrandBulkTest
    SET EditDate = GETDATE()
        ,EditName = @EditName
        ,ReportDate = @ReportDate
        ,Result = @Result
        ,Remark = @Remark
        ,FabricRefno = @FabricRefno
        ,FabricColor = @FabricColor
        ,AccessoryRefno = @AccessoryRefno
        ,AccessoryColor = @AccessoryColor
        ,Artwork = @Artwork
        ,ArtworkRefno = @ArtworkRefno
        ,ArtworkColor = @ArtworkColor
WHERE ReportNo = @ReportNo
";

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }
        public bool Delete_BrandBulkTest(BrandBulkTest_ViewModel sources)
        {
            BrandBulkTest result = new BrandBulkTest();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ReportNo", sources.Main.ReportNo } ,
            };

            string sqlcmd = $@"
DELETE FROM BrandBulkTest where ReportNo = @ReportNo
DELETE FROM BrandBulkTestDox where ReportNo = @ReportNo
";

            ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
            return true;
        }


        public bool Processe_BrandBulkTestDox(BrandBulkTest_ViewModel sources, string UserID ,bool isSaveDetailPage = false)
        {

            List<BrandBulkTestDox> oldDetailData = this.GetBrandBulkTestDoxList(new BrandBulkTest_Request() { ReportNo = sources.Main.ReportNo }).ToList();

            // 若是Detail頁面的Save，只需比對相同ReportNo的資料
            if (isSaveDetailPage && sources.BrandBulkTestDoxList.Any())
            {
                oldDetailData = oldDetailData.Where(o => o.ReportNo == sources.BrandBulkTestDoxList.FirstOrDefault().ReportNo).ToList();
            }

            List<BrandBulkTestDox> needUpdateDetailList =
                PublicClass.CompareListValue<BrandBulkTestDox>(
                    sources.BrandBulkTestDoxList,
                    oldDetailData,
                    "ReportNo,Ukey",
                    "FileName");


            string insertDetail = $@" ----寫入 BrandBulkTestDox
INSERT INTO BrandBulkTestDox
           (ReportNo
           ,FileName
           ,FilePath
           ,AddDate
           ,AddName)
VALUES 
           (@ReportNo
           ,@FileName
           ,@FilePath
           ,GETDATE()
           ,@AddName)
;
";
            string updateDetail = $@" ----更新 BrandBulkTestDox

UPDATE BrandBulkTestDox
SET EditDate = GETDATE() , EditName = @EditName
    ,FileName = @FileName
    ,FilePath = @FilePath
WHERE ReportNo = @ReportNo AND Ukey = @Ukey
";
            string deleteDetail = $@" ----刪除 BrandBulkTestDox

DELETE FROM BrandBulkTestDox where ReportNo = @ReportNo  AND Ukey = @Ukey
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:

                        listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                        listDetailPar.Add(new SqlParameter($"@FileName", detailItem.FileName));
                        listDetailPar.Add(new SqlParameter($"@FilePath", detailItem.FilePath));
                        listDetailPar.Add(new SqlParameter($"@AddName", UserID));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);

                        break;
                    //case CompareStateType.Edit: 檔案只有新增和刪除
                    //    listDetailPar.Add(new SqlParameter($"@ReportNo", sources.Main.ReportNo));
                    //    listDetailPar.Add(new SqlParameter($"@FileName", detailItem.FileName));
                    //    listDetailPar.Add(new SqlParameter($"@Ukey", detailItem.Ukey));
                    //    listDetailPar.Add(new SqlParameter($"@AddName", UserID));

                    //    ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                    //    break;
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

    }
}
