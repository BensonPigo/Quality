using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using ToolKit;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class DailyMoistureProvider : SQLDAL
    {
        #region 底層連線
        public DailyMoistureProvider(string ConString) : base(ConString) { }
        public DailyMoistureProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<SelectListItem> GetReportNo_Source(DailyMoisture_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select  [Text]= ReportNo, [Value]= ReportNo
from BulkMoistureTest  WITH(NOLOCK)
WHERE 1=1
"; ;
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "AND  BrandID = @BrandID " + Environment.NewLine;
                objParameter.Add("BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "AND  SeasonID = @SeasonID " + Environment.NewLine;
                objParameter.Add("SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "AND  StyleID  = @StyleID " + Environment.NewLine;
                objParameter.Add("StyleID", DbType.String, Req.StyleID);
            }
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql += "AND  OrderID = @OrderID " + Environment.NewLine;
                objParameter.Add("OrderID", DbType.String, Req.OrderID);
            }

            SbSql += " ORDER BY ReportNo DESC";

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, objParameter).ToList();
        }

        public DailyMoisture_Result GetMainData(DailyMoisture_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select b.ReportNo
      ,b.OrderID
      ,b.BrandID
      ,b.SeasonID
      ,b.StyleID
      ,b.Article
      ,b.ReportDate

      ,b.Instrument
      ,b.Fabrication
	  ,Standard = CAST( ISNULL(e.Standard,0) as decimal(6,2))
      ,b.Action
      ,b.Result
      ,b.Remark
      ,b.Status
      ,b.AddName
      ,b.AddDate
      ,b.EditName
      ,b.EditDate
      ,b.Line
      ,MRHandleEmail = ISNULL(p.Email, p2.Email)
from BulkMoistureTest b
left join EndlineMoisture e on  e.Instrument=b.Instrument and b.Fabrication=e.Fabrication 
inner join SciProduction_Orders o on b.OrderID = o.ID
left join SciProduction_Pass1 p on o.MRHandle = p.ID
left join Pass1 p2 on o.MRHandle = p2.ID
where 1 = 1 
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                SbSql += "and b.ReportNo=@ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", DbType.String, Req.ReportNo);
            }
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql += "and b.OrderID=@OrderID" + Environment.NewLine;
                objParameter.Add("@OrderID", DbType.String, Req.OrderID);
            }
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "and b.BrandID=@BrandID" + Environment.NewLine;
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "and b.SeasonID=@SeasonID" + Environment.NewLine;
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "and b.StyleID=@StyleID" + Environment.NewLine;
                objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            }

            List<DailyMoisture_Result> res = ExecuteList<DailyMoisture_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new DailyMoisture_Result();
        }

        public IList<DailyMoisture_Detail_Result> GetDetailData(string ReportNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select b.*
from BulkMoistureTest_Detail b WITH(NOLOCK)
where b.ReportNo = @ReportNo
";

            objParameter.Add("ReportNo", DbType.String, ReportNo);
            IList<DailyMoisture_Detail_Result> res = ExecuteList<DailyMoisture_Detail_Result>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList() : new List<DailyMoisture_Detail_Result>();
        }

        public int Insert_DailyMoisture(DailyMoisture_Result Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "BM", "BulkMoistureTest", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@OrderID", DbType.String, Req.OrderID ?? ""} ,
                { "@BrandID", DbType.String, Req.BrandID ?? ""} ,
                { "@SeasonID", DbType.String, Req.SeasonID ?? ""} ,
                { "@StyleID", DbType.String, Req.StyleID ?? ""} ,
                { "@Instrument", DbType.String, Req.Instrument ?? ""} ,
                { "@Fabrication", DbType.String, Req.Fabrication ?? "" } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@Action", DbType.String, Req.Action ?? "" } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
                { "@Line", DbType.String, Req.Line ?? "" } ,
            };

            SbSql.Append($@"
INSERT INTO dbo.BulkMoistureTest
           (ReportNo
           ,OrderID
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Status
           ,Instrument
           ,Fabrication
           ,Remark
           ,Action
           ,Line
           ,AddName
           ,AddDate)
VALUES
           (@ReportNo
           ,@OrderID
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,'New'
           ,@Instrument
           ,@Fabrication
           ,@Remark
           ,@Action
           ,@Line
           ,@AddName
           ,GETDATE()
)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Insert_DailyMoisture_Detail(DailyMoisture_Detail_Result Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.ReportNo } ,
                { "@Point1",  Req.Point1 } ,
                { "@Point2",  Req.Point2 } ,
                { "@Point3",  Req.Point3 } ,
                { "@Point4",  Req.Point4 } ,
                { "@Point5",  Req.Point5 } ,
                { "@Area", DbType.String, Req.Area ?? "" } ,
                { "@Fabric", DbType.String, Req.Fabric ?? "" } ,
                { "@Result", DbType.String, Req.Result ?? "" } ,
                { "@EditName", DbType.String, Req.EditName ?? "" } ,
            };

            SbSql.Append($@"
INSERT INTO dbo.BulkMoistureTest_Detail
           (ReportNo
           ,Point1
           ,Point2
           ,Point3
           ,Point4
           ,Point5
           ,Area
           ,Fabric
           ,Result
           ,EditName
           ,EditDate)
VALUES      (@ReportNo
           ,@Point1
           ,@Point2
           ,@Point3
           ,@Point4
           ,@Point5
           ,@Area
           ,@Fabric
           ,@Result
           ,@EditName
           ,GETDATE())
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update_DailyMoisture(DailyMoisture_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, Req.Main.ReportNo ?? string.Empty);
            objParameter.Add("@Instrument", DbType.String, Req.Main.Instrument ?? string.Empty);
            objParameter.Add("@Fabrication", DbType.String, Req.Main.Fabrication ?? string.Empty);
            objParameter.Add("@Standard", Req.Main.Standard);
            objParameter.Add("@Action", DbType.String, Req.Main.Action ?? string.Empty);
            objParameter.Add("@Remark", DbType.String, Req.Main.Remark ?? string.Empty);
            objParameter.Add("@Line", DbType.String, Req.Main.Line ?? string.Empty);
            objParameter.Add("@Editname", DbType.String, Req.Main.EditName ?? string.Empty);


            string head = $@"
Update BulkMoistureTest
Set Instrument = @Instrument
    ,Fabrication = @Fabrication
    ,Action = @Action
    ,Line = @Line
    ,Remark = @Remark
    ,EditDate =GETDATE()
    ,Editname = @Editname
where ReportNo = @ReportNo
";


            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Delete_DailyMoisture(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo ?? string.Empty);

            string head = $@"
DELETE FROM  BulkMoistureTest
where ReportNo = @ReportNo

DELETE FROM  BulkMoistureTest_Detail
where ReportNo = @ReportNo
";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Confirm_DailyMoisture(DailyMoisture_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Editname", DbType.String, Req.Main.EditName ?? string.Empty);
            objParameter.Add("@ReportNo", DbType.String, Req.Main.ReportNo ?? string.Empty);
            objParameter.Add("@Status", DbType.String, Req.Main.Status ?? string.Empty);

            if (Req.Main.Status.ToUpper() == "NEW")
            {
                objParameter.Add("@ReportDate", DbType.DateTime, DBNull.Value);
            }
            if (Req.Main.Status.ToLower() == "confirmed")
            {
                objParameter.Add("@ReportDate", DbType.DateTime, DateTime.Now);
            }

            string head = $@"

Update b
Set  ReportDate = @ReportDate
    ,Status = @Status
    ,Result = IIF(@Status ='New' 
                        ,''
                        , IIF( (select COUNT(1) from BulkMoistureTest_Detail a where a.ReportNo=b.ReportNo and a.Result='Fail') > 0 , 'Fail' ,'Pass')
                    )
    ,EditDate = GETDATE()
    ,Editname = @Editname
from BulkMoistureTest b
where ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public void Update_DailyMoisture_Detail(DailyMoisture_ViewModel Req)
        {
            List<DailyMoisture_Detail_Result> oldData = this.GetDetailData(Req.Main.ReportNo).ToList();

            List<DailyMoisture_Detail_Result> needUpdateDetailList =
                PublicClass.CompareListValue<DailyMoisture_Detail_Result>(
                    Req.Details,
                    oldData,
                    "Ukey",
                    "Point1,Point2,Point3,Point4,Point5,Area,Fabric,Result");

            string insert = $@"
INSERT INTO dbo.BulkMoistureTest_Detail
           (ReportNo
           ,Point1
           ,Point2
           ,Point3
           ,Point4
           ,Point5
           ,Area
           ,Fabric
           ,Result
           ,EditName
           ,EditDate)
     VALUES(
            @ReportNo
           ,@Point1
           ,@Point2
           ,@Point3
           ,@Point4
           ,@Point5
           ,@Area
           ,@Fabric
           ,@Result
           ,@EditName
           ,GETDATE())

";
            string update = $@"
UPDATE dbo.BulkMoistureTest_Detail
SET  Point1 = @Point1
    ,Point2 = @Point2
    ,Point3 = @Point3
    ,Point4 = @Point4
    ,Point5 = @Point5
    ,Area = @Area
    ,Fabric = @Fabric
    ,Result = @Result
    ,EditName = @EditName
    ,EditDate = GETDATE()
WHERE UKey = @Ukey

";
            string delete = @"
delete BulkMoistureTest_Detail where Ukey = @Ukey
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                listDetailPar.Add("@Point1", detailItem.Point1);
                listDetailPar.Add("@Point2", detailItem.Point2);
                listDetailPar.Add("@Point3", detailItem.Point3);
                listDetailPar.Add("@Point4", detailItem.Point4);
                listDetailPar.Add("@Point5", detailItem.Point5);
                listDetailPar.Add("@Area", DbType.String, detailItem.Area ?? string.Empty);
                listDetailPar.Add("@Fabric", DbType.String, detailItem.Fabric ?? string.Empty);
                listDetailPar.Add("@Result", DbType.String, detailItem.Result ?? string.Empty);
                listDetailPar.Add("@EditName", DbType.String, detailItem.EditName ?? string.Empty);

                switch (detailItem.StateType)
                {
                    case DatabaseObject.Public.CompareStateType.Add:
                        listDetailPar.Add("@ReportNo", DbType.String, detailItem.ReportNo);
                        ExecuteNonQuery(CommandType.Text, insert, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add("@Ukey", DbType.Int64, detailItem.Ukey);
                        ExecuteNonQuery(CommandType.Text, update, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Delete:
                        listDetailPar.Add("@Ukey", detailItem.Ukey);
                        ExecuteNonQuery(CommandType.Text, delete, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.None:
                        break;
                    default:
                        break;
                }
            }
        }

        public IList<EndlineMoisture> GetEndlineMoisture()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlGetEndlineMoisture = @"
select  Instrument
        ,Fabrication
        ,Standard
        ,Unit
        ,Junk
        ,Description
        ,AddDate
        ,AddName
        ,EditDate
        ,Editname
from    EndlineMoisture with (nolock)

";
            return ExecuteList<EndlineMoisture>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        public IList<SelectListItem> GetAction()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlGetEndlineMoisture = @"
select  [Text] = '', [Value] = ''
union
select  [Text] = Name, [Value] = Name 
from Production.dbo.DropDownList ddl WITH(NOLOCK) 
where type='PMS_MoistureAction'

";
            return ExecuteList<SelectListItem>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

    }
}
