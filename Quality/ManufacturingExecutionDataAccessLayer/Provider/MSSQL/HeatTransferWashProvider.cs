using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
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
    public class HeatTransferWashProvider : SQLDAL
    {
        #region 底層連線
        public HeatTransferWashProvider(string ConString) : base(ConString) { }
        public HeatTransferWashProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<SelectListItem> GetReportNo_Source(HeatTransferWash_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select  [Text]= ReportNo, [Value]= ReportNo
from HeatTransferWash  WITH(NOLOCK)
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
            if (!string.IsNullOrEmpty(Req.Article))
            {
                SbSql += "AND  Article = @Article " + Environment.NewLine;
                objParameter.Add("Article", DbType.String, Req.Article);
            }

            SbSql += " ORDER BY ReportNo DESC";

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, objParameter).ToList();
        }

        public HeatTransferWash_Result GetMainData(HeatTransferWash_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select h.ReportNo
      ,h.OrderID
      ,h.BrandID
      ,h.SeasonID
      ,h.StyleID
      ,h.Article
      ,h.Line
      ,h.Machine
      ,IsTeamwear = CAST(h.IsTeamwear as bit)
      ,h.ReportDate
      ,h.Result
      ,h.Remark
      ,h.Status
      ,h.Temperature
      ,h.Time
      ,h.Pressure
      ,h.PeelOff
      ,h.Cycles
      ,h.TemperatureUnit
      ,h.AddName
      ,h.AddDate
      ,h.EditName
      ,h.EditDate
      ,pmsfile.TestBeforePicture
      ,pmsfile.TestAfterPicture
      ,MRHandleEmail = ISNULL(p.Email, p2.Email)
from HeatTransferWash h
inner join SciPMSFile_HeatTransferWash pmsfile on h.ReportNo=pmsfile.ReportNo
inner join SciProduction_Orders o on h.OrderID = o.ID
left join SciProduction_Pass1 p on o.MRHandle = p.ID
left join Pass1 p2 on o.MRHandle = p2.ID
where 1 = 1 
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                SbSql += "and h.ReportNo=@ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", DbType.String, Req.ReportNo);
            }
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "and h.BrandID=@BrandID" + Environment.NewLine;
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "and h.SeasonID=@SeasonID" + Environment.NewLine;
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "and h.StyleID=@StyleID" + Environment.NewLine;
                objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                SbSql += "and h.Article=@Article" + Environment.NewLine;
                objParameter.Add("@Article", DbType.String, Req.Article);
            }




            List<HeatTransferWash_Result> res = ExecuteList<HeatTransferWash_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new HeatTransferWash_Result();
        }
        public IList<HeatTransferWash_Detail_Result> GetDetailData(string ReportNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select    h.*
from HeatTransferWash_Detail h WITH(NOLOCK)
where h.ReportNo = @ReportNo
";

            objParameter.Add("ReportNo", DbType.String, ReportNo);
            IList<HeatTransferWash_Detail_Result> res = ExecuteList<HeatTransferWash_Detail_Result>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList() : new List<HeatTransferWash_Detail_Result>();
        }

        public int Insert_HeatTransferWash(HeatTransferWash_Result Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "HW", "HeatTransferWash", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@OrderID", DbType.String, Req.OrderID ?? ""} ,
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@Line", DbType.String,  Req.Line ?? "" } ,
                { "@Result", DbType.String,  Req.Result ?? "" } ,
                { "@Machine", DbType.String,  Req.Machine ?? "" } ,
                { "@IsTeamwear", DbType.Boolean, Req.IsTeamwear } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Temperature } ,
                { "@Time", DbType.Int32, Req.Time } ,
                { "@Pressure", DbType.Decimal, Req.Pressure } ,
                { "@PeelOff", DbType.String, Req.PeelOff ?? "" } ,
                { "@Cycles", DbType.Int32, Req.Cycles } ,
                { "@TemperatureUnit", DbType.Int32, Req.TemperatureUnit } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
            };

            if (Req.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.TestAfterPicture != null)
            {
                objParameter.Add("@TestAfterPicture", Req.TestAfterPicture);
            }
            else
            {
                objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"
SET XACT_ABORT ON

INSERT INTO HeatTransferWash
           (ReportNo
           ,OrderID
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,Line
           ,Machine
           ,IsTeamwear
           ,Remark
           ,Status
           ,Temperature
           ,Time
           ,Result
           ,Pressure
           ,PeelOff
           ,Cycles
           ,TemperatureUnit
           ,AddName
           ,AddDate)
VALUES  (     
            @ReportNo
           ,@OrderID
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,@Article
           ,@Line
           ,@Machine
           ,@IsTeamwear
           ,@Remark
           ,'New'
           ,@Temperature
           ,@Time
           ,@Result
           ,@Pressure
           ,@PeelOff
           ,@Cycles
           ,@TemperatureUnit
           ,@AddName
           ,GETDATE() )

INSERT INTO SciPMSFile_HeatTransferWash
           (ReportNo
            ,TestBeforePicture
            ,TestAfterPicture)
VALUES(    
            @ReportNo
            ,@TestBeforePicture
            ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Insert_HeatTransferWash_Detail(HeatTransferWash_Detail_Result Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.ReportNo } ,
                { "@FabricRefNo", DbType.String, Req.FabricRefNo ?? ""} ,
                { "@HTRefNo", DbType.String, Req.HTRefNo  ?? ""} ,
                { "@Result", DbType.String, Req.Result ?? "" } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@EditName", DbType.String, Req.EditName ?? "" } ,
            };

            SbSql.Append($@"
INSERT INTO dbo.HeatTransferWash_Detail
           (ReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate)
VALUES      (@ReportNo
           ,@FabricRefNo
           ,@HTRefNo
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE())
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update_HeatTransferWash(HeatTransferWash_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Line", DbType.String, Req.Main.Line ?? string.Empty);
            objParameter.Add("@Machine", DbType.String, Req.Main.Machine ?? string.Empty);
            objParameter.Add("@Result", DbType.String, Req.Main.Result ?? string.Empty);
            objParameter.Add("@Remark", DbType.String, Req.Main.Remark ?? string.Empty);
            objParameter.Add("@IsTeamwear", DbType.Boolean, Req.Main.IsTeamwear);
            objParameter.Add("@Temperature", DbType.Int32, Req.Main.Temperature);
            objParameter.Add("@Time", DbType.Int32, Req.Main.Time);
            objParameter.Add("@Pressure", DbType.Decimal, Req.Main.Pressure);
            objParameter.Add("@PeelOff", DbType.String, Req.Main.PeelOff ?? string.Empty);
            objParameter.Add("@Cycles", DbType.Int32, Req.Main.Cycles);
            objParameter.Add("@TemperatureUnit", DbType.Int32, Req.Main.TemperatureUnit);
            objParameter.Add("@Editname", DbType.String, Req.Main.EditName ?? string.Empty);
            objParameter.Add("@ReportNo", DbType.String, Req.Main.ReportNo ?? string.Empty);


            if (Req.Main.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            if (Req.Main.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Req.Main.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string head = $@"
SET XACT_ABORT ON

Update HeatTransferWash
Set Line = @Line
    ,Machine = @Machine
    ,Result = @Result
    ,Remark = @Remark
    ,IsTeamwear = @IsTeamwear
    ,Temperature = @Temperature
    ,Time = @Time
    ,Pressure = @Pressure
    ,PeelOff = @PeelOff
    ,Cycles = @Cycles 
    ,TemperatureUnit = @TemperatureUnit
    ,EditDate =GETDATE()
    ,Editname = @Editname
where ReportNo = @ReportNo


if not exists (select 1 from SciPMSFile_HeatTransferWash where ReportNo = @ReportNo)
begin
    INSERT INTO SciPMSFile_HeatTransferWash (ReportNo,TestBeforePicture,TestAfterPicture)
    VALUES (@ReportNo,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    UPDATE SciPMSFile_HeatTransferWash
    SET
        TestBeforePicture=@TestBeforePicture
        ,TestAfterPicture=@TestAfterPicture
    WHERE ReportNo = @ReportNo
end

";


            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Delete_HeatTransferWash(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo ?? string.Empty);

            string head = $@"
SET XACT_ABORT ON

DELETE FROM  HeatTransferWash
where ReportNo = @ReportNo

DELETE FROM  HeatTransferWash_Detail
where ReportNo = @ReportNo

DELETE FROM  SciPMSFile_HeatTransferWash
where ReportNo = @ReportNo
";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Confirm_HeatTransferWash(HeatTransferWash_ViewModel Req)
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
                objParameter.Add("@ReportDate", DbType.DateTime,DateTime.Now);
            }

            string head = $@"

Update HeatTransferWash
Set  ReportDate = @ReportDate
    ,Status = @Status
    ,EditDate = GETDATE()
    ,Editname = @Editname
where ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public void Update_HeatTransferWash_Detail(HeatTransferWash_ViewModel Req)
        {
            List<HeatTransferWash_Detail_Result> oldData = this.GetDetailData(Req.Main.ReportNo).ToList();

            List<HeatTransferWash_Detail_Result> needUpdateDetailList =
                PublicClass.CompareListValue<HeatTransferWash_Detail_Result>(
                    Req.Details,
                    oldData,
                    "Ukey",
                    "FabricRefNo,HTRefNo,Result,Remark");

            string insert = $@"
INSERT INTO dbo.HeatTransferWash_Detail
           (ReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate)
     VALUES(
            @ReportNo
           ,@FabricRefNo
           ,@HTRefNo
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE())

";
            string update = $@"
UPDATE dbo.HeatTransferWash_Detail
SET  FabricRefNo=@FabricRefNo
    ,HTRefNo = @HTRefNo
    ,Result = @Result
    ,Remark = @Remark
    ,EditName = @EditName
    ,EditDate = GETDATE()
WHERE UKey = @Ukey

";
            string delete = @"
delete HeatTransferWash_Detail where Ukey = @Ukey
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                listDetailPar.Add("@FabricRefNo", DbType.String, detailItem.FabricRefNo ?? string.Empty);
                listDetailPar.Add("@HTRefNo", DbType.String, detailItem.HTRefNo ?? string.Empty);
                listDetailPar.Add("@Result", DbType.String, detailItem.Result);
                listDetailPar.Add("@Remark", DbType.String, detailItem.Remark ?? string.Empty);
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
    }
}
